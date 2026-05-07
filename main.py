import sys
import os
import json
import time
import socket
import shutil
import asyncio
import multiprocessing
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from fastapi.responses import StreamingResponse, JSONResponse
from pydantic import BaseModel
from openai import OpenAI


# ---------------------------------------------------------------------------
# Path resolution — works both in dev and PyInstaller frozen mode
# ---------------------------------------------------------------------------

def get_base_dir() -> str:
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def get_frontend_dir() -> str:
    candidate = os.path.join(get_base_dir(), 'frontend')
    if os.path.isdir(candidate):
        return candidate
    meipass = getattr(sys, '_MEIPASS', None)
    if meipass:
        return os.path.join(meipass, 'frontend')
    return os.path.join(os.path.dirname(__file__), 'frontend')


def get_addon_dir() -> str:
    candidate = os.path.join(get_base_dir(), 'wpsaddon')
    if os.path.isdir(candidate):
        return candidate
    meipass = getattr(sys, '_MEIPASS', None)
    if meipass:
        return os.path.join(meipass, 'wpsaddon')
    return os.path.join(os.path.dirname(__file__), 'wpsaddon')


BASE_DIR = get_base_dir()
CONFIG_PATH = os.path.join(BASE_DIR, 'config.json')
LOCK_PATH = os.path.join(BASE_DIR, 'wps_agent.lock')
HISTORIES_DIR = os.path.join(BASE_DIR, 'histories')
SNAPSHOTS_DIR = os.path.join(BASE_DIR, 'snapshots')
CURRENT_PORT = 3889  # updated at startup

# ---------------------------------------------------------------------------
# System prompt
# ---------------------------------------------------------------------------

SYSTEM_PROMPT = """你是一个智能助手，集成在 WPS Office 任务窗格中，既能帮用户操作文档，也能回答各种问题、进行日常对话。

## 响应格式

**需要操作文档时**，返回以下 JSON（不加任何 Markdown 包裹）：
{
  "message": "简短说明你做了什么",
  "jsa_code": "// JSA 代码"
}

**不需要操作文档时**（聊天、问答、分析、建议等），同样返回 JSON，jsa_code 设为 null：
{
  "message": "你的完整回答",
  "jsa_code": null
}

## WPS JSAPI 参考（任务窗格中使用 window.Application，不是 wps）

- 入口：var app = Application;
- 文字文档：app.ActiveDocument
- 选区：app.ActiveDocument.ActiveWindow.Selection
- 插入文字：Selection.TypeText("文字")
- 插入段落：Selection.TypeParagraph()
- 加粗：Selection.Font.Bold = true
- 保存：app.ActiveDocument.Save()
- 表格单元格：Application.ActiveWorkbook.ActiveSheet.Cells(行, 列).Value = 值
- 弹框：alert("消息")
- 注意：不要使用 wps.WpsApplication() 等旧式入口

## 注意事项

- 生成的 JSA 代码在用户点击「执行」后才会运行，用户自己决定是否执行
- 如果用户提问与文档无关，直接在 message 里正常回答即可
- 默认用中文回答，用户用其他语言则跟随切换"""

# ---------------------------------------------------------------------------
# Config helpers
# ---------------------------------------------------------------------------

DEFAULT_CONFIG = {
    "api_key": "",
    "model": "deepseek-chat",
    "base_url": "https://api.deepseek.com/v1",
    "temperature": 0.3,
    "max_tokens": 2048,
}


def load_config() -> dict:
    if os.path.exists(CONFIG_PATH):
        try:
            with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            return {**DEFAULT_CONFIG, **cfg}
        except Exception:
            pass
    return dict(DEFAULT_CONFIG)


def save_config(data: dict):
    cfg = load_config()
    cfg.update(data)
    with open(CONFIG_PATH, 'w', encoding='utf-8') as f:
        json.dump(cfg, f, ensure_ascii=False, indent=2)


def mask_key(key: str) -> str:
    if not key or len(key) < 4:
        return '****'
    return '*' * (len(key) - 4) + key[-4:]


# ---------------------------------------------------------------------------
# Port + single-instance logic
# ---------------------------------------------------------------------------

def find_free_port(start: int = 3889, end: int = 3950) -> int:
    for port in range(start, end):
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
            try:
                s.bind(('127.0.0.1', port))
                return port
            except OSError:
                continue
    raise RuntimeError("No free port found in range 3889-3950")


def read_lock_port() -> int | None:
    if not os.path.exists(LOCK_PATH):
        return None
    try:
        with open(LOCK_PATH, 'r') as f:
            return int(f.read().strip())
    except Exception:
        return None


def write_lock(port: int):
    with open(LOCK_PATH, 'w') as f:
        f.write(str(port))


def delete_lock():
    try:
        os.remove(LOCK_PATH)
    except Exception:
        pass


def server_is_alive(port: int) -> bool:
    import urllib.request
    try:
        urllib.request.urlopen(f'http://127.0.0.1:{port}/api/health', timeout=2)
        return True
    except Exception:
        return False


def register_wps_addin(port: int):
    """Ensure publish.xml has entries for wps/et/wpp types pointing to the current port."""
    import re
    jsaddons_dir = os.path.join(os.path.expandvars('%APPDATA%'),
                                'kingsoft', 'wps', 'jsaddons')
    publish_path = os.path.join(jsaddons_dir, 'publish.xml')
    url = f'http://127.0.0.1:{port}/'
    app_types = ['wps', 'et', 'wpp', 'pdf']

    entries = '\n  '.join(
        f'<jspluginonline name="wps-ai-agent" type="{t}" url="{url}" debug="" enable="enable_dev" install="null"/>'
        for t in app_types
    )
    fresh = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        '<jsplugins>\n'
        f'  {entries}\n'
        '</jsplugins>\n'
    )

    if os.path.exists(publish_path):
        try:
            with open(publish_path, 'r', encoding='utf-8') as f:
                content = f.read()
            # Check all types already point to the right port
            if all(f'type="{t}"' in content and url in content for t in app_types):
                print(f"WPSJS addon already registered for wps/et/wpp at port {port}")
                return
            # Remove old wps-ai-agent entries, preserve other plugins
            content = re.sub(r'\s*<jspluginonline[^>]*name="wps-ai-agent"[^/]*/>', '', content)
            content = content.replace('</jsplugins>', f'  {entries}\n</jsplugins>')
            with open(publish_path, 'w', encoding='utf-8') as f:
                f.write(content)
            print(f"publish.xml updated: wps-ai-agent registered for {app_types} at port {port}")
            return
        except Exception as e:
            print(f"publish.xml update failed: {e}")
            return

    # File doesn't exist — create it
    try:
        os.makedirs(jsaddons_dir, exist_ok=True)
        with open(publish_path, 'w', encoding='utf-8') as f:
            f.write(fresh)
        print(f"publish.xml created for {app_types} at port {port}")
    except Exception as e:
        print(f"publish.xml creation failed: {e}")


# ---------------------------------------------------------------------------
# FastAPI app
# ---------------------------------------------------------------------------

app = FastAPI(title="WPS AI Agent")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*", "null"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


# ---------------------------------------------------------------------------
# Pydantic models
# ---------------------------------------------------------------------------

class ChatRequest(BaseModel):
    messages: list
    stream: bool = True


class ConfigRequest(BaseModel):
    api_key: str | None = None
    model: str | None = None
    base_url: str | None = None
    temperature: float | None = None
    max_tokens: int | None = None


class SessionSaveRequest(BaseModel):
    messages: list
    title: str | None = None


class PDFExtractRequest(BaseModel):
    file_path: str
    max_pages: int = 50
    max_chars: int = 60000


class SnapshotRequest(BaseModel):
    source_path: str


class RestoreRequest(BaseModel):
    snapshot_path: str
    original_path: str


# ---------------------------------------------------------------------------
# Routes — debug log (让 WPS 后台 JS 回写状态，替代 DevTools)
# ---------------------------------------------------------------------------

DEBUG_LOG_PATH = os.path.join(BASE_DIR, 'wps_debug.log')


@app.post("/api/debug")
async def post_debug(request: Request):
    try:
        body = await request.json()
    except Exception:
        body = {}
    line = f"[{datetime.now().strftime('%H:%M:%S')}] {json.dumps(body, ensure_ascii=False)}\n"
    with open(DEBUG_LOG_PATH, 'a', encoding='utf-8') as f:
        f.write(line)
    return {"ok": True}


@app.get("/api/debug")
async def get_debug():
    if not os.path.exists(DEBUG_LOG_PATH):
        return {"log": "（暂无日志）"}
    with open(DEBUG_LOG_PATH, 'r', encoding='utf-8') as f:
        return {"log": f.read()}


# ---------------------------------------------------------------------------
# Routes — health & config
# ---------------------------------------------------------------------------

@app.get("/api/open-ui")
async def open_ui():
    import webbrowser
    webbrowser.open(f"http://127.0.0.1:{CURRENT_PORT}/ui/index.html?v=7")
    return {"ok": True}


@app.get("/api/health")
async def health():
    cfg = load_config()
    return {"status": "ok", "config_loaded": bool(cfg.get("api_key"))}


@app.get("/api/config")
async def get_config():
    cfg = load_config()
    return {
        "api_key_masked": mask_key(cfg.get("api_key", "")),
        "model": cfg.get("model", ""),
        "base_url": cfg.get("base_url", ""),
        "temperature": cfg.get("temperature", 0.3),
        "max_tokens": cfg.get("max_tokens", 2048),
    }


@app.post("/api/config")
async def post_config(req: ConfigRequest):
    data = {k: v for k, v in req.model_dump().items() if v is not None}
    save_config(data)
    return {"ok": True}


# ---------------------------------------------------------------------------
# Routes — chat (SSE streaming)
# ---------------------------------------------------------------------------

@app.post("/api/chat")
async def chat(req: ChatRequest):
    cfg = load_config()
    if not cfg.get("api_key"):
        raise HTTPException(status_code=400, detail="API Key 未配置，请先在设置中填写")

    client = OpenAI(api_key=cfg["api_key"], base_url=cfg.get("base_url"))
    full_messages = [{"role": "system", "content": SYSTEM_PROMPT}] + req.messages

    async def generate():
        try:
            stream = client.chat.completions.create(
                model=cfg["model"],
                messages=full_messages,
                temperature=cfg.get("temperature", 0.3),
                max_tokens=cfg.get("max_tokens", 2048),
                stream=True,
            )
            for chunk in stream:
                content = chunk.choices[0].delta.content
                if content:
                    data = json.dumps({"choices": [{"delta": {"content": content}}]})
                    yield f"data: {data}\n\n"
                await asyncio.sleep(0)
            yield "data: [DONE]\n\n"
        except Exception as e:
            err = json.dumps({"error": str(e)})
            yield f"data: {err}\n\n"
            yield "data: [DONE]\n\n"

    return StreamingResponse(generate(), media_type="text/event-stream")


# ---------------------------------------------------------------------------
# Routes — sessions (conversation history)
# ---------------------------------------------------------------------------

@app.get("/api/sessions")
async def list_sessions():
    os.makedirs(HISTORIES_DIR, exist_ok=True)
    sessions = []
    for f in sorted(Path(HISTORIES_DIR).glob("*.json")):
        if f.name == "_draft.json":
            continue
        try:
            data = json.loads(f.read_text(encoding='utf-8'))
            sessions.append({
                "id": data.get("id", f.stem),
                "title": data.get("title", "无标题"),
                "created_at": data.get("created_at", ""),
                "message_count": len(data.get("messages", [])),
            })
        except Exception:
            pass
    return list(reversed(sessions))


@app.get("/api/sessions/{session_id}")
async def get_session(session_id: str):
    path = os.path.join(HISTORIES_DIR, f"{session_id}.json")
    if not os.path.exists(path):
        raise HTTPException(status_code=404, detail="会话不存在")
    return json.loads(Path(path).read_text(encoding='utf-8'))


@app.post("/api/sessions")
async def save_session(req: SessionSaveRequest):
    os.makedirs(HISTORIES_DIR, exist_ok=True)
    session_id = datetime.now().strftime("%Y-%m-%dT%H-%M-%S")
    title = req.title
    if not title and req.messages:
        first_user = next((m["content"] for m in req.messages if m["role"] == "user"), "")
        title = first_user[:20] + ("..." if len(first_user) > 20 else "")
    data = {
        "id": session_id,
        "title": title or "无标题",
        "created_at": datetime.now().isoformat(),
        "messages": req.messages,
    }
    path = os.path.join(HISTORIES_DIR, f"{session_id}.json")
    Path(path).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding='utf-8')
    return {"ok": True, "id": session_id}


@app.delete("/api/sessions/{session_id}")
async def delete_session(session_id: str):
    path = os.path.join(HISTORIES_DIR, f"{session_id}.json")
    if os.path.exists(path):
        os.remove(path)
    return {"ok": True}


@app.post("/api/sessions/draft")
async def save_draft(req: SessionSaveRequest):
    os.makedirs(HISTORIES_DIR, exist_ok=True)
    path = os.path.join(HISTORIES_DIR, "_draft.json")
    data = {"messages": req.messages}
    Path(path).write_text(json.dumps(data, ensure_ascii=False), encoding='utf-8')
    return {"ok": True}


@app.get("/api/sessions/draft/load")
async def load_draft():
    path = os.path.join(HISTORIES_DIR, "_draft.json")
    if not os.path.exists(path):
        return {"messages": []}
    return json.loads(Path(path).read_text(encoding='utf-8'))


# ---------------------------------------------------------------------------
# Routes — PDF extraction
# ---------------------------------------------------------------------------

@app.post("/api/pdf/extract")
async def extract_pdf(req: PDFExtractRequest):
    path = req.file_path
    if not os.path.exists(path):
        return JSONResponse(status_code=404, content={"ok": False, "error": "文件不存在: " + path})
    if not path.lower().endswith('.pdf'):
        return JSONResponse(status_code=400, content={"ok": False, "error": "只支持 PDF 文件"})
    try:
        import fitz  # PyMuPDF
        doc = fitz.open(path)
        page_count = len(doc)
        max_p = req.max_pages if req.max_pages > 0 else page_count
        extracted_pages = min(max_p, page_count)

        parts = []
        for i in range(extracted_pages):
            text = doc[i].get_text().strip()
            if text:
                parts.append(f"--- 第 {i + 1} 页 ---\n{text}")
        doc.close()

        full_text = '\n\n'.join(parts)
        char_count = len(full_text)
        truncated = char_count > req.max_chars
        if truncated:
            full_text = full_text[:req.max_chars]

        return {
            "ok": True,
            "text": full_text,
            "page_count": page_count,
            "extracted_pages": extracted_pages,
            "char_count": char_count,
            "truncated": truncated,
        }
    except Exception as e:
        return JSONResponse(status_code=500, content={"ok": False, "error": str(e)})


# ---------------------------------------------------------------------------
# Routes — snapshot (OS-level file copy, no WPS side effects)
# ---------------------------------------------------------------------------

@app.post("/api/snapshot")
async def create_snapshot(req: SnapshotRequest):
    os.makedirs(SNAPSHOTS_DIR, exist_ok=True)
    dst = os.path.join(SNAPSHOTS_DIR, f"snap_{int(time.time())}.docx")
    try:
        shutil.copy2(req.source_path, dst)
        return {"ok": True, "snapshot_path": dst}
    except PermissionError:
        return JSONResponse(status_code=423, content={
            "ok": False,
            "error": "文档当前被系统锁定，无法创建防灾快照，请重试或手动备份"
        })
    except Exception as e:
        return JSONResponse(status_code=500, content={"ok": False, "error": str(e)})


@app.post("/api/snapshot/restore")
async def restore_snapshot(req: RestoreRequest):
    real_snap = os.path.realpath(req.snapshot_path)
    real_snap_dir = os.path.realpath(SNAPSHOTS_DIR)
    if not real_snap.startswith(real_snap_dir + os.sep):
        return JSONResponse(status_code=403, content={"ok": False, "error": "非法的快照路径"})
    try:
        shutil.copy2(real_snap, req.original_path)
        return {"ok": True}
    except PermissionError:
        return JSONResponse(status_code=423, content={"ok": False, "error": "目标文件被锁定，无法恢复"})
    except Exception as e:
        return JSONResponse(status_code=500, content={"ok": False, "error": str(e)})


# ---------------------------------------------------------------------------
# Static file serving — mount last so API routes take priority
# ---------------------------------------------------------------------------

FRONTEND_DIR = get_frontend_dir()
if os.path.isdir(FRONTEND_DIR):
    app.mount("/ui", StaticFiles(directory=FRONTEND_DIR, html=True), name="ui")

ADDON_DIR = get_addon_dir()
if os.path.isdir(ADDON_DIR):
    app.mount("/", StaticFiles(directory=ADDON_DIR, html=True), name="addon")


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    multiprocessing.freeze_support()

    # Single-instance guard
    existing_port = read_lock_port()
    if existing_port and server_is_alive(existing_port):
        sys.exit(0)  # Another instance is already serving

    port = find_free_port()
    write_lock(port)
    register_wps_addin(port)

    import uvicorn
    import atexit
    CURRENT_PORT = port
    atexit.register(delete_lock)

    uvicorn.run(app, host="127.0.0.1", port=port, log_level="warning")
