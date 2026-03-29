import asyncio
import json
import os
import re
from pathlib import Path
from typing import Any

from fastapi import APIRouter, Depends, File, HTTPException, UploadFile, WebSocket, WebSocketDisconnect
from fastapi.responses import StreamingResponse
from sqlalchemy.orm import Session

from agent import chat_with_agent, chat_with_agent_stream, storage
from auth import authenticate_user, create_access_token, get_current_user, get_db, get_password_hash, require_admin, resolve_role
from document_loader import DocumentLoader
from embedding import EmbeddingService
from milvus_client import MilvusManager
from milvus_writer import MilvusWriter
from models import User
from parent_chunk_store import ParentChunkStore
from task import task_manager
from schemas import (
    AuthResponse,
    ChatRequest,
    ChatResponse,
    CurrentUserResponse,
    DocumentBatchDeleteRequest,
    DocumentBatchDeleteResponse,
    DocumentDeleteResponse,
    DocumentInfo,
    DocumentListResponse,
    DocumentUploadResponse,
    LoginRequest,
    MessageInfo,
    RegisterRequest,
    SessionDeleteResponse,
    SessionInfo,
    SessionListResponse,
    SessionMessagesResponse,
    UploadTaskCreateResponse,
    UploadTaskStatus,
)

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR.parent / "data"
UPLOAD_DIR = DATA_DIR / "documents"
MAX_FILE_SIZE = 50 * 1024 * 1024

loader = DocumentLoader()
parent_chunk_store = ParentChunkStore()
milvus_manager = MilvusManager()
embedding_service = EmbeddingService()
milvus_writer = MilvusWriter(embedding_service=embedding_service, milvus_manager=milvus_manager)

router = APIRouter()


class WebSocketManager:
    def __init__(self):
        self.active_connections: dict[str, list[WebSocket]] = {}

    async def connect(self, websocket: WebSocket, task_id: str):
        await websocket.accept()
        if task_id not in self.active_connections:
            self.active_connections[task_id] = []
        self.active_connections[task_id].append(websocket)

        def send_update(task_data: dict[str, Any]):
            try:
                asyncio.create_task(websocket.send_json(task_data))
            except Exception:
                pass

        task_manager.register_callback(task_id, send_update)

    def disconnect(self, websocket: WebSocket, task_id: str):
        if task_id in self.active_connections:
            if websocket in self.active_connections[task_id]:
                self.active_connections[task_id].remove(websocket)
            if not self.active_connections[task_id]:
                del self.active_connections[task_id]
        task_manager.unregister_callback(task_id)


ws_manager = WebSocketManager()


@router.post("/auth/register", response_model=AuthResponse)
async def register(request: RegisterRequest, db: Session = Depends(get_db)):
    username = (request.username or "").strip()
    password = (request.password or "").strip()
    if not username or not password:
        raise HTTPException(status_code=400, detail="用户名和密码不能为空")

    exists = db.query(User).filter(User.username == username).first()
    if exists:
        raise HTTPException(status_code=409, detail="用户名已存在")

    role = resolve_role(request.role, request.admin_code)
    user = User(username=username, password_hash=get_password_hash(password), role=role)
    db.add(user)
    db.commit()

    token = create_access_token(username=username, role=role)
    return AuthResponse(access_token=token, username=username, role=role)


@router.post("/auth/login", response_model=AuthResponse)
async def login(request: LoginRequest, db: Session = Depends(get_db)):
    user = authenticate_user(db, request.username, request.password)
    if not user:
        raise HTTPException(status_code=401, detail="用户名或密码错误")
    token = create_access_token(username=user.username, role=user.role)
    return AuthResponse(access_token=token, username=user.username, role=user.role)


@router.get("/auth/me", response_model=CurrentUserResponse)
async def me(current_user: User = Depends(get_current_user)):
    return CurrentUserResponse(username=current_user.username, role=current_user.role)


@router.get("/sessions/{session_id}", response_model=SessionMessagesResponse)
async def get_session_messages(session_id: str, current_user: User = Depends(get_current_user)):
    """获取指定会话的所有消息"""
    try:
        messages = [
            MessageInfo(
                type=msg["type"],
                content=msg["content"],
                timestamp=msg["timestamp"],
                rag_trace=msg.get("rag_trace"),
            )
            for msg in storage.get_session_messages(current_user.username, session_id)
        ]
        return SessionMessagesResponse(messages=messages)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.get("/sessions", response_model=SessionListResponse)
async def list_sessions(current_user: User = Depends(get_current_user)):
    """获取当前用户的所有会话列表"""
    try:
        sessions = [SessionInfo(**item) for item in storage.list_session_infos(current_user.username)]
        sessions.sort(key=lambda x: x.updated_at, reverse=True)
        return SessionListResponse(sessions=sessions)
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.delete("/sessions/{session_id}", response_model=SessionDeleteResponse)
async def delete_session(session_id: str, current_user: User = Depends(get_current_user)):
    """删除当前用户的指定会话"""
    try:
        deleted = storage.delete_session(current_user.username, session_id)
        if not deleted:
            raise HTTPException(status_code=404, detail="会话不存在")
        return SessionDeleteResponse(session_id=session_id, message="成功删除会话")
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))


@router.post("/chat", response_model=ChatResponse)
async def chat_endpoint(request: ChatRequest, current_user: User = Depends(get_current_user)):
    try:
        session_id = request.session_id or "default_session"
        resp = chat_with_agent(request.message, current_user.username, session_id)
        if isinstance(resp, dict):
            return ChatResponse(**resp)
        return ChatResponse(response=resp)
    except Exception as e:
        message = str(e)
        match = re.search(r"Error code:\s*(\d{3})", message)
        if match:
            code = int(match.group(1))
            if code == 429:
                raise HTTPException(
                    status_code=429,
                    detail=(
                        "上游模型服务触发限流/额度限制（429）。请检查账号额度/模型状态。\n"
                        f"原始错误：{message}"
                    ),
                )
            if code in (401, 403):
                raise HTTPException(status_code=code, detail=message)
            raise HTTPException(status_code=code, detail=message)
        raise HTTPException(status_code=500, detail=message)


@router.post("/chat/stream")
async def chat_stream_endpoint(request: ChatRequest, current_user: User = Depends(get_current_user)):
    """跟 Agent 对话 (流式)"""

    async def event_generator():
        try:
            session_id = request.session_id or "default_session"
            async for chunk in chat_with_agent_stream(request.message, current_user.username, session_id):
                yield chunk
        except Exception as e:
            error_data = {"type": "error", "content": str(e)}
            yield f"data: {json.dumps(error_data)}\n\n"

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={
            "Cache-Control": "no-cache, no-store, must-revalidate",
            "Connection": "keep-alive",
            "X-Accel-Buffering": "no",
        },
    )


@router.get("/documents", response_model=DocumentListResponse)
async def list_documents(_: User = Depends(require_admin)):
    """获取已上传的文档列表（管理员）"""
    try:
        milvus_manager.init_collection()

        results = milvus_manager.query(
            output_fields=["filename", "file_type"],
            limit=10000,
        )

        file_stats = {}
        for item in results:
            filename = item.get("filename", "")
            file_type = item.get("file_type", "")
            if filename not in file_stats:
                file_stats[filename] = {
                    "filename": filename,
                    "file_type": file_type,
                    "chunk_count": 0,
                }
            file_stats[filename]["chunk_count"] += 1

        documents = [DocumentInfo(**stats) for stats in file_stats.values()]
        return DocumentListResponse(documents=documents)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"获取文档列表失败: {str(e)}")


def _process_document_upload(file_path: str, filename: str, task_id: str):
    """实际的文件上传处理逻辑（在后台线程中执行）"""
    task_manager.update_progress(task_id, 10, "初始化向量数据库...")

    milvus_manager.init_collection()

    task_manager.update_progress(task_id, 15, "清理旧数据...")
    delete_expr = f'filename == "{filename}"'
    try:
        milvus_manager.delete(delete_expr)
    except Exception:
        pass
    try:
        parent_chunk_store.delete_by_filename(filename)
    except Exception:
        pass

    task_manager.update_progress(task_id, 25, "加载文档并分块...")
    try:
        new_docs = loader.load_document(file_path, filename)
    except Exception as doc_err:
        raise Exception(f"文档处理失败: {doc_err}")

    if not new_docs:
        raise Exception("文档处理失败，未能提取内容")

    task_manager.update_progress(task_id, 50, "文档分块完成，准备向量化...")
    parent_docs = [doc for doc in new_docs if int(doc.get("chunk_level", 0) or 0) in (1, 2)]
    leaf_docs = [doc for doc in new_docs if int(doc.get("chunk_level", 0) or 0) == 3]
    if not leaf_docs:
        raise Exception("文档处理失败，未生成可检索叶子分块")

    task_manager.update_progress(task_id, 60, "存储父级分块...")
    parent_chunk_store.upsert_documents(parent_docs)

    task_manager.update_progress(task_id, 70, "正在生成向量并写入数据库...")
    milvus_writer.write_documents(leaf_docs)

    task_manager.update_progress(task_id, 95, "清理临时文件...")
    try:
        os.remove(file_path)
    except Exception:
        pass

    return {
        "filename": filename,
        "chunks_processed": len(leaf_docs),
        "message": (
            f"成功上传并处理 {filename}，叶子分块 {len(leaf_docs)} 个，"
            f"父级分块 {len(parent_docs)} 个"
        ),
    }


@router.post("/documents/upload", response_model=UploadTaskCreateResponse)
async def upload_document(file: UploadFile = File(...), _: User = Depends(require_admin)):
    """上传文档（异步处理，立即返回 task_id）"""
    filename = file.filename or ""
    file_lower = filename.lower()
    if not filename:
        raise HTTPException(status_code=400, detail="文件名不能为空")
    if not re.match(r'^[\w\-. ]+$', filename):
        raise HTTPException(status_code=400, detail="文件名包含非法字符")
    if not (
        file_lower.endswith(".pdf")
        or file_lower.endswith((".docx", ".doc"))
        or file_lower.endswith((".xlsx", ".xls"))
    ):
        raise HTTPException(status_code=400, detail="仅支持 PDF、Word 和 Excel 文档")

    content = await file.read()
    if len(content) > MAX_FILE_SIZE:
        raise HTTPException(status_code=400, detail=f"文件大小超过限制（最大 {MAX_FILE_SIZE // 1024 // 1024}MB）")

    os.makedirs(UPLOAD_DIR, exist_ok=True)
    file_path = str(UPLOAD_DIR / filename)
    with open(file_path, "wb") as f:
        f.write(content)

    task_id = task_manager.create_task(filename)

    MAX_RETRIES = 3
    RETRY_DELAY = 5

    async def run_background():
        last_error = None
        for attempt in range(1, MAX_RETRIES + 1):
            try:
                if attempt > 1:
                    task_manager.update_progress(task_id, 0, f"重试第 {attempt} 次...")
                    await asyncio.sleep(RETRY_DELAY)
                result = await asyncio.to_thread(_process_document_upload, file_path, filename, task_id)
                task_manager.complete_task(task_id, result)
                return
            except Exception as e:
                last_error = e
                if attempt < MAX_RETRIES:
                    task_manager.update_progress(task_id, 0, f"第 {attempt} 次失败，{RETRY_DELAY}秒后重试...")
                else:
                    task_manager.fail_task(task_id, str(last_error))

    asyncio.create_task(run_background())

    return UploadTaskCreateResponse(
        task_id=task_id,
        filename=filename,
        message="文件已接收，正在后台处理",
    )


@router.get("/documents/tasks/{task_id}", response_model=UploadTaskStatus)
async def get_upload_task_status(task_id: str, _: User = Depends(require_admin)):
    """查询上传任务状态"""
    task = task_manager.get_task(task_id)
    if not task:
        raise HTTPException(status_code=404, detail="任务不存在或已过期")
    return UploadTaskStatus(
        task_id=task["task_id"],
        filename=task["filename"],
        status=task["status"],
        progress=task.get("progress", 0),
        message=task.get("message", ""),
        result=task.get("result"),
        error=task.get("error"),
    )


@router.delete("/documents/{filename}", response_model=DocumentDeleteResponse)
async def delete_document(filename: str, _: User = Depends(require_admin)):
    """删除文档在 Milvus 中的向量（保留本地文件，管理员）"""
    try:
        milvus_manager.init_collection()

        delete_expr = f'filename == "{filename}"'
        result = milvus_manager.delete(delete_expr)
        parent_chunk_store.delete_by_filename(filename)

        return DocumentDeleteResponse(
            filename=filename,
            chunks_deleted=result.get("delete_count", 0) if isinstance(result, dict) else 0,
            message=f"成功删除文档 {filename} 的向量数据（本地文件已保留）",
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"删除文档失败: {str(e)}")


@router.delete("/documents/batch", response_model=DocumentBatchDeleteResponse)
async def batch_delete_documents(request: DocumentBatchDeleteRequest, _: User = Depends(require_admin)):
    """批量删除文档（管理员）"""
    if not request.filenames:
        raise HTTPException(status_code=400, detail="文件名列表不能为空")

    results = []
    total_chunks = 0

    for filename in request.filenames:
        try:
            milvus_manager.init_collection()
            delete_expr = f'filename == "{filename}"'
            result = milvus_manager.delete(delete_expr)
            parent_chunk_store.delete_by_filename(filename)

            chunks = result.get("delete_count", 0) if isinstance(result, dict) else 0
            total_chunks += chunks
            results.append(DocumentDeleteResponse(
                filename=filename,
                chunks_deleted=chunks,
                message=f"成功删除 {filename}",
            ))
        except Exception as e:
            results.append(DocumentDeleteResponse(
                filename=filename,
                chunks_deleted=0,
                message=f"删除失败: {str(e)}",
            ))

    return DocumentBatchDeleteResponse(
        results=results,
        total_deleted=len([r for r in results if r.chunks_deleted > 0]),
        total_chunks_deleted=total_chunks,
    )


@router.websocket("/ws/documents/{task_id}")
async def websocket_task_progress(websocket: WebSocket, task_id: str):
    await ws_manager.connect(websocket, task_id)
    try:
        while True:
            await websocket.receive_text()
    except WebSocketDisconnect:
        ws_manager.disconnect(websocket, task_id)
    except Exception:
        ws_manager.disconnect(websocket, task_id)
