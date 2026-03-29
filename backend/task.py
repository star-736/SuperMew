import asyncio
import uuid
from datetime import datetime
from typing import Any, Callable, Optional

from cache import cache


class TaskStatus:
    PENDING = "pending"
    PROCESSING = "processing"
    COMPLETED = "completed"
    FAILED = "failed"


class TaskManager:
    PREFIX = "upload_task"
    TTL = 3600

    def create_task(self, filename: str) -> tuple[str, asyncio.Event]:
        task_id = str(uuid.uuid4())
        task_data = {
            "task_id": task_id,
            "filename": filename,
            "status": TaskStatus.PENDING,
            "progress": 0,
            "message": "任务已创建",
            "result": None,
            "error": None,
            "created_at": datetime.now().isoformat(),
        }
        cache.set_json(f"{self.PREFIX}:{task_id}", task_data, self.TTL)
        return task_id

    def update_progress(self, task_id: str, progress: int, message: str, status: str = TaskStatus.PROCESSING):
        key = f"{self.PREFIX}:{task_id}"
        task_data = cache.get_json(key) or {}
        task_data["progress"] = progress
        task_data["message"] = message
        task_data["status"] = status
        task_data["updated_at"] = datetime.now().isoformat()
        cache.set_json(key, task_data, self.TTL)

    def complete_task(self, task_id: str, result: Any):
        key = f"{self.PREFIX}:{task_id}"
        task_data = cache.get_json(key) or {}
        task_data["status"] = TaskStatus.COMPLETED
        task_data["progress"] = 100
        task_data["message"] = "处理完成"
        task_data["result"] = result
        task_data["updated_at"] = datetime.now().isoformat()
        cache.set_json(key, task_data, self.TTL)

    def fail_task(self, task_id: str, error: str):
        key = f"{self.PREFIX}:{task_id}"
        task_data = cache.get_json(key) or {}
        task_data["status"] = TaskStatus.FAILED
        task_data["message"] = f"处理失败: {error}"
        task_data["error"] = error
        task_data["updated_at"] = datetime.now().isoformat()
        cache.set_json(key, task_data, self.TTL)

    def get_task(self, task_id: str) -> Optional[dict]:
        return cache.get_json(f"{self.PREFIX}:{task_id}")


task_manager = TaskManager()
