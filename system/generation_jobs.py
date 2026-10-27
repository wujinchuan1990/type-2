"""Persistent generation snapshots; unfinished jobs never masquerade as exports."""

import logging
import threading
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime, timezone

from fastapi import HTTPException

LOGGER = logging.getLogger("report_workbench.jobs")


class GenerationJobs:
    def __init__(self, store):
        self.store = store
        self.executor = ThreadPoolExecutor(max_workers=1, thread_name_prefix="report-generation")
        self.lock = threading.Lock()
        # 轮询接口的高频读取走内存缓存，避免每次都全量读取并解析jobs表
        self._cache = {}
        for job in store.list("generation_jobs"):
            self._cache[job["id"]] = job
            if job["status"] in ("queued", "running"):
                self.update(job["id"], status="interrupted", stage="服务重启，未自动重新生成",
                            error="上次生成中断，请核对历史版本后主动重试。")

    def update(self, identity, **values):
        with self.lock:
            job = dict(self._cache.get(identity) or self.store.get("generation_jobs", identity))
            job.update(values)
            updated = datetime.now(timezone.utc).isoformat()
            job["updated"] = updated
            saved = self.store.save("generation_jobs", job)
            saved["updated_at"] = updated
            self._cache[identity] = saved
            return saved

    def list_public(self):
        with self.lock:
            jobs = [dict(job) for job in self._cache.values()]
        jobs.sort(key=lambda job: job.get("updated_at", ""), reverse=True)
        return [self.public(job) for job in jobs]

    @staticmethod
    def public(job):
        return {key: value for key, value in job.items() if key not in ("request", "source_hashes")}

    def submit(self, request, generate, source_hashes=None):
        job = self.store.save("generation_jobs", {
            "status": "queued", "stage": "已保存生成快照，等待处理", "error": "", "export_id": "",
            "task_id": request.task.id, "task_revision": request.task.revision,
            "task_name": request.task.name, "created": datetime.now(timezone.utc).isoformat(),
            "request": request.model_dump(),
            "source_hashes": source_hashes or {},
        })
        job["updated_at"] = datetime.now(timezone.utc).isoformat()
        with self.lock:
            self._cache[job["id"]] = job
        self.executor.submit(self.execute, job["id"], request.model_copy(deep=True), generate)
        return self.public(job)

    def execute(self, identity, request, generate):
        try:
            self.update(identity, status="running", stage="正在校验生成快照")
            result = generate(request, snapshot=True,
                              source_hashes=self.store.get("generation_jobs", identity)["source_hashes"],
                              progress=lambda text: self.update(identity, stage=text))
            self.update(identity, status="completed", stage="两份Word已生成",
                        export_id=result["id"])
        except Exception as exc:
            LOGGER.exception("Background report generation failed: %s", identity)
            error = str(exc.detail) if isinstance(exc, HTTPException) else str(exc)
            self.update(identity, status="failed", stage="生成失败，未发布未完成版本", error=error)
