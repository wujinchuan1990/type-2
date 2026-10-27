import hashlib
import json
import logging
import os
import shutil
import sqlite3
import threading
import uuid
from datetime import datetime, timezone

from native_word import render
from report_types import normalize_report_type

LOGGER = logging.getLogger("report_workbench.exports")


def public_export(manifest, store=None):
    return {
        **{key: value for key, value in manifest.items() if key != "task"},
        "report_type": normalize_report_type(manifest.get("task", {}).get("report_type", "水表")),
        "generated_at": store.first_saved_at("exports", manifest["id"]) if store is not None else "",
    }


def save_manifest(store, manifest):
    data = dict(manifest)
    if "export_revision" not in data:
        try:
            old = store.get("exports", data["id"])
        except KeyError:
            pass
        else:
            data["export_revision"] = old.get("export_revision", old["revision"])
    saved = store.save("exports", data, "export_revision")
    folder = store.directory / "exports" / saved["id"]
    temporary = folder / "manifest.json.tmp"
    temporary.write_text(json.dumps(saved, ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(temporary, folder / "manifest.json")
    return saved


class LayoutChecks:
    def __init__(self, store):
        self.store = store
        self.lock = threading.Lock()
        for manifest in store.list("exports"):
            if manifest.get("layout", {}).get("status") == "running":
                manifest["layout"].update(
                    status="interrupted",
                    error="服务已重启，上次排版检查未完成。原始 Word 下载仍然有效，可手工核对。",
                )
                save_manifest(store, manifest)

    def begin(self, identity):
        if not self.lock.acquire(blocking=False):
            raise ValueError("已有 Word 排版检查正在运行；仍可下载文件、编辑或生成其他报告。")
        try:
            manifest = self.store.get("exports", identity)
            source = self.store.directory / "exports" / identity / manifest["url"].rsplit("/", 1)[-1]
            if hashlib.sha256(source.read_bytes()).hexdigest() != manifest["sha256"]:
                raise ValueError("生成文件已被外部修改，不能继续自动检查。请重新生成；原文件未被覆盖。")
            attempt = uuid.uuid4().hex
            manifest["layout"] = {
                "status": "running", "stage": "queued", "attempt_id": attempt,
                "started_at": datetime.now(timezone.utc).isoformat(), "error": "",
            }
            manifest["pdf_url"] = ""
            manifest["checked_url"] = ""
            manifest.pop("pages", None)
            return save_manifest(self.store, manifest)
        except (KeyError, ValueError, OSError, sqlite3.Error):
            self.lock.release()
            raise

    def execute(self, identity):
        manifest = self.store.get("exports", identity)
        attempt = manifest["layout"]["attempt_id"]
        folder = self.store.directory / "exports" / identity

        def progress(stage):
            current = self.store.get("exports", identity)
            current["layout"]["stage"] = stage
            save_manifest(self.store, current)

        try:
            check_folder = folder / "checks" / attempt
            check_folder.mkdir(parents=True)
            source = folder / manifest["url"].rsplit("/", 1)[-1]
            working = check_folder / f"word-{attempt}.docx"
            pdf = check_folder / "preview.pdf"
            shutil.copyfile(source, working)
            if hashlib.sha256(working.read_bytes()).hexdigest() != manifest["sha256"]:
                raise ValueError("文件在准备工作副本时发生变化，请重新生成后再检查。")
            information = render(working, pdf, progress=progress, timeout_seconds=90)
            current = self.store.get("exports", identity)
            current["layout"].update(
                status="rendered", stage="complete", error="",
                completed_at=datetime.now(timezone.utc).isoformat(),
                checked_sha256=hashlib.sha256(working.read_bytes()).hexdigest(),
            )
            current.update(
                pages=information["pages"],
                checked_url=f"/api/exports/{identity}/checks/{attempt}/{working.name}",
                pdf_url=f"/api/exports/{identity}/checks/{attempt}/preview.pdf",
            )
            save_manifest(self.store, current)
        except (ValueError, OSError, RuntimeError, sqlite3.Error) as exc:
            LOGGER.exception("Independent Word layout check failed for export %s", identity)
            current = self.store.get("exports", identity)
            current["layout"].update(
                status="failed", error=str(exc),
                completed_at=datetime.now(timezone.utc).isoformat(),
            )
            save_manifest(self.store, current)
        finally:
            self.lock.release()
