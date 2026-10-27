"""Independent local report-splitting jobs and explicit directory grants."""

from copy import deepcopy
from datetime import datetime, timezone
import hashlib
import json
import logging
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import threading
import unicodedata
import uuid

from fastapi import APIRouter, BackgroundTasks, Depends, File, UploadFile
from fastapi.responses import Response
from download_transport import DownloadOptions, download_options, download_response
from pydantic import BaseModel, ConfigDict, Field, StrictInt, field_validator
from typing import List, Literal, Optional
from starlette.concurrency import run_in_threadpool
import pymupdf

import native_word
from docx_images import remove_unused_images
import separation_documents as sd
import separation_layout as layout
from storage import ConflictError


LOGGER = logging.getLogger("report_workbench")
MAX_UPLOAD = 32 * 1024 * 1024


def _macos_folder_prompt():
    script = '''on run
    try
        return POSIX path of (choose folder with prompt "选择拆分Word文件的保存目录（不覆盖已有文件）")
    on error number -128
        return "__CANCELLED__"
    end try
end run'''
    try:
        result = subprocess.run(["/usr/bin/osascript", "-e", script],
                                capture_output=True, text=True, timeout=120)
    except subprocess.TimeoutExpired as exc:
        raise ValueError("选择目录超时，未保存目录；请重新打开目录选择窗口。") from exc
    if result.returncode:
        raise ValueError("无法打开本机目录选择窗口：" + result.stderr.strip())
    return result.stdout.strip()


def _windows_folder_prompt():
    """弹出Windows目录选择对话框，取消时返回None。"""
    import ctypes
    from ctypes import wintypes

    ole32 = ctypes.windll.ole32
    shell32 = ctypes.windll.shell32
    ole32.CoInitializeEx(None, 0)  # COINIT_APARTMENTTHREADED

    class BROWSEINFOW(ctypes.Structure):
        _fields_ = [("hwndOwner", wintypes.HWND), ("pidlRoot", ctypes.c_void_p),
                    ("pszDisplayName", wintypes.LPWSTR), ("lpszTitle", wintypes.LPCWSTR),
                    ("ulFlags", wintypes.UINT), ("lpfn", ctypes.c_void_p),
                    ("lParam", ctypes.c_void_p), ("iImage", ctypes.c_int)]

    BIF_RETURNONLYFSDIRS, BIF_NEWDIALOGSTYLE = 0x1, 0x40
    display = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
    info = BROWSEINFOW(None, None, display, "选择拆分Word文件的保存目录（不覆盖已有文件）",
                       BIF_RETURNONLYFSDIRS | BIF_NEWDIALOGSTYLE, None, None, 0)
    pidl = shell32.SHBrowseForFolderW(ctypes.byref(info))
    if not pidl:
        return None
    try:
        buffer = ctypes.create_unicode_buffer(wintypes.MAX_PATH)
        if not shell32.SHGetPathFromIDListW(pidl, buffer):
            raise ValueError("选择的保存位置不是现有目录。")
        return buffer.value
    finally:
        ole32.CoTaskMemFree(pidl)


class SplitPart(BaseModel):
    model_config = ConfigDict(extra="forbid")
    id: str
    kind: Literal["observe", "items", "result"]
    enabled: bool = True
    start_page: StrictInt = Field(ge=1)
    end_page: StrictInt = Field(ge=1)
    filename: str
    directory_token: str = ""

    @field_validator("id")
    @classmethod
    def safe_id(cls, value):
        if not re.fullmatch(r"[A-Za-z0-9_-]{1,80}", value):
            raise ValueError("拆分文件标识无效。")
        return value


class SplitPlan(BaseModel):
    model_config = ConfigDict(extra="forbid")
    model_count: StrictInt = Field(ge=1, le=1000)
    root_directory_token: str
    parts: List[SplitPart] = Field(min_length=1, max_length=100)


class FolderRequest(BaseModel):
    part_id: Optional[str] = None


class ExportRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    plan_revision: StrictInt
    confirm_adjustments: bool = False


def file_hash(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def safe_filename(value):
    if (not value or value != value.strip() or len(value.encode("utf-8")) > 240
            or re.search(r'[/\\:*?"<>|\x00-\x1f]', value) or value.startswith(".")
            or not value.lower().endswith(".docx") or "待填写" in value or "待核对序号" in value):
        raise ValueError("请填写有效且已核对的DOCX文件名；不能含目录、控制字符或待填写占位内容。")
    return value


class Separations:
    def __init__(self, directory, check_docx):
        self.directory = Path(directory) / "separations"
        self.directory.mkdir(exist_ok=True)
        self.lock = threading.RLock()
        self.check_docx = check_docx
        for path in self.directory.glob("*/job.json"):
            job = json.loads(path.read_text(encoding="utf-8"))
            if job["status"] in ("analyzing", "exporting"):
                job.update(status="failed", error="服务在处理中停止。未确认完成，请检查所选目录并主动重试；不会自动重复写入。")
                self.save(job)

    def folder(self, identity):
        if not re.fullmatch(r"[a-f0-9]{32}", identity):
            raise KeyError(identity)
        return self.directory / identity

    def get(self, identity):
        with self.lock:
            path = self.folder(identity) / "job.json"
            if not path.is_file():
                raise KeyError(identity)
            return json.loads(path.read_text(encoding="utf-8"))

    def save(self, job):
        with self.lock:
            folder = self.folder(job["id"])
            path = folder / "job.json"
            temporary = folder / ("state-" + uuid.uuid4().hex + ".tmp")
            temporary.write_text(json.dumps(job, ensure_ascii=False), encoding="utf-8")
            os.replace(temporary, path)

    def public(self, job):
        result = {key: value for key, value in job.items() if not key.startswith("_")}
        result["warnings"] = [warning.replace("保留原位置", "按系统报告规范重新排版")
                              for warning in job.get("warnings", [])]
        return result

    def update(self, identity, **values):
        with self.lock:
            job = self.get(identity)
            job.update(values)
            self.save(job)
            return job

    def progress(self, identity, stage):
        self.update(identity, stage=stage)

    def source(self, job):
        path = self.folder(job["id"]) / "source.docx"
        if not path.is_file() or file_hash(path) != job["_source_hash"]:
            raise ValueError("上传报告副本已丢失或改变，请重新上传；不能复用旧分页。")
        return path

    def begin_analysis(self, identity):
        with self.lock:
            job = self.get(identity)
            if job["status"] in ("analyzing", "exporting"):
                raise ConflictError("本报告正在处理中，请等待完成。")
            return self.update(identity, status="analyzing", stage="等待Word分页", error="", plan=None,
                               _plan=None, _units=[], _bookmarks={}, _numbered=[], _pdf_hash="")

    def analyze(self, identity):
        try:
            job = self.get(identity)
            source = self.source(job)
            folder = self.folder(identity)
            units, count = sd.discover(source)
            marked = folder / "marked.docx"
            names, numbered = sd.prepare_marked(source, marked, units)
            unmarked = folder / "unmarked.docx"
            shutil.copyfile(source, unmarked)
            baseline = native_word.render(unmarked, folder / "source.pdf", timeout_seconds=180,
                                          progress=lambda stage: self.progress(identity, "原报告 · " + stage))
            result = native_word.render(marked, folder / "marked.pdf", bookmark_names=names,
                                        timeout_seconds=180, progress=lambda stage: self.progress(identity, stage))
            sd.compare_pages(folder / "source.pdf", folder / "marked.pdf", 1, baseline["pages"],
                             folder / "analysis-diagnostics")
            pages = []
            with pymupdf.open(folder / "source.pdf") as pdf:
                for index, page in enumerate(pdf):
                    pages.append({"number": index + 1, "text": page.get_text(),
                                  "image_url": f"/api/separations/{identity}/pages/{index + 1}.png"})
            warnings, parts = [], []
            for unit in units:
                unit["start_page"] = result["bookmarks"][unit["start_marker"]]["page"]
                unit["end_page"] = result["bookmarks"][unit["end_marker"]]["page"]
                unit["printed_page"] = result["bookmarks"][unit["start_marker"]]["printed"]
                for candidate in unit.get("leading_markers", []):
                    if result["bookmarks"][candidate["marker"]]["page"] == unit["start_page"]:
                        unit["start_block"] = min(unit["start_block"], candidate["block"])
                after = result["bookmarks"].get(unit.get("after_marker"))
                unit["end_y"] = after["y"] if after and after["page"] == unit["end_page"] else None
                if unit["end_y"] is not None:
                    warnings.append(f"{unit['project']}末页P{unit['end_page']}与其他章节共页：只输出本部分，按系统报告规范重新排版，不带入其他章节。")
                if not 1 <= unit["start_page"] <= unit["end_page"] <= result["pages"]:
                    raise ValueError("Word返回了无效的项目页码，停止拆分。")
                parts.append({**{key: unit[key] for key in ("id", "kind", "project", "start_page", "end_page")},
                              "filename": sd.default_filename(unit, count), "enabled": True, "directory_token": ""})
                if unit["kind"] == "result" and not unit.get("sequence"):
                    warnings.append(unit["project"] + "的项目序号未识别，请手动确认文件名。")
            if count is None:
                warnings.append("未可靠识别型号/规格组数量，请填写数量并核对各文件名。")
            self.source(job)
            self.update(identity, status="ready", stage="已完成Word实际分页，请核对各输出范围", error="",
                        total_pages=result["pages"], model_count=count, pages=pages, parts=parts, warnings=warnings,
                        _units=units, _numbered=numbered, _bookmarks=result["bookmarks"], _pdf_hash=file_hash(folder / "source.pdf"))
        except Exception as exc:
            LOGGER.exception("Report separation analysis failed: %s", identity)
            self.update(identity, status="failed", error=str(exc), stage="分析失败；原文件未修改")

    def choose_folder(self, identity):
        self.get(identity)
        if sys.platform == "win32":
            selected = _windows_folder_prompt()
        else:
            selected = _macos_folder_prompt()
        if selected is None or selected == "__CANCELLED__":
            return {"cancelled": True}
        path = Path(selected).resolve(strict=True)
        if not selected or not path.is_dir():
            raise ValueError("选择的保存位置不是现有目录。")
        token = uuid.uuid4().hex
        with self.lock:
            job = self.get(identity)
            job.setdefault("_directories", {})[token] = str(path)
            self.save(job)
        return {"token": token, "path": str(path)}

    def destination(self, job, root_token, part):
        token = part.directory_token or root_token
        directory = job.get("_directories", {}).get(token)
        if not directory:
            raise ValueError("请通过本机目录选择窗口选择保存位置，不能直接提交任意路径。")
        root = Path(directory)
        if not root.is_dir() or str(root.resolve()) != directory:
            raise ValueError("所选保存目录已改变或不存在，请重新选择。")
        target = root if part.directory_token else root / sd.FOLDERS[part.kind]
        if not target.resolve().is_relative_to(root):
            raise ValueError("分类子目录链接到所选目录之外，请重新选择目录。")
        return target

    def plan(self, identity, request):
        with self.lock:
            job = self.get(identity)
            if job["status"] not in ("ready", "completed", "failed") or not job.get("_units"):
                raise ConflictError("请等待本报告分析完成后再核对拆分清单。")
            self.source(job)
            if len({p.id for p in request.parts}) != len(request.parts):
                raise ValueError("拆分清单包含重复文件标识。")
            outputs, adjustments, pages_seen, names_seen = [], [], set(), set()
            for part in request.parts:
                if not part.enabled:
                    continue
                if not part.start_page <= part.end_page <= job["total_pages"]:
                    raise ValueError("页码须按原报告物理页顺序填写，且不能超过总页数。")
                units = [unit for unit in job["_units"] if unit["kind"] == part.kind
                         and unit["start_page"] <= part.end_page and unit["end_page"] >= part.start_page]
                if not units:
                    raise ValueError(f"{part.filename}：所选页未包含该类可完整拆分的章节/项目。")
                units.sort(key=lambda u: u["start_block"])
                first, last = units[0], units[-1]
                start, end = first["start_page"], last["end_page"]
                page_set = set(range(start, end + 1))
                if page_set & pages_seen:
                    raise ValueError("多个输出文件覆盖了相同原报告页。请合并同类项目范围或取消重复文件，不能重复导出这些页。")
                pages_seen |= page_set
                if (start, end) != (part.start_page, part.end_page):
                    adjustments.append({"id": part.id, "requested_start": part.start_page, "requested_end": part.end_page,
                                        "start_page": start, "end_page": end,
                                        "message": f"为保留完整章节/试验项目及跨页表格，建议P{part.start_page}～P{part.end_page}调整为P{start}～P{end}。"})
                filename = safe_filename(part.filename)
                directory = self.destination(job, request.root_directory_token, part)
                target = directory / filename
                folded = unicodedata.normalize("NFC", str(target)).casefold()
                existing = {unicodedata.normalize("NFC", p.name).casefold() for p in directory.iterdir()} if directory.is_dir() else set()
                if folded in names_seen or unicodedata.normalize("NFC", filename).casefold() in existing:
                    raise ValueError(f"文件名冲突，不能覆盖已有文件：{target}")
                names_seen.add(folded)
                outputs.append({"id": part.id, "kind": part.kind, "start_page": start, "end_page": end,
                                "filename": filename, "directory": str(directory), "full_path": str(target),
                                "_start_block": first["start_block"], "_end_block": last["end_block"],
                                "_printed_page": first["printed_page"], "_end_y": last.get("end_y")})
            if not outputs:
                raise ValueError("请至少选择一个输出文件。")
            revision = job.get("_plan_revision", 0) + 1
            internal = {"revision": revision, "outputs": outputs, "adjustments": adjustments,
                        "layout_policy": layout.POLICY}
            public = deepcopy(internal)
            public["outputs"] = [{k: v for k, v in output.items() if not k.startswith("_")} for output in outputs]
            job.update(parts=[p.model_dump() for p in request.parts], model_count=request.model_count,
                       root_directory_token=request.root_directory_token,
                       root_directory_label=job.get("_directories", {}).get(request.root_directory_token, ""),
                       plan=public, _plan=internal, _plan_revision=revision, status="ready", error="")
            self.save(job)
            return job

    def begin_export(self, identity, request):
        with self.lock:
            job = self.get(identity)
            if job["status"] != "ready" or not job.get("_plan"):
                raise ConflictError("请先重新核对并确认输出清单。")
            if job["_plan"]["revision"] != request.plan_revision or not request.confirm_adjustments:
                raise ConflictError("输出清单已改变或尚未确认建议范围，请重新核对。")
            if job["_plan"].get("layout_policy") != layout.POLICY:
                raise ConflictError("拆分输出已改为报告规范排版，请重新保存并检查计划后确认。")
            return self.update(identity, status="exporting", stage="正在按报告规范排版并核验内容，尚未写入目标目录", error="", exports=[])

    def export(self, identity):
        current_file = ""
        try:
            job = self.get(identity)
            source = self.source(job)
            folder = self.folder(identity)
            reference = folder / "source.pdf"
            if file_hash(reference) != job["_pdf_hash"]:
                raise ValueError("原报告分页基准已改变，请重新分析。")
            staging = folder / ("export-" + uuid.uuid4().hex)
            staging.mkdir()
            verified = []
            for output in job["_plan"]["outputs"]:
                current_file = output["filename"]
                self.progress(identity, "核验 " + output["filename"])
                path = staging / (output["id"] + ".docx")
                sd.extract(source, path, output["_start_block"], output["_end_block"],
                           1, job["total_pages"], job["_numbered"], job["_bookmarks"], freeze_total_pages=False,
                           report_number=sd.EXPORT_REPORT_NUMBER)
                ranges = [(max(unit["start_block"], output["_start_block"]) - output["_start_block"],
                           min(unit["end_block"], output["_end_block"]) - output["_start_block"])
                          for unit in job["_units"] if unit["kind"] == "result"
                          and output["_start_block"] <= unit["start_block"] < output["_end_block"]]
                signature = layout.normalize(path, output["kind"], ranges,
                                             result_format=output["kind"] == "result")
                native_word.render(path, path.with_suffix(".pdf"), timeout_seconds=180,
                                   verify_editing_pages=output["kind"] == "result",
                                   equalize_result_rows=output["kind"] == "result",
                                   progress=lambda stage, name=output["filename"]: self.progress(identity, name + " · " + stage))
                checks = layout.verify_rendered(path, path.with_suffix(".pdf"), signature, output["kind"],
                                                output["start_page"], output["end_page"])
                self.progress(identity, "清理未使用图片 · " + output["filename"])
                checks["image_cleanup"] = remove_unused_images(path)
                verified.append((output, path, checks))
            self.source(job)
            self.progress(identity, "全部核验通过，正在写入已确认的目录")
            created, exported = [], []
            try:
                for output, path, checks in verified:
                    target = Path(output["full_path"])
                    part = next(SplitPart.model_validate(p) for p in job["parts"] if p["id"] == output["id"])
                    directory = self.destination(job, job["root_directory_token"], part)
                    if directory != target.parent:
                        raise ValueError("保存位置已变化，请重新核对清单。")
                    directory.mkdir(exist_ok=True)
                    if any(unicodedata.normalize("NFC", p.name).casefold() == unicodedata.normalize("NFC", target.name).casefold()
                           for p in directory.iterdir()):
                        raise FileExistsError(f"文件已存在，未覆盖：{target}")
                    with target.open("xb") as handle, path.open("rb") as data:
                        stat = os.fstat(handle.fileno())
                        created.append((target, stat.st_dev, stat.st_ino))
                        shutil.copyfileobj(data, handle)
                        handle.flush()
                        os.fsync(handle.fileno())
                    if file_hash(target) != file_hash(path):
                        raise OSError("保存后的文件校验失败：" + str(target))
                    exported.append({"id": output["id"], "filename": output["filename"], "path": str(target),
                                     "url": f"/api/separations/{identity}/files/{output['id']}",
                                     "checks": checks, "_copy": str(path.relative_to(folder))})
            except (OSError, ValueError):
                for target, device, inode in created:
                    if target.exists() and (target.stat().st_dev, target.stat().st_ino) == (device, inode):
                        target.unlink()
                raise
            self.update(identity, status="completed", stage="拆分文件已核验并保存", error="",
                        exports=[{k: v for k, v in item.items() if not k.startswith("_")} for item in exported],
                        _exports=exported)
        except Exception as exc:
            LOGGER.exception("Report separation export failed: %s", identity)
            self.update(identity, status="failed", error=(current_file + "：" if current_file else "") + str(exc),
                        stage="停止交付；请查看错误并核对范围/版式")


def routes(directory, check_docx):
    service = Separations(directory, check_docx)
    router = APIRouter(prefix="/api/separations")

    @router.get("")
    def listing():
        return sorted([{"id": job["id"], "name": job["name"], "status": job["status"], "created_at": job["created_at"]}
                       for path in service.directory.glob("*/job.json")
                       for job in [service.get(path.parent.name)]], key=lambda j: j["created_at"], reverse=True)

    @router.post("/upload")
    async def upload(background: BackgroundTasks, file: UploadFile = File(...)):
        if not file.filename or not file.filename.lower().endswith(".docx"):
            raise ValueError("请选择可编辑DOCX报告，不支持PDF、扫描件或仅改后缀的文件。")
        identity = uuid.uuid4().hex
        folder = service.folder(identity)
        folder.mkdir()
        path = folder / "source.docx"
        try:
            size = 0
            with path.open("xb") as handle:
                while True:
                    chunk = await file.read(1024 * 1024)
                    if not chunk:
                        break
                    size += len(chunk)
                    if size > MAX_UPLOAD:
                        raise ValueError("报告最大支持32MB。")
                    handle.write(chunk)
            check_docx(path)
        except (ValueError, OSError):
            path.unlink(missing_ok=True)
            folder.rmdir()
            raise
        finally:
            await file.close()
        job = {"id": identity, "name": Path(file.filename).name, "status": "uploaded", "stage": "", "error": "",
               "created_at": datetime.now(timezone.utc).isoformat(), "_source_hash": file_hash(path),
               "pages": [], "parts": [], "warnings": [], "exports": [], "plan": None,
               "total_pages": 0, "model_count": None, "root_directory_token": "", "root_directory_label": ""}
        service.save(job)
        job = service.begin_analysis(identity)
        background.add_task(service.analyze, identity)
        return service.public(job)

    @router.get("/{identity}")
    def detail(identity: str):
        return service.public(service.get(identity))

    @router.post("/{identity}/analyze")
    def analyze(identity: str, background: BackgroundTasks):
        job = service.begin_analysis(identity)
        background.add_task(service.analyze, identity)
        return service.public(job)

    @router.post("/{identity}/folder")
    async def folder(identity: str, request: FolderRequest):
        return await run_in_threadpool(service.choose_folder, identity)

    @router.post("/{identity}/plan")
    def plan(identity: str, request: SplitPlan):
        return service.public(service.plan(identity, request))

    @router.post("/{identity}/export")
    def export(identity: str, request: ExportRequest, background: BackgroundTasks):
        job = service.begin_export(identity, request)
        background.add_task(service.export, identity)
        return service.public(job)

    @router.get("/{identity}/source.pdf")
    def pdf(identity: str, options: DownloadOptions = Depends(download_options)):
        job = service.get(identity)
        if not job.get("_pdf_hash"):
            raise ValueError("原报告尚未完成Word分页。")
        return download_response(service.folder(identity) / "source.pdf", options, media_type="application/pdf")

    @router.get("/{identity}/pages/{page}.png")
    def page(identity: str, page: int):
        job = service.get(identity)
        if not 1 <= page <= job["total_pages"]:
            raise ValueError("预览页码超出报告范围。")
        with pymupdf.open(service.folder(identity) / "source.pdf") as document:
            image = document[page - 1].get_pixmap(matrix=pymupdf.Matrix(.9, .9), alpha=False)
            return Response(image.tobytes("png"), media_type="image/png")

    @router.get("/{identity}/files/{file_id}")
    def output(identity: str, file_id: str, options: DownloadOptions = Depends(download_options)):
        job = service.get(identity)
        item = next((item for item in job.get("_exports", []) if item["id"] == file_id), None)
        if not item:
            raise KeyError(file_id)
        return download_response(service.folder(identity) / item["_copy"], options, filename=item["filename"])

    return router
