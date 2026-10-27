import base64
import hashlib
import json
import logging
import os
import re
import time
import uuid
import zipfile
from contextlib import asynccontextmanager
from pathlib import Path
from typing import Literal, Optional
from urllib.parse import urlparse

from fastapi import BackgroundTasks, Depends, FastAPI, File, Form, HTTPException, Query, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response
from PIL import Image
from fastapi.staticfiles import StaticFiles
from starlette.concurrency import run_in_threadpool
from starlette.middleware.trustedhost import TrustedHostMiddleware

from contracts import (
    AUTOMATIC_OVERVIEW_FIELDS, OBSERVATION_MANUAL_FIELDS, OUTLINE_OVERVIEW_FIELDS, PHOTO_SLOTS, REPORT_ENVIRONMENT_FIELDS, Candidate, Equipment,
    GenerateRequest, ObservationSummary, OverviewField, PhotoEdit, Project, RecoveryRequest, ReportEnvironment, ResultMaterialRequest, ResultMaterialScheme, SampleOutputScheme, Task, photo_slot_group,
    photo_layout_ids, GAS_UNUSED_PHOTO_SLOTS, GAS_PHOTO_REPORT_TYPES,
)
from environment import aggregate_project_environment
from atmospheric_pressure import summarize_pressure
from upload_types import classified_upload
from report_types import FLOWMETER_TYPES, INDEPENDENT_METER_TYPES, normalize_report_type
from evaluation_summary import summarize_evaluation
from export_service import LayoutChecks, public_export, save_manifest
from export_comparison import compare_exports
from construction import construction_text
from curve_results import refresh_curve_results
import flowmeter_records
import flowmeter_tasks
import gas_records
import gas_flowmeter_excel
from photo_documents import extract_component_names, photo_width_limit
from photo_uploads import MAX_PHOTO_BYTES, edited_photo_bytes, edited_photo_size, photo_path, save_photo
from photo_preview import PhotoPreviews
from record_documents import compose_original_record
from observation_documents import manual_originals, outline_values
from observation_certificates import refresh_certificates
from result_requirements import sync_project_requirements
from result_pagination import prepare_metadata_groups
from storage import ConflictError, Store
from generation_jobs import GenerationJobs
from parse_cache import ParseCache
from backups import BackupGate, routes as backup_routes
from validation import parse_date, validate_task
from validation_issues import IssueMessages
from server_access import ServerAccess
from download_transport import DownloadOptions, download_options, download_response

BASE = Path(__file__).resolve().parent
ROOT = BASE.parent
LOGGER = logging.getLogger("report_workbench")
MAX_UPLOAD = 32 * 1024 * 1024


def check_docx(path):
    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            if len(names) != len(set(names)):
                raise ValueError("DOCX 包中存在重复部件。")
            if "word/document.xml" not in names:
                raise ValueError("不是有效的 Word DOCX 文档。")
            if len(names) > 10000 or sum(item.file_size for item in archive.infolist()) > 160 * 1024 * 1024:
                raise ValueError("文档解压后过大，无法安全处理。")
            if any("vbaProject" in name for name in names):
                raise ValueError("不支持含宏的文档，请另存为无宏 DOCX。")
            for name in names:
                if name.endswith(".xml") or name.endswith(".rels"):
                    data = archive.read(name)
                    if b"<!DOCTYPE" in data or b"<!ENTITY" in data:
                        raise ValueError("不支持带外部实体定义的文档。")
    except zipfile.BadZipFile as exc:
        raise ValueError("文件不是有效 DOCX，请勿仅修改文件扩展名。") from exc


def create_app(data_dir=None, *, template_root=None, server_access: Optional[ServerAccess] = None,
               cloud_photo_optimization=False):
    import documents

    if server_access is not None and not isinstance(server_access, ServerAccess):
        raise ValueError("服务器访问策略必须经过明确配置。")
    if template_root is not None:
        template_root = Path(template_root).resolve(strict=True)
        if not template_root.is_dir():
            raise ValueError("模板根目录必须是已存在的目录。")

    def library_root():
        return template_root if template_root is not None else ROOT

    store = Store(Path(data_dir or os.environ.get("REPORT_DATA_DIR", BASE / "data")))
    layout_checks = LayoutChecks(store)
    photo_previews = PhotoPreviews(store, library_root())
    generation_jobs = GenerationJobs(store)

    @asynccontextmanager
    async def lifespan(_app):
        yield
        await run_in_threadpool(generation_jobs.executor.shutdown, wait=True)

    app = FastAPI(title="型式评价报告工作台", docs_url=None, redoc_url=None, lifespan=lifespan)
    app.state.store = store
    app.state.generation_jobs = generation_jobs
    record_cache = ParseCache()
    backup_gate = BackupGate()
    app.include_router(backup_routes(store, library_root(), backup_gate))
    record_parser_version = hashlib.sha256(
        (BASE / "documents.py").read_bytes() + (BASE / "flowmeter_records.py").read_bytes()).hexdigest()

    app.add_middleware(TrustedHostMiddleware, allowed_hosts=(
        [server_access.hostname] if server_access else ["127.0.0.1", "localhost", "testserver"]))

    @app.middleware("http")
    async def local_only(request: Request, call_next):
        origin = request.headers.get("origin")
        if server_access:
            rejected = server_access.reject(request)
            if rejected is not None:
                return rejected
        elif origin and (urlparse(origin).netloc != request.headers.get("host")
                         or urlparse(origin).scheme != "http"):
            return JSONResponse({"detail": "仅允许本机同源访问。"}, status_code=403)
        guarded = not request.url.path.startswith(("/api/backups", "/static"))
        if guarded:
            await run_in_threadpool(backup_gate.enter)
        try:
            response = await call_next(request)
        finally:
            if guarded:
                backup_gate.leave()
        response.headers["X-Content-Type-Options"] = "nosniff"
        response.headers["X-Frame-Options"] = "DENY"
        response.headers["Cache-Control"] = "no-store"
        response.headers["Content-Security-Policy"] = (
            "default-src 'self'; script-src 'self'; style-src 'self'; "
            "img-src 'self' data:; object-src 'none'; base-uri 'none'; frame-ancestors 'none'"
        )
        return response

    @app.exception_handler(ConflictError)
    async def conflict(_request, exc):
        return JSONResponse({"detail": str(exc)}, status_code=409)

    @app.exception_handler(ValueError)
    async def invalid(_request, exc):
        return JSONResponse({"detail": str(exc)}, status_code=422)

    @app.exception_handler(KeyError)
    async def missing(_request, _exc):
        return JSONResponse({"detail": "记录不存在，请刷新后重试。"}, status_code=404)

    @app.exception_handler(Exception)
    async def unexpected(_request, exc):
        LOGGER.exception("Report workbench request failed", exc_info=exc)
        return JSONResponse({"detail": "处理失败，未保存为成功结果。请查看终端错误记录。"}, status_code=500)

    def upload_record(identity, kind=None):
        info = store.get("uploads", identity)
        if kind and info["kind"] != kind:
            raise ValueError("上传文件类型与任务引用不一致。")
        extension = info.get("extension", ".docx")
        if extension not in (".docx", ".xlsx"):
            raise ValueError("不支持的上传文件格式。")
        path = store.directory / "uploads" / (identity + extension)
        if not path.is_file():
            raise ValueError("上传文件丢失，请重新上传。")
        if hashlib.sha256(path.read_bytes()).hexdigest() != info["sha256"]:
            raise ValueError("上传文件内容已被外部修改，请重新上传并核对。")
        return info, path

    def validate_assets(task):
        if task.report_type in INDEPENDENT_METER_TYPES:
            task.projects = [Project.model_validate(documents.reconcile_project_rows(library_root(), project.model_dump()))
                             for project in task.projects]
        catalog = documents.catalog(library_root())
        assets = {asset["id"]: asset for asset in catalog["assets"]}
        for identity, kind in ((task.items_asset_id, "items"), (task.observe_asset_id, "observe")):
            if identity and (identity not in assets or assets[identity]["kind"] != kind):
                raise ValueError("选中的素材已丢失或版本发生变化，请重新选择。")
        for project in task.projects:
            from water_results import validate_selection

            validate_selection(project.model_dump(), task.model_dump())
            if project.asset_id not in assets or assets[project.asset_id]["kind"] != "result":
                raise ValueError(f"{project.name}：结果素材已丢失或发生变化。")
            definition = documents.project_from_asset(library_root(), project.asset_id)
            if task.report_type == "水表" and project.water_result_labels:
                from water_results import project_schema

                doc = documents._load(documents._asset(library_root(), project.asset_id, "result")["_path"])
                definition["rows"] = project_schema(doc, project.model_dump())
            if project.name != definition["name"] or not documents.result_variant_matches(
                    assets[project.asset_id], project.variant):
                raise ValueError("项目名称或方案与素材不一致，请重新选择。")
            if [row.id for row in project.rows] != [row["id"] for row in definition["rows"]]:
                raise ValueError(f"{project.name}：结果行与素材不一致，不能截断或追加未映射数据。")
            if [row.sample_label for row in project.rows] != [row["sample_label"] for row in definition["rows"]]:
                raise ValueError(f"{project.name}：样机代号与素材不一致，请重新选择素材。")
        required_keys = {field["key"] for field in catalog["overview"]}
        keys = [field.key for field in task.overview]
        if set(keys) != required_keys or len(keys) != len(set(keys)):
            raise ValueError("基本信息总览字段不完整或重复，请重新载入任务。")
    def order_task(task):
        sync_report_type(task)
        if task.items_asset_id and task.projects:
            task.projects = [
                Project.model_validate(project)
                for project in documents.order_projects(
                    library_root(), task.items_asset_id, [project.model_dump() for project in task.projects]
                )
            ]
        task = sync_samples(task)
        from water_results import validate_selection

        for project in task.projects:
            validate_selection(project.model_dump(), task.model_dump())
        if task.report_type in INDEPENDENT_METER_TYPES:
            task.projects = [Project.model_validate(documents.reconcile_project_rows(library_root(), project.model_dump()))
                             for project in task.projects]
        return task

    def current_records(task, refresh_records=False):
        parsed = []
        for identity in dict.fromkeys(task.record_ids):
            info, path = upload_record(identity, "record")
            validate_upload_type(task, info, "record")
            excel_parser = {"燃气表": gas_records, "流量计（气体）": gas_flowmeter_excel}.get(task.report_type)
            if path.suffix == ".xlsx" and excel_parser is None:
                raise ValueError("Excel原始记录识别仅适用于燃气表和流量计（气体）。")
            excel_cache_valid = (
                path.suffix == ".xlsx"
                and info.get("parser") == excel_parser.PARSER
                and info.get("report_type") == task.report_type
                and isinstance(info.get("fields"), dict)
                and isinstance(info.get("records"), list)
                and isinstance(info.get("warnings"), list)
            )
            if excel_cache_valid and not refresh_records:
                current = info
            else:
                current = (excel_parser.parse_upload(path) if path.suffix == ".xlsx"
                           else record_cache.read(
                               path, (task.report_type, record_parser_version),
                               lambda: (flowmeter_records.parse_upload(path) if task.report_type in FLOWMETER_TYPES
                                        else documents.parse_upload(path, "record")),
                               refresh=refresh_records))
            current = json.loads(json.dumps(current, ensure_ascii=False).replace(path.name, info["name"]))
            if path.suffix == ".xlsx" and (refresh_records or not excel_cache_valid):
                current = store.save("uploads", {**info, **current})
            parsed.append(current)
        return parsed

    def sync_record_fields(task, parsed):
        field = next((field for field in task.overview if field.key == "report_number"), None)
        if field is None:
            return task
        previous = {candidate.value for candidate in field.candidates}
        was_auto = not field.confirmed and not field.manual_override and previous == {field.value}
        candidates = []
        for current in parsed:
            candidates.extend(Candidate.model_validate(value) for value in current["fields"].get("report_number", []))
        field.candidates = candidates
        unique = {candidate.value for candidate in candidates}
        previous_value = field.value
        if not field.manual_override and (not field.value.strip() or was_auto):
            field.value = next(iter(unique)) if len(unique) == 1 else ""
        if previous != unique or previous_value != field.value:
            field.confirmed = False
        return task

    def sync_samples(task, strict=True, parsed_records=None, fill_flow_conditions=False):
        sync_report_type(task)
        master = documents._load(library_root() / "report model/report.docx")
        supply = next((field for field in task.overview if field.key == "supply_voltage"), None)
        if supply is not None:
            defaults = documents._field_candidates(master, "报告模板")
            supply.old = "\n".join(dict.fromkeys(item["value"] for item in defaults.get("supply_voltage", [])))
        task.evaluation_summary = ReportEnvironment.model_validate(summarize_evaluation(task.model_dump()))
        task.report_environment = ReportEnvironment.model_validate(
            aggregate_project_environment([project.model_dump() for project in task.projects]))
        for field in task.overview:
            if field.key in REPORT_ENVIRONMENT_FIELDS:
                field.value = task.report_environment.values.get(field.key, "")
                field.confirmed = False
                field.candidates = []
            elif field.key in AUTOMATIC_OVERVIEW_FIELDS:
                field.value = task.evaluation_summary.values.get(field.key, "")
                field.confirmed = False
                field.candidates = []
        parsed = current_records(task) if parsed_records is None else parsed_records
        task = sync_record_fields(task, parsed)
        schema = {}
        outline = None
        if task.outline_id:
            info, path = upload_record(task.outline_id, "outline")
            outline = documents._load(path)
            schema = documents.outline_samples(outline)[0]
        task = Task.model_validate(documents.sync_test_basis(
            task.model_dump(), outline, info["name"] if task.outline_id else "当前大纲"))
        next(field for field in task.overview if field.key == "test_basis").old = "\n".join(
            documents._text(p) for p in documents._report_test_basis_paragraphs(master) if documents._text(p))
        if (task.sample_table_excluded_columns
                and task.sample_schema.get("headers") != schema.get("headers")):
            raise ValueError("大纲样机表表头已改变，请清除旧列选择并重新核对，不能套用原隐藏列。")
        documents.sample_table_columns(schema, task.sample_table_excluded_columns)
        task.production_basis_issues = []
        field = next((field for field in task.overview if field.key == "production_basis"), None)
        if field is not None:
            previous = [candidate.model_dump() for candidate in field.candidates]
            was_auto = not field.confirmed and not field.manual_override and {c.value for c in field.candidates} == {field.value}
            field.candidates = []
            if outline is not None:
                try:
                    field.candidates = [Candidate.model_validate(
                        documents.outline_production_basis(outline, info["name"]))]
                except ValueError as exc:
                    task.production_basis_issues.append(str(exc))
            old_value = field.value
            if not field.manual_override and (not field.value.strip() or was_auto):
                field.value = field.candidates[0].value if field.candidates else ""
            if previous != [c.model_dump() for c in field.candidates] or old_value != field.value:
                field.confirmed = False
        if task.report_type in INDEPENDENT_METER_TYPES:
            task = Task.model_validate(flowmeter_tasks.reconcile_samples(
                task.model_dump(), schema, parsed, strict=strict, report_type=task.report_type))
        elif task.outline_id or task.sample_schema or task.samples:
            from water_results import reconcile_task

            task = Task.model_validate(reconcile_task(task.model_dump(), schema, strict=strict, parsed=parsed))
        old_values = {}
        observation_issues = []
        task.observation_certificates = refresh_certificates(
            None, outline, task.observation_certificates, info["name"] if task.outline_id else "当前大纲")
        observation_parameters = outline_values(
            outline, info["name"] if task.outline_id else "当前大纲", report_type=task.report_type)
        if task.report_type in INDEPENDENT_METER_TYPES:
            refreshed = []
            for project in task.projects:
                result = (gas_records.apply_records(project.model_dump(), parsed, include_conditions=fill_flow_conditions)
                          if task.report_type == "燃气表" else flowmeter_records.apply_records(
                              project.model_dump(), parsed, outline, info["name"] if task.outline_id else "当前大纲"))
                if not fill_flow_conditions:
                    for key in ("start", "end", "temperature", "humidity"):
                        result[key] = getattr(project, key)
                refreshed.append(Project.model_validate(result))
            task.projects = refreshed
            task.report_environment = ReportEnvironment.model_validate(
                aggregate_project_environment([project.model_dump() for project in task.projects]))
            task.evaluation_summary = ReportEnvironment.model_validate(summarize_evaluation(task.model_dump()))
            for field in task.overview:
                if field.key in REPORT_ENVIRONMENT_FIELDS:
                    field.value = task.report_environment.values.get(field.key, "")
                elif field.key in AUTOMATIC_OVERVIEW_FIELDS:
                    field.value = task.evaluation_summary.values.get(field.key, "")
        else:
            # Gas result requirements keep their existing rules independently of observe copying.
            result_parameters = (outline_values(outline, info["name"] if task.outline_id else "当前大纲")
                                 if task.report_type == "燃气表" and task.projects else observation_parameters)
            if task.report_type == "水表":
                from water_results import sync_project

                task.projects = [Project.model_validate(sync_project(
                    library_root(), project.model_dump(), task.model_dump(), parsed, result_parameters,
                    refresh_conditions=fill_flow_conditions)) for project in task.projects]
            task.projects = [Project.model_validate(sync_project_requirements(library_root(), project.model_dump(), result_parameters))
                             for project in task.projects]
            task.projects = [Project.model_validate(refresh_curve_results(project.model_dump(), parsed))
                             for project in task.projects]
        for project in task.projects:
            pressure = summarize_pressure(project.model_dump(), parsed, task.report_type)
            project.pressure_auto_value = pressure["atmospheric_pressure"]
            if not project.pressure_manual_override:
                project.atmospheric_pressure = project.pressure_auto_value
            project.pressure_sources = pressure["pressure_sources"]
            project.pressure_issues = pressure["pressure_issues"]
        observe = None
        if task.observe_asset_id:
            try:
                observe_path = documents._asset(library_root(), task.observe_asset_id, "observe")["_path"]
                observe = documents._load(observe_path)
                task.observation_certificates = refresh_certificates(
                    observe, outline, task.observation_certificates,
                    info["name"] if task.outline_id else "当前大纲")
                if task.report_type == "水表":
                    old_values = manual_originals(observe)
            except ValueError as exc:
                if strict:
                    raise
                observe = None
                observation_issues.append("观察素材无法读取，请重新选择：" + str(exc))
                task.observation_certificates.issues.append("观察素材无法读取，请重新选择：" + str(exc))
        if task.items_asset_id:
            try:
                items = documents._load(documents._asset(library_root(), task.items_asset_id, "items")["_path"])
                certificate_issues = list(task.observation_certificates.issues)
                task.observation_certificates = refresh_certificates(
                    observe, outline, task.observation_certificates,
                    info["name"] if task.outline_id else "当前大纲", items=items)
                task.observation_certificates.issues.extend(certificate_issues)
            except ValueError as exc:
                if strict:
                    raise
                task.observation_certificates.issues.append("项目一览表证书备注无法读取，请重新选择：" + str(exc))
        for key, label in OBSERVATION_MANUAL_FIELDS.items():
            field = next((field for field in task.overview if field.key == key), None)
            if field is None:
                field = OverviewField(key=key, label=label)
                task.overview.append(field)
            old = field.old if observation_issues else old_values.get(key, "")
            if field.old != old:
                field.confirmed = False
            field.old = old
            field.candidates = []
        task.observation_summary = ObservationSummary.model_validate(
            observation_parameters)
        task.observation_summary.issues.extend(observation_issues)
        task = sync_photos(task, outline, strict)
        return Task.model_validate(documents.sync_other_information(
            task.model_dump(), outline, info["name"] if task.outline_id else "当前大纲"))

    def used_photo_ids(task):
        return {identity for slot, identity in task.photos.items()
                if task.report_type not in GAS_PHOTO_REPORT_TYPES or photo_slot_group(slot) not in GAS_UNUSED_PHOTO_SLOTS}

    def sync_photos(task, outline, strict):
        task.photo_components = []
        task.photo_issues = []
        if outline is not None:
            try:
                names = extract_component_names(documents._outline_regions(outline)[0])
                task.photo_components = [
                    {"id": f"component:{task.outline_id}:{index}", "name": name}
                    for index, name in enumerate(names, 1)
                ]
            except ValueError as exc:
                task.photo_issues.append("附件2零部件名称识别失败：" + str(exc))
        allowed = set(PHOTO_SLOTS) | {item["id"] for item in task.photo_components}
        if set(task.photo_layout) - photo_layout_ids(task.photos, task.photo_components, task.display_types, task.report_type):
            raise ValueError("照片排版设置与当前大纲不匹配，请移除旧零部件排版设置后重试")
        if set(task.photo_edits) - set(task.photos):
            raise ValueError("照片编辑设置没有对应照片，请移除旧编辑设置")
        maximum_width = (photo_width_limit(documents._load(library_root() / "report model/report.docx"))
                         if task.photo_edits else None)
        for slot, edit in task.photo_edits.items():
            if task.report_type in GAS_PHOTO_REPORT_TYPES and photo_slot_group(slot) in GAS_UNUSED_PHOTO_SLOTS:
                continue
            if edit.width_cm is not None and maximum_width is not None and edit.width_cm > maximum_width:
                raise ValueError("照片输出宽度超出附件2可用页面宽度（%.2fcm）" % maximum_width)
            with Image.open(photo_path(store, task.photos[slot])) as image:
                edited_photo_size(image.size, edit)
        if {photo_slot_group(slot) for slot in task.photos} - allowed:
            message = "零部件照片与当前大纲不匹配，请先移除旧零部件照片关联，再更换大纲并逐项上传。"
            if strict:
                raise ValueError(message)
            task.photo_issues.append(message)
        for identity in used_photo_ids(task):
            try:
                photo_path(store, identity)
            except (KeyError, ValueError) as exc:
                if strict:
                    raise
                task.photo_issues.append("照片无法读取，请移除关联后重新上传：" + str(exc))
        task.construction_instrument = ""
        task.construction_issues = []
        if outline is not None:
            names = {
                entry["value"].strip()
                for entry in documents._field_candidates(outline, "当前大纲").get("instrument_name", [])
                if entry["value"].strip()
            }
            if len(names) == 1:
                task.construction_instrument = next(iter(names))
        if task.construction_usage:
            field = next((field for field in task.overview if field.key == "construction"), None)
            try:
                text = construction_text(
                    task.construction_instrument,
                    [part["name"] for part in task.photo_components],
                    task.construction_usage, task.construction_custom,
                )
            except ValueError as exc:
                task.construction_issues.append(str(exc))
                text = ""
            if field is None:
                task.construction_issues.append("构造和应用场合字段缺失，请重新打开任务。")
            elif field.value != text:
                field.value = text
                field.confirmed = False
        return task

    def sync_report_type(task):
        def has_materials(value):
            return value.get("items_asset_id") or value.get("observe_asset_id") or value.get("projects")

        if task.id:
            previous = store.get("tasks", task.id)
            if previous.get("report_type", "水表") != task.report_type and has_materials(previous):
                raise ValueError("任务已选择素材或结果项目，不能切换报告对象类型；请新建任务。")
        for identity, kind in [(identity, "record") for identity in task.record_ids] + (
                [(task.outline_id, "outline")] if task.outline_id else []):
            info = store.get("uploads", identity)
            validate_upload_type(task, info, kind)
        for key, kind in (("items_asset_id", "items"), ("observe_asset_id", "observe")):
            identity = getattr(task, key)
            if identity:
                asset = documents._asset(library_root(), identity, kind)
                if asset["report_type"] != task.report_type:
                    raise ValueError(f"{task.report_type}任务不能使用{asset['report_type']}素材：{asset['name']}")
                setattr(task, key, asset["id"])
        for project in task.projects:
            asset = documents._asset(library_root(), project.asset_id, "result")
            from water_results import selectable

            if (project.water_result_labels or project.water_output_labels is not None) and (
                    task.report_type != "水表" or not selectable(project.model_dump())):
                raise ValueError("动态样机明细目前仅适用于水表单型号及B/O/E1/E2等级素材，不支持2 models。")
            if asset["report_type"] != task.report_type:
                raise ValueError(f"{task.report_type}任务不能使用{asset['report_type']}结果素材：{asset['name']}")
            project.asset_id = asset["id"]
            if task.report_type in INDEPENDENT_METER_TYPES:
                project.name = asset["project"]

    def validate_upload_type(task, info, kind):
        report_type = info.get("report_type")
        if ((report_type and report_type != task.report_type)
                or (not report_type and (task.report_type == "流量计（气体）"
                                        or (kind == "record" and task.report_type in FLOWMETER_TYPES)))):
            label = "大纲" if kind == "outline" else "原始记录"
            raise ValueError(f"{label}与当前报告对象不一致，水表、流量计（气体）、流量计（液体）和燃气表不能混用资料，请重新上传。")

    def apply_outline_fields(task, strict=True):
        fields = {}
        if task.outline_id:
            info, path = upload_record(task.outline_id, "outline")
            validate_upload_type(task, info, "outline")
            # Re-read the selected upload so older cached parses also get current mappings.
            parsed = documents.parse_upload(path, "outline")
            fields = json.loads(json.dumps(parsed["fields"], ensure_ascii=False).replace(path.name, info["name"]))
        for field in task.overview:
            if field.key not in OUTLINE_OVERVIEW_FIELDS or field.key in ("production_basis", "test_basis"):
                continue
            previous = {candidate.value for candidate in field.candidates}
            auto_value = not field.confirmed and not field.manual_override and previous == {field.value}
            field.candidates = [Candidate.model_validate(value) for value in fields.get(field.key, [])]
            unique = {candidate.value for candidate in field.candidates}
            if not field.confirmed and not field.manual_override and (not field.value or auto_value):
                field.value = next(iter(unique)) if len(unique) == 1 else ""
        return sync_samples(task, strict=strict)

    @app.get("/api/health")
    def health():
        return {"status": "ok", "local_only": server_access is None,
                **({"deployment": "protected_server"} if server_access else {})}

    @app.get("/api/catalog")
    def catalog():
        result = dict(documents.catalog(library_root()))
        warning = (
            "当前为试用版：完整20项目模拟样例已通过 Word 自动分页检查，实际报告仍须人工审版。"
            "生成时可选择直接输出Word或等待本机Word自动重排。直接输出需人工核对分页及更新页码；自动重排失败不发布未完成的新版本。"
        )
        result["warnings"] = [*result.get("warnings", []), warning]
        result["photo_slots"] = [
            {"id": slot, "name": name, "width_cm": width}
            for slot, (name, width) in PHOTO_SLOTS.items()
        ]
        result["photo_editing"] = True
        result["water_sample_selection"] = True
        result["gas_photo_rules"] = 2
        result["photo_max_width_cm"] = photo_width_limit(
            documents._load(library_root() / "report model/report.docx"))
        result["result_material_version"] = documents.result_material_version(library_root())
        result["result_material_batch_limit"] = 24
        return result

    @app.post("/api/assets/projects")
    def result_materials(request: ResultMaterialRequest):
        root = library_root()
        version = documents.result_material_version(root)
        if request.version != version:
            raise ConflictError("结果素材版本已变化，请刷新素材目录后重试；当前任务未改变。")
        projects, errors = [], []
        with documents.document_read_session():
            for identity in request.asset_ids:
                try:
                    projects.append(documents.project_from_asset(root, identity))
                except ValueError as exc:
                    errors.append({"asset_id": identity, "detail": str(exc)})
        if documents.result_material_version(root) != version:
            raise ConflictError("读取期间结果素材版本发生变化，请刷新素材目录后重试。")
        return {"version": version, "projects": projects, "errors": errors}

    @app.get("/api/assets/{identity}/project")
    def project(identity: str):
        return documents.project_from_asset(library_root(), identity)

    @app.get("/api/assets/{identity}/rows")
    def rows(identity: str):
        return documents.section_rows(library_root(), identity)

    @app.get("/api/equipment")
    def list_equipment():
        return store.list("equipment")

    @app.get("/api/principles")
    def list_principles():
        return store.list("principles")

    def material_scheme_assets(scheme):
        assets = [documents._asset(library_root(), identity, "result") for identity in scheme.asset_ids]
        if any(asset["report_type"] != scheme.report_type for asset in assets):
            raise ValueError("素材方案只能包含对应报告对象的结果素材，不能混用水表、流量计和燃气表。")
        if len({asset["id"] for asset in assets}) != len(assets):
            raise ValueError("素材方案包含指向同一文件的重复素材，请重新核对。")
        return assets

    @app.get("/api/material-schemes")
    def list_material_schemes():
        return store.list("material_schemes")

    @app.post("/api/material-schemes")
    def save_material_scheme(scheme: ResultMaterialScheme):
        assets = material_scheme_assets(scheme)
        if scheme.id:
            previous = store.get("material_schemes", scheme.id)
            if previous["report_type"] != scheme.report_type:
                raise ValueError("不能更改已有素材方案的报告对象，请保存为新方案。")
        if any(item["id"] != scheme.id and item["report_type"] == scheme.report_type
               and item["name"] == scheme.name for item in store.list("material_schemes")):
            raise ValueError("当前报告对象已有同名素材方案，请更新原方案或使用新名称。")
        scheme.asset_ids = [asset["id"] for asset in assets]
        return store.save("material_schemes", scheme.model_dump())

    @app.post("/api/material-schemes/{identity}/apply")
    def apply_material_scheme(identity: str, task: Task, revision: int = Query(..., ge=1)):
        scheme = ResultMaterialScheme.model_validate(store.get("material_schemes", identity))
        if scheme.revision != revision:
            raise ConflictError("素材方案已在其他窗口更新，请刷新方案库后重新选择应用。")
        if task.report_type != scheme.report_type:
            raise ValueError("素材方案与当前报告对象不一致，不能跨对象应用。")
        if not task.items_asset_id:
            raise ValueError("请先选择项目一览表，再应用素材方案。")
        assets = material_scheme_assets(scheme)
        aliases = {identity: asset["id"] for asset in assets
                   for identity in (asset["id"], asset["_path_id"], asset["_legacy_path_id"])}
        existing = {}
        for project in task.projects:
            canonical = aliases.get(project.asset_id)
            if canonical:
                if canonical in existing:
                    raise ValueError("当前任务重复选择同一结果素材，不能丢弃其中的编辑数据，请先核对。")
                project.asset_id = canonical
                existing[canonical] = project
        task.projects = [
            existing[asset["id"]] if asset["id"] in existing else Project.model_validate(
                documents.project_from_asset(library_root(), asset["id"]))
            for asset in assets
        ]
        return order_task(task)

    def scheme_assets(scheme):
        from water_results import selectable

        assets = []
        for entry in scheme.projects:
            asset = documents._asset(library_root(), entry.asset_id, "result")
            if asset["report_type"] != "水表" or not selectable(asset):
                raise ValueError("样机方案目前仅支持水表单型号及B/O/E1/E2等级素材，不支持2 models：" + asset["name"])
            assets.append(asset)
        if len({asset["id"] for asset in assets}) != len(assets):
            raise ValueError("样机方案存在重复结果素材，请重新核对。")
        return assets

    @app.get("/api/sample-schemes")
    def list_sample_schemes():
        return store.list("sample_schemes")

    @app.post("/api/sample-schemes")
    def save_sample_scheme(scheme: SampleOutputScheme):
        assets = scheme_assets(scheme)
        if scheme.id:
            store.get("sample_schemes", scheme.id)
        if any(item["id"] != scheme.id and item["name"] == scheme.name
               for item in store.list("sample_schemes")):
            raise ValueError("已有同名样机方案，请选择原方案更新，或使用新名称。")
        for entry, asset in zip(scheme.projects, assets):
            entry.asset_id = asset["id"]
        return store.save("sample_schemes", scheme.model_dump())

    @app.post("/api/sample-schemes/{identity}/apply")
    def apply_sample_scheme(identity: str, task: Task):
        if task.report_type != "水表":
            raise ValueError("样机方案目前仅适用于水表。")
        if not task.items_asset_id:
            raise ValueError("请先选择项目一览表，再应用样机方案。")
        task = order_task(task)
        scheme = SampleOutputScheme.model_validate(store.get("sample_schemes", identity))
        assets = scheme_assets(scheme)
        available = {sample.label for sample in task.samples if not sample.disassembly}
        existing = {project.asset_id: project for project in task.projects}
        issues = []
        for entry, asset in zip(scheme.projects, assets):
            project = existing.get(asset["id"])
            if project is None:
                project = Project.model_validate(documents.project_from_asset(library_root(), asset["id"]))
                if task.projects:
                    ordinary = {"吴锦川", "姚新红"}
                    project.evaluators = [person for person in task.projects[0].evaluators if person in ordinary]
                    project.reviewers = [person for person in task.projects[0].reviewers if person in ordinary]
                task.projects.append(project)
            missing = [label for label in entry.labels if label not in available]
            if missing:
                issues.append(project.name + "：以下方案代号不存在或已选拆样，未自动补造样机：" +
                              "、".join(missing))
            project.water_output_labels = [label for label in entry.labels if label in available]
        return {"task": order_task(task), "issues": issues}

    @app.post("/api/equipment")
    def save_equipment(item: Equipment):
        from equipment_typography import equipment_identity, normalize_equipment_typography

        item = Equipment.model_validate(normalize_equipment_typography(item.trimmed().model_dump()))
        if not all((item.name, item.model, item.performance, item.number, item.valid_until,
                    item.storage_location, item.trial_group)):
            raise ValueError("请完整填写设备七项信息；不适用项请明确填写不适用。")
        parse_date(item.valid_until)
        if any(old["id"] != item.id and equipment_identity(
            *(old.get(key, "") for key in ("name", "model", "number"))
        ) == equipment_identity(item.name, item.model, item.number) for old in store.list("equipment")):
            raise ValueError("名称、型号和编号均相同的设备已存在，请修改原设备，避免重复录入。")
        return store.save("equipment", item.model_dump())

    @app.get("/api/tasks")
    def list_tasks():
        return [
            {**{key: item[key] for key in ("id", "name", "updated_at", "revision")},
             "report_type": item.get("report_type", "水表")}
            for item in store.list("tasks")
        ]

    @app.get("/api/tasks/{identity}")
    def get_task(identity: str):
        task = Task.model_validate(store.get("tasks", identity))
        known = {field.key for field in task.overview}
        for field in documents.catalog(library_root())["overview"]:
            if field["key"] not in known:
                task.overview.append(OverviewField.model_validate(field))
        return apply_outline_fields(task, strict=False)

    @app.delete("/api/tasks/{identity}")
    def delete_task(identity: str, revision: int = Query(..., ge=1)):
        return store.delete_task(identity, revision)

    @app.post("/api/tasks")
    def save_task(task: Task):
        task = order_task(task)
        return store.save("tasks", task.model_dump())

    @app.post("/api/tasks/order")
    def order(task: Task):
        return order_task(task)

    @app.post("/api/tasks/outline-fields")
    def outline_fields(task: Task):
        return apply_outline_fields(task)

    @app.post("/api/tasks/record-fields")
    def record_fields(task: Task):
        return sync_samples(task)

    @app.post("/api/environment-summary")
    def environment_summary(task: Task):
        return aggregate_project_environment([project.model_dump() for project in task.projects])

    @app.post("/api/evaluation-summary")
    def evaluation_summary(task: Task):
        return summarize_evaluation(task.model_dump())

    @app.post("/api/other-information")
    def other_information(task: Task):
        outline, source = None, "当前大纲"
        if task.outline_id:
            info, path = upload_record(task.outline_id, "outline")
            outline, source = documents._load(path), info["name"]
        current = documents.sync_other_information(task.model_dump(), outline, source)
        return {key: current[key] for key in (
            "overview", "sealing_model", "sealing_model_source", "sealing_issues",
            "supplementary_notes", "supplementary_note_ids", "other_information_order",
            "supplementary_source", "supplementary_issues", "other_sealing_number")}

    @app.post("/api/gas-curve-previews/{project_id}")
    def gas_curve_preview(project_id: str, task: Task):
        from gas_curves import CURVE_HEADINGS, project_curve_images

        if task.report_type != "燃气表":
            raise ValueError("此曲线预览仅适用于燃气表的误差曲线和耐久性项目。")
        selected = [project for project in task.projects if project.id == project_id]
        if len(selected) != 1 or gas_records.project_kind(selected[0].model_dump()) not in CURVE_HEADINGS:
            raise ValueError("未找到唯一对应的误差曲线或耐久性项目，请重新打开实测结果。")
        for project in task.projects:
            project.gas_curves = []
        task = sync_samples(task)
        projects = [project.model_dump() for project in task.projects]
        current = next(project for project in projects if project["id"] == project_id)
        charts = project_curve_images(current, projects)
        if not charts:
            raise ValueError("本项目没有可预览的样机结果行，请先核对样机对应关系。")
        for chart in charts:
            image = chart.pop("image")
            chart["image_url"] = ("data:image/png;base64," + base64.b64encode(image.getvalue()).decode("ascii")
                                  if image is not None else "")
        return {"project_id": project_id, "kind": gas_records.project_kind(current), "charts": charts}

    @app.get("/api/photos")
    def photos():
        return store.list("photos")

    @app.post("/api/photo-previews")
    def begin_photo_preview(task: Task, background_tasks: BackgroundTasks):
        outline = documents._load(upload_record(task.outline_id, "outline")[1]) if task.outline_id else None
        task = sync_photos(task, outline, True)
        if task.photo_issues:
            raise ValueError("；".join(task.photo_issues))
        if task.display_types is None:
            from basic_options import template_options
            task.display_types = template_options(documents._load(library_root() / "report model/report.docx"))["display_types"]
        paths = {identity: photo_path(store, identity) for identity in used_photo_ids(task)}
        result = photo_previews.begin(task.model_dump(), paths)
        background_tasks.add_task(photo_previews.execute, result["id"])
        return result

    @app.get("/api/photo-previews/{identity}")
    def get_photo_preview(identity: str):
        photo_previews.folder(identity)
        return store.get("photo_previews", identity)

    @app.get("/api/photo-previews/{identity}/{filename}")
    def photo_preview_file(identity: str, filename: str, options: DownloadOptions = Depends(download_options)):
        try:
            path = photo_previews.file(identity, filename)
        except FileNotFoundError as exc:
            raise HTTPException(404, str(exc)) from exc
        return download_response(path, options)

    @app.post("/api/photos")
    async def upload_photo(file: UploadFile = File(...), photo_mode: Literal["standard", "detail"] = Form("standard")):
        try:
            content = await file.read(MAX_PHOTO_BYTES + 1)
            return await run_in_threadpool(save_photo, store, content, file.filename,
                                           cloud_mode=photo_mode if cloud_photo_optimization else None)
        finally:
            await file.close()

    @app.get("/api/photos/{identity}/image")
    def photo_image(identity: str, rotation: int = 0, crop: str = "",
                    encoding: Literal["binary", "base64"] = "binary",
                    preview: Literal["full", "thumbnail"] = "full"):
        path = photo_path(store, identity)
        if rotation not in (0, 90, 180, 270):
            raise ValueError("照片旋转仅支持0、90、180、270度")
        try:
            box = [int(value) for value in crop.split(",")] if crop else None
            edit = PhotoEdit(rotation=rotation, crop=box)
        except ValueError as exc:
            raise ValueError("照片裁剪参数无效：" + str(exc)) from exc
        media_type = "image/jpeg" if path.suffix == ".jpg" else "image/png"
        thumbnail = preview == "thumbnail"
        if encoding == "base64":
            content = (edited_photo_bytes(path, edit, thumbnail=thumbnail)
                       if edit.rotation or edit.crop is not None or thumbnail else path.read_bytes())
            return JSONResponse({
                "media_type": media_type,
                "data_base64": base64.b64encode(content).decode("ascii"),
            }, headers={"Cache-Control": "no-store"})
        if not edit.rotation and edit.crop is None and not thumbnail:
            return FileResponse(path, media_type=media_type)
        return Response(edited_photo_bytes(path, edit, thumbnail=thumbnail), media_type=media_type)

    @app.get("/api/uploads")
    def uploads():
        items = [classified_upload(item) for item in store.list("uploads")]
        for item in items:
            if item["kind"] == "record":
                item["warnings"] = [
                    warning.replace("证书编号不自动等同报告编号；",
                                    "报告编号读取原始记录报告编号栏或样品信息首行右侧证书编号栏；")
                    for warning in item.get("warnings", [])
                ]
        return items

    @app.get("/api/workspace")
    def workspace():
        loaders = {
            "catalog": catalog, "equipment": list_equipment, "uploads": uploads,
            "photos": photos, "principles": list_principles,
            "sample_schemes": list_sample_schemes, "material_schemes": list_material_schemes,
        }
        result = {}
        with documents.document_read_session():
            for name, loader in loaders.items():
                try:
                    result[name] = {"data": loader()}
                except (ValueError, KeyError) as exc:
                    LOGGER.warning("Workspace section %s failed: %s", name, exc)
                    result[name] = {"error": str(exc) if isinstance(exc, ValueError) else "记录不存在，请刷新后重试。"}
                except Exception:
                    LOGGER.exception("Workspace section %s failed", name)
                    result[name] = {"error": "处理失败，未保存为成功结果。请查看终端错误记录。"}
        return result

    @app.delete("/api/uploads/{identity}")
    def delete_upload(identity: str):
        return store.delete_upload(identity)

    @app.get("/api/uploads/{identity}/file")
    def download_upload(identity: str, options: DownloadOptions = Depends(download_options)):
        info, path = upload_record(identity)
        return download_response(path, options, filename=info["name"])

    @app.post("/api/uploads")
    async def upload(file: UploadFile = File(...), kind: str = Form(...),
                     report_type: Literal["水表", "流量计（气体）", "流量计（液体）", "燃气表", "流量计"] = Form("水表")):
        report_type = normalize_report_type(report_type)
        if kind not in ("record", "outline"):
            raise ValueError("请选择原始记录或大纲文件类型。")
        extension = Path(file.filename or "").suffix.lower()
        if extension != ".docx" and not (extension == ".xlsx" and kind == "record"
                                        and report_type in ("燃气表", "流量计（气体）")):
            raise ValueError("大纲、水表及流量计（液体）记录仅支持DOCX；燃气表和流量计（气体）原始记录另支持各自格式的无宏XLSX。")
        identity = uuid.uuid4().hex
        path = store.directory / "uploads" / (identity + extension)
        complete = False
        try:
            size = 0
            with path.open("xb") as target:
                while True:
                    chunk = await file.read(1024 * 1024)
                    if not chunk:
                        break
                    size += len(chunk)
                    if size > MAX_UPLOAD:
                        raise ValueError("单个文档最大支持 32 MB。")
                    target.write(chunk)
            if extension == ".xlsx":
                excel_parser = gas_flowmeter_excel if report_type == "流量计（气体）" else gas_records
                parsed = await run_in_threadpool(excel_parser.parse_upload, path)
            elif kind == "record" and report_type in FLOWMETER_TYPES:
                check_docx(path)
                parsed = await run_in_threadpool(flowmeter_records.parse_upload, path)
            else:
                check_docx(path)
                parsed = await run_in_threadpool(documents.parse_upload, path, kind)
            parsed["report_type"] = report_type
            name = Path(file.filename).name
            # Internal filenames are opaque; sources shown to people use the original name.
            parsed = json.loads(json.dumps(parsed, ensure_ascii=False).replace(path.name, name))
            info = dict(
                parsed, id=identity, name=name, kind=kind, extension=extension,
                sha256=hashlib.sha256(path.read_bytes()).hexdigest(),
            )
            saved = store.save("uploads", info)
            complete = True
            return saved
        finally:
            await file.close()
            if not complete:
                path.unlink(missing_ok=True)

    @app.post("/api/tasks/parse")
    def parse_task(task: Task, refresh_records: bool = Query(False)):
        parsed = current_records(task, refresh_records=refresh_records)
        task = sync_samples(task, parsed_records=parsed, fill_flow_conditions=True)
        all_sources = list(parsed)
        if task.outline_id:
            all_sources.append(upload_record(task.outline_id, "outline")[0])
        for field in task.overview:
            if field.key in OUTLINE_OVERVIEW_FIELDS | REPORT_ENVIRONMENT_FIELDS | AUTOMATIC_OVERVIEW_FIELDS | OBSERVATION_MANUAL_FIELDS.keys() or field.key in ("report_number", "sealing"):
                continue
            existing = {(candidate.value, candidate.source) for candidate in field.candidates}
            for source in all_sources:
                for candidate in source.get("fields", {}).get(field.key, []):
                    pair = (candidate["value"], candidate["source"])
                    if pair not in existing:
                        field.candidates.append(Candidate.model_validate(candidate))
                        existing.add(pair)
            unique = {candidate.value for candidate in field.candidates}
            if not field.confirmed and not field.manual_override and not field.value and len(unique) == 1:
                field.value = next(iter(unique))
        if task.report_type == "水表":
            task.projects = [
                Project.model_validate(refresh_curve_results(project.model_dump(), parsed,
                    include_conditions=not bool(project.water_result_labels)))
                for project in task.projects
            ]
        return order_task(apply_outline_fields(task))

    @app.post("/api/tasks/recovery-check")
    def check_recovery(request: RecoveryRequest):
        return request

    @app.post("/api/validate")
    def validate(task: Task):
        task = sync_samples(task)
        validate_assets(task)
        component_text = upload_record(task.outline_id, "outline")[0].get("outline", {}).get("components", "") if task.outline_id else ""
        result = validate_task(task, component_text)
        if task.outline_id:
            info, _ = upload_record(task.outline_id, "outline")
            outline = info.get("outline", {})
            if any(not outline.get(key) for key in ("components", "coverage")):
                extra = IssueMessages(result["issues"], "error")
                extra.target = {"step": 0}
                extra.append("本次大纲未完整定位两处复制区域，请检查大纲章节标题和结构。")
                result["errors"].extend(extra)
        for identity in task.record_ids:
            upload_record(identity, "record")
        return result

    @app.post("/api/generate")
    def generate(request: GenerateRequest):
        return generate_snapshot(request)

    @app.post("/api/generation-jobs")
    def begin_generation(request: GenerateRequest):
        if request.mode not in ("auto", "draft", "review", "final"):
            raise ValueError("未知的输出模式。")
        task = order_task(request.task)
        validation = validate(task)
        if request.mode not in ("auto", "draft") and validation["errors"]:
            raise HTTPException(422, validation["errors"])
        saved = Task.model_validate(store.save("tasks", task.model_dump()))
        request = request.model_copy(update={"task": saved})
        from record_documents import TEMPLATE_PATH
        sources = {library_root() / "report model/report.docx", library_root() / TEMPLATE_PATH}
        if task.projects:
            from signature_templates import signature_template_path

            sources.add(signature_template_path(library_root(), task.include_atmospheric_pressure))
        sources.update(documents._asset(library_root(), identity)["_path"] for identity in
                       [task.items_asset_id, task.observe_asset_id, *[p.asset_id for p in task.projects]] if identity)
        sources.update(upload_record(identity)[1] for identity in [task.outline_id, *task.record_ids] if identity)
        sources.update(photo_path(store, identity) for identity in used_photo_ids(task))
        hashes = {str(path): hashlib.sha256(path.read_bytes()).hexdigest() for path in sources}
        job = generation_jobs.submit(request, generate_snapshot, hashes)
        return {**job, "saved_task": saved.model_dump()}

    @app.get("/api/generation-jobs")
    def list_generation_jobs():
        return generation_jobs.list_public()

    @app.get("/api/generation-jobs/{identity}")
    def get_generation_job(identity: str):
        return generation_jobs.public(store.get("generation_jobs", identity))

    def generate_snapshot(request: GenerateRequest, *, snapshot=False, progress=None, source_hashes=None):
        def stage(text):
            if progress:
                progress(text)

        source_hashes = dict(source_hashes or {})
        if request.task.projects:
            from signature_templates import signature_template_path

            path = signature_template_path(library_root(), request.task.include_atmospheric_pressure)
            source_hashes.setdefault(str(path), hashlib.sha256(path.read_bytes()).hexdigest())

        def check_sources():
            for name, expected in (source_hashes or {}).items():
                path = Path(name)
                if not path.is_file() or hashlib.sha256(path.read_bytes()).hexdigest() != expected:
                    raise ValueError("生成快照的源资料已改变或丢失，请重新提交：" + path.name)

        check_sources()
        started = time.monotonic()
        timings = {}
        if request.mode not in ("auto", "draft", "review", "final"):
            raise ValueError("未知的输出模式。")
        task = request.task if snapshot else order_task(request.task)
        if task.photo_issues:
            raise ValueError("；".join(task.photo_issues))
        if task.construction_issues:
            raise ValueError("；".join(task.construction_issues))
        component_text = upload_record(task.outline_id, "outline")[0].get("outline", {}).get("components", "") if snapshot and task.outline_id else ""
        validation = validate_task(task, component_text) if snapshot else validate(task)
        draft = request.mode == "draft" or (request.mode == "auto" and bool(validation["errors"]))
        if not draft and validation["errors"]:
            raise HTTPException(status_code=422, detail=validation["errors"])
        outline_path = upload_record(task.outline_id, "outline")[1] if task.outline_id else None
        if not snapshot and task.id and store.get("tasks", task.id)["revision"] != task.revision:
            raise ConflictError("任务已更新，请重新打开后再生成。")
        identity = uuid.uuid4().hex
        folder = store.directory / "exports" / identity
        folder.mkdir()
        filename = ("draft" if draft else "report") + "-" + identity + ".docx"
        output = folder / filename
        report_task = task.model_dump()
        report_task["_photo_paths"] = {
            identity: photo_path(store, identity) for identity in used_photo_ids(task)
        }
        stage_started = time.monotonic()
        timings["资料校验及准备"] = stage_started - started
        stage("正在合成型评报告")
        composition_options = {}
        if request.pagination_strategy == "optimized_v4":
            from optimized_observation import rebuild

            composition_options["observation_builder"] = rebuild
        result = documents.compose_report(library_root(), report_task, outline_path, output, draft, **composition_options)
        check_docx(output)
        timings["型评报告合成"] = time.monotonic() - stage_started
        stage_started = time.monotonic()
        record_filename = "original-record-" + identity + ".docx"
        record_output = folder / record_filename
        uploads = [upload_record(identity, "record") for identity in dict.fromkeys(task.record_ids)]
        record_paths = [path for _, path in uploads]
        pagination = {}
        stage("正在合成配套记录及Excel打印页")
        record_result = compose_original_record(
            library_root(), report_task, outline_path, output, record_output, record_paths, **composition_options)
        from negative_numbers import protect_negative_numbers

        for path in (output, record_output):
            protect_negative_numbers(path)
        prepare_metadata_groups(output)
        prepare_metadata_groups(record_output)
        timings["配套记录合成及页底约束"] = time.monotonic() - stage_started
        if request.native_word:
            from optimized_v4 import finalize_v4_pages

            finalize = finalize_v4_pages
            try:
                for key, label, path in (("report", "型评报告重排", output), ("record", "配套记录重排", record_output)):
                    stage_started = time.monotonic()
                    stage(label + "：等待Word处理")
                    pagination[key] = (finalize(path, progress=lambda text: stage(label + "：" + text))
                                       if progress else finalize(path))
                    timings[label] = time.monotonic() - stage_started
            except ValueError as exc:
                raise ValueError("自动分页未完成，本次未发布新版本：" + str(exc)) from exc
        stage_started = time.monotonic()
        stage("正在检查两份文件及比较历史版本")
        check_docx(output)
        check_docx(record_output)
        warnings = (validation["warnings"] + validation["errors"] + result.get("warnings", []) +
                    record_result["warnings"])
        if request.native_word:
            warnings = [warning for warning in warnings if warning != "Word 页码域已标记更新；需本机 Word 重分页核对。"]
        warnings.extend(warning for value in pagination.values() for warning in value["warnings"])
        if any(path.suffix == ".xlsx" for _, path in uploads):
            warnings.append("Excel“打印页”按报告项目及编号筛选、排序后追加到配套原始记录末尾，共用页只追加一次；无对应本次实测结果的页不追加。生成结果不再单独输出Excel，上传原件保留，其他计算工作表不追加。")
        warnings.append(
            "两份Word已完成结果续表自动分页及页码更新，仍须人工核对整份报告，不代表正式审版或签发。"
            if request.native_word else
            "本版为直接生成Word，未调用本机Word自动重排；请打开两份文件更新页码，并人工核对跨页表头、续页说明及页底信息，不代表正式审版或签发。"
        )
        check_sources()
        saved = task.model_dump() if snapshot else store.save("tasks", task.model_dump())
        manifest = {
            "id": identity, "mode": "draft" if draft else "review", "task": saved,
            "native_word": request.native_word,
            "pagination_strategy": request.pagination_strategy,
            "warnings": list(dict.fromkeys(warnings)),
            "issues": validation.get("issues", []),
            "layout": {"status": "pending", "stage": "not_started", "error": ""},
            "sha256": hashlib.sha256(output.read_bytes()).hexdigest(),
            "url": f"/api/exports/{identity}/{filename}",
            "record_url": f"/api/exports/{identity}/{record_filename}",
            "record_sha256": hashlib.sha256(record_output.read_bytes()).hexdigest(),
            "record_cover": record_result["cover"],
            "excel_print_sheets": record_result["excel_print_sheets"],
            "excel_print_order": record_result["excel_print_order"],
            "result_pagination": pagination,
            "pdf_url": "", "checked_url": "",
            "revision": saved["revision"], "task_id": saved["id"],
            "excel_records": [],
        }
        manifest["comparison"] = compare_exports(store, manifest, previous_export(manifest))
        timings["完整性检查及版本比较"] = time.monotonic() - stage_started
        manifest["generation_timing"] = {
            "total_seconds": round(time.monotonic() - started, 3),
            "stage_seconds": {key: round(value, 3) for key, value in timings.items()},
        }
        return public_export(save_manifest(store, manifest), store)

    def previous_export(manifest):
        candidates = [
            item for item in store.list("exports")
            if item["task_id"] == manifest["task_id"] and item["revision"] < manifest["revision"]
        ]
        return max(candidates, key=lambda item: item["revision"]) if candidates else None

    @app.get("/api/tasks/{identity}/exports")
    def task_exports(identity: str):
        store.get("tasks", identity)
        return [
            public_export(item, store) for item in sorted(
                store.list("exports"), key=lambda item: item["revision"], reverse=True
            ) if item["task_id"] == identity
        ]

    @app.get("/api/tasks/{identity}/workspace")
    def task_workspace(identity: str):
        with documents.document_read_session():
            task = get_task(identity)
            previous_order = [project.id for project in task.projects]
            if task.items_asset_id:
                task = order_task(task)
            return {
                "task": task, "exports": task_exports(identity), "photos": photos(),
                "order_verified": bool(task.items_asset_id),
                "order_changed": previous_order != [project.id for project in task.projects],
            }

    @app.get("/api/exports/{identity}/comparison")
    def export_comparison(identity: str):
        manifest = store.get("exports", identity)
        if "comparison" not in manifest or manifest["comparison"]["status"] == "unavailable":
            manifest["comparison"] = compare_exports(store, manifest, previous_export(manifest))
            manifest = save_manifest(store, manifest)
        return manifest["comparison"]

    @app.get("/api/exports/{identity}/records/{filename}")
    def export_excel_record(identity: str, filename: str, options: DownloadOptions = Depends(download_options)):
        manifest = store.get("exports", identity)
        match = next((item for item in manifest.get("excel_records", [])
                      if item["url"].rsplit("/", 1)[-1] == filename), None)
        if not match or not re.fullmatch(r"[0-9a-f]{32}\.xlsx", filename):
            raise HTTPException(404, "本版报告没有该Excel原始记录。")
        path = store.directory / "exports" / identity / filename
        if not path.is_file():
            raise HTTPException(404, "本版Excel原始记录文件丢失。")
        if hashlib.sha256(path.read_bytes()).hexdigest() != match["sha256"]:
            raise ValueError("本版Excel原件内容被外部修改，不能作为保存时的原件下载。")
        return download_response(path, options, filename=match["name"])

    @app.get("/api/exports/{identity}")
    def export_status(identity: str):
        return public_export(store.get("exports", identity), store)

    @app.get("/api/retained-reports")
    def retained_reports():
        current = {task["id"] for task in store.list("tasks")}
        return [
            {**{key: item.get(key, "") for key in
                ("id", "task_id", "revision", "url", "record_url", "pdf_url", "checked_url", "updated_at")},
             "excel_records": item.get("excel_records", []),
             "generated_at": store.first_saved_at("exports", item["id"]),
             "task_name": item.get("task", {}).get("name", "未命名任务"),
             "report_type": item.get("task", {}).get("report_type", "水表")}
            for item in store.list("exports") if item["task_id"] not in current
        ]

    @app.post("/api/exports/{identity}/layout-check")
    def layout_check(identity: str, background_tasks: BackgroundTasks):
        manifest = layout_checks.begin(identity)
        background_tasks.add_task(layout_checks.execute, identity)
        return public_export(manifest, store)

    @app.get("/api/exports/{identity}/checks/{attempt}/{filename}")
    def checked_export(identity: str, attempt: str, filename: str, options: DownloadOptions = Depends(download_options)):
        manifest = store.get("exports", identity)
        if (not re.fullmatch(r"[a-f0-9]{32}", attempt)
                or manifest.get("layout", {}).get("attempt_id") != attempt
                or manifest["layout"]["status"] != "rendered"
                or filename not in (f"word-{attempt}.docx", "preview.pdf")):
            raise HTTPException(404, "没有已完成的排版检查文件。")
        path = store.directory / "exports" / identity / "checks" / attempt / filename
        if not path.is_file():
            raise HTTPException(404, "文件不存在。")
        return download_response(path, options, filename=("已分页-仍需人工核对.docx" if filename.endswith(".docx") else filename))

    @app.get("/api/exports/{identity}/{filename}")
    def export(identity: str, filename: str, options: DownloadOptions = Depends(download_options)):
        manifest = store.get("exports", identity)
        if not re.fullmatch(r"[a-f0-9]{32}", identity) or filename not in (
            f"draft-{identity}.docx", f"report-{identity}.docx", f"original-record-{identity}.docx",
            "preview.pdf", "manifest.json",
        ):
            raise HTTPException(404, "文件不存在。")
        path = store.directory / "exports" / identity / filename
        if not path.is_file():
            raise HTTPException(404, "文件不存在。")
        if filename.endswith(".docx"):
            report_name = re.sub(r"[/\\\x00-\x1f]", "_", manifest["task"]["name"]).strip()[:100]
            if filename.startswith("original-record-"):
                display_name = (report_name or "型式评价") + "-原始记录-待填写.docx"
            else:
                display_name = (report_name or "型式评价报告") + (
                    "-待补录草稿" if manifest["mode"] == "draft" else "-排版待核验"
                ) + ".docx"
        else:
            display_name = filename
        return download_response(path, options, filename=display_name)

    @app.get("/")
    def home():
        return FileResponse(BASE / "static" / "index.html")

    from separation import routes as separation_routes
    app.include_router(separation_routes(store.directory, check_docx))
    from separation_experimental import routes as experimental_separation_routes
    app.include_router(experimental_separation_routes(store.directory, check_docx))

    @app.get("/separate-experimental")
    def experimental_separation_page():
        return FileResponse(BASE / "static" / "separate-experimental.html")

    @app.get("/separate")
    def separation_page():
        return FileResponse(BASE / "static" / "separate.html")

    app.mount("/static", StaticFiles(directory=BASE / "static", check_dir=False), name="static")
    return app
