"""Task-local appendix previews using the same photo composition and native Word renderer."""

import logging
import math
import re
import sqlite3
import threading
import uuid

import pymupdf
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

import documents
from native_word import render
from photo_documents import _appendix_region, fill_attachment_photos

LOGGER = logging.getLogger("report_workbench.photo_preview")


def compose_preview(root, task, paths, output, layout_targets=None):
    doc = documents._load(root / "report model/report.docx")
    _appendix_region(doc)
    section = doc.sections[-1]
    if section.header.is_linked_to_previous or section.footer.is_linked_to_previous:
        raise ValueError("附件2预览需要独立页眉页脚，当前模板不能安全拆出预览。")
    number = next((item["value"] for item in task["overview"] if item["key"] == "report_number"), "")
    for part in (section.header, section.footer):
        for paragraph in part._element.xpath(".//w:p"):
            documents._replace_text_nodes(paragraph, number)
    documents._validate_attachment_headers(doc)
    body = doc.element.body
    start = documents._unique_anchor(body, "样机的整体外形照片")
    for element in list(body)[:body.index(start)]:
        body.remove(element)
    targets = []
    warnings = fill_attachment_photos(
        doc, task["photos"], task["photo_components"], paths, task["photo_edits"],
        display_types=task["display_types"], photo_order=task["photo_order"],
        photo_layout=task.get("photo_layout", {}), layout_targets=targets,
        report_type=task.get("report_type", "水表"),
    )
    if layout_targets is not None:
        bookmark_id = max([int(value) for value in doc.element.xpath(".//w:bookmarkStart/@w:id")
                           if value.isdigit()] + [0]) + 1
        prefix = "photo_" + uuid.uuid4().hex[:16] + "_"
        for index, target in enumerate(targets):
            paragraph = target.pop("paragraph")
            name = prefix + str(index)
            begin, end = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
            begin.set(qn("w:id"), str(bookmark_id + index))
            begin.set(qn("w:name"), name)
            end.set(qn("w:id"), str(bookmark_id + index))
            paragraph._p.insert(1, begin)
            paragraph._p.append(end)
            target["bookmark"] = name
            layout_targets.append(target)
    page_number = body.sectPr.find(qn("w:pgNumType"))
    if page_number is not None:
        body.sectPr.remove(page_number)
    update = doc.settings.element.find(qn("w:updateFields"))
    if update is not None:
        doc.settings.element.remove(update)
    documents._editable_page_breaks(doc)
    doc.save(output)
    documents._normalize_report_fonts(output)
    return warnings


class PhotoPreviews:
    def __init__(self, store, root):
        self.store, self.root = store, root
        self.lock = threading.Lock()
        for item in store.list("photo_previews"):
            if item["status"] == "running":
                item.update(status="failed", error="服务已重启，照片排版预览中断，请重新生成预览。")
                store.save("photo_previews", item)

    def folder(self, identity):
        if not re.fullmatch(r"[a-f0-9]{32}", identity):
            raise ValueError("无效的照片预览标识")
        return self.store.directory / "photo-previews" / identity

    def begin(self, task, paths):
        if not self.lock.acquire(blocking=False):
            raise ValueError("照片排版预览正在运行，请等待本次预览完成。")
        try:
            identity = uuid.uuid4().hex
            folder = self.folder(identity)
            folder.mkdir(parents=True)
            targets = []
            warnings = compose_preview(self.root, task, paths, folder / "appendix.docx", targets)
            return self.store.save("photo_previews", {
                "id": identity, "status": "running", "stage": "queued", "error": "",
                "warnings": warnings, "pages": 0,
                "targets": targets,
                "word_url": f"/api/photo-previews/{identity}/appendix.docx",
            })
        except (ValueError, OSError, KeyError, RuntimeError, sqlite3.Error):
            self.lock.release()
            raise

    def execute(self, identity):
        def progress(stage):
            item = self.store.get("photo_previews", identity)
            item["stage"] = stage
            self.store.save("photo_previews", item)

        try:
            folder = self.folder(identity)
            targets = self.store.get("photo_previews", identity)["targets"]
            rendered = render(folder / "appendix.docx", folder / "preview.pdf",
                              progress=progress, timeout_seconds=90,
                              bookmark_names=[target["bookmark"] for target in targets])
            progress("rendering_pages")
            with pymupdf.open(folder / "preview.pdf") as pdf:
                if not pdf.page_count:
                    raise ValueError("Word照片预览没有有效页面。")
                for index, page in enumerate(pdf, 1):
                    page.get_pixmap(matrix=pymupdf.Matrix(1.3, 1.3), alpha=False).save(
                        str(folder / f"page-{index}.png"))
                count = pdf.page_count
                page_sizes = [{"width": page.rect.width, "height": page.rect.height} for page in pdf]
                for target in targets:
                    position = rendered.get("bookmarks", {}).get(target["bookmark"])
                    if position is None or not 1 <= position["page"] <= count:
                        raise ValueError("Word未返回完整的照片/标题页码，不能提供可编辑定位。")
                    page_height = page_sizes[position["page"] - 1]["height"]
                    if not math.isfinite(position["y"]) or not 0 <= position["y"] < page_height:
                        raise ValueError("Word返回的照片/标题位置无效，不能提供可编辑定位。")
                    target.update(page=position["page"], top=position["y"] / page_height,
                                  height=min(target["height"], page_height - position["y"]) / page_height)
            item = self.store.get("photo_previews", identity)
            item.update(status="ready", stage="complete", pages=count,
                        targets=targets, page_sizes=page_sizes,
                        pdf_url=f"/api/photo-previews/{identity}/preview.pdf",
                        page_urls=[f"/api/photo-previews/{identity}/page-{i}.png" for i in range(1, count + 1)])
            self.store.save("photo_previews", item)
        except (ValueError, OSError, RuntimeError, sqlite3.Error) as exc:
            LOGGER.exception("Photo layout preview failed: %s", identity)
            item = self.store.get("photo_previews", identity)
            item.update(status="failed", error=str(exc))
            self.store.save("photo_previews", item)
        finally:
            self.lock.release()

    def file(self, identity, filename):
        item = self.store.get("photo_previews", identity)
        allowed = {"appendix.docx"}
        if item["status"] == "ready":
            allowed.add("preview.pdf")
            allowed.update(f"page-{i}.png" for i in range(1, item["pages"] + 1))
        if filename not in allowed:
            raise FileNotFoundError("照片预览尚未生成或文件不存在。")
        path = self.folder(identity) / filename
        if not path.is_file():
            raise FileNotFoundError("照片预览文件不存在，请重新生成。")
        return path
