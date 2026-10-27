"""Local, read-only DOCX import and editable report assembly.

Row identifiers address physical OOXML cells, not python-docx's expanded merged
cell list.  Input files are never saved.  Only explicitly mapped measurements
are filled; a template value is never evidence for the current report.
"""

import hashlib
import json
import re
from contextlib import contextmanager
from contextvars import ContextVar
from copy import deepcopy
from datetime import datetime, time
from decimal import Decimal
from pathlib import Path
from typing import Iterator, Optional
from zipfile import ZipFile

from docx import Document
from docx.document import Document as DocumentObject
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from docx.parts.hdrftr import HeaderPart
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph
from docx.text.run import Run
from docxcompose.composer import Composer
from lxml import etree

from contracts import (AUTOMATIC_OVERVIEW_FIELDS, Equipment, HIDDEN_OVERVIEW_FIELDS, REPORT_ENVIRONMENT_FIELDS, REPORT_TYPE_FOLDERS, SAMPLE_OVERVIEW_FIELDS,
                       TEMPLATE_DEFAULT_FIELDS, TEMPLATE_FALLBACK_WARNING)
from environment import aggregate_project_environment, environment_bounds as _environment_bounds
from validation import parse_date, required_staff_groups
from classification import CLASSIFICATIONS, CLASSIFICATION_ISSUE, resolve_classification
from report_types import FLOWMETER_TYPES, INDEPENDENT_METER_TYPES

PROJECT_METADATA_ROW_HEIGHT_TWIPS = 454


_DOCUMENT_READ_CACHE: ContextVar[Optional[dict[tuple[Path, int, int], DocumentObject]]] = ContextVar(
    "document_read_cache", default=None)
_ASSET_READ_CACHE: ContextVar[Optional[dict[Path, list[dict]]]] = ContextVar(
    "asset_read_cache", default=None)


@contextmanager
def document_read_session() -> Iterator[None]:
    """Reuse read-only parsing within one request, never mutable caller documents."""
    if _DOCUMENT_READ_CACHE.get() is not None:
        yield
        return
    documents_token = _DOCUMENT_READ_CACHE.set({})
    assets_token = _ASSET_READ_CACHE.set({})
    try:
        yield
    finally:
        _ASSET_READ_CACHE.reset(assets_token)
        _DOCUMENT_READ_CACHE.reset(documents_token)


NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
}
FIELDS = {
    "instrument_name": ("计量器具名称", "器具名称", "样品名称"),
    "manufacturer": ("制造单位", "制造厂"),
    "agent": ("代理人",),
    "commissioned_date": ("委托日期",),
    "application_number": ("申请书编号",),
    "classification_code": ("分类编码",),
    "principle": ("原理",),
    "model": ("规格型号", "型号/规格", "样机型号规格", "型号规格", "型号"),
    "sample_number": ("样机编号", "样品编号"),
    "report_number": ("报告编号",),
    "applicant": ("申请单位",),
    "commissioned_unit": ("委托单位",),
    "commissioned_contact": ("委托负责人",),
    "q3": ("Q3",),
    "q3_q1": ("Q3/Q1",),
    "accuracy_class": ("准确度等级", "准确度"),
    "measurement_parameter": ("测量参数",),
    "parameter_unit": ("测量参数单位",),
    "measurement_range": ("测量区间",),
    "display_digits": ("显示位数",),
    "metrological_performance": ("计量性能指标",),
    "display_type": ("显示型式",),
    "environment_temperature": ("型式评价环境温度",),
    "environment_humidity": ("型式评价环境相对湿度",),
    "supply_voltage": ("型式评价电源",),
    "test_location": ("型式评价地点",),
    "other_environment": ("其他环境条件",),
    "construction": ("构造和应用场合",),
    "production_basis": ("生产依据",),
    "test_basis": ("试验依据",),
    "technical_review": ("技术资料审查结论",),
    "evaluation_conclusion": ("型式评价结论",),
    "sealing": ("其他", "封样信息"),
    "disassembly_sample": ("拆样样机编号",),
    "evaluation_period": ("型式评价时间",),
}
MAIN_FIELDS = tuple(key for key in FIELDS
                    if key not in HIDDEN_OVERVIEW_FIELDS
                    and key not in ("instrument_name", "manufacturer", "agent",
                                   "commissioned_date", "application_number",
                                   "classification_code", "principle", "model",
                                   "sample_number", "report_number", "applicant"))
ENVIRONMENT_FIELDS = {"温度": "environment_temperature",
                      "相对湿度": "environment_humidity", "电源": "supply_voltage",
                      "地点": "test_location", "其他": "other_environment"}
ALIASES = {"反向流试验": "反向流", "静磁场试验": "静磁场"}


def _xp(element, expression):
    return etree.XPath(expression, namespaces=NS)(element)


def _text(element):
    return "".join(_xp(element, ".//w:t/text()")).strip()


def _norm(value):
    return re.sub(r"[\s：:，,。．.、（）()－—_]", "", str(value))


def _project_name(value):
    value = re.sub(r"^\d+[.．、]\s*", "", value)
    return ALIASES.get(value, value)


def _root(root):
    root = Path(root).resolve()
    return root.parent if root.name == "system" else root


def _load(path):
    path = Path(path)
    if path.suffix.lower() != ".docx" or path.name.startswith("~$"):
        raise ValueError("仅支持有效的 DOCX 文档，不能使用 Word 锁文件")
    cache = _DOCUMENT_READ_CACHE.get()
    key = None
    try:
        if cache is not None:
            stat = path.stat()
            key = (path.resolve(), stat.st_mtime_ns, stat.st_size)
            if key in cache:
                return deepcopy(cache[key])
        doc = Document(path)
    except Exception as exc:
        raise ValueError("无法读取 DOCX：%s" % path.name) from exc
    _normalize_body(doc.element.body)
    if _xp(doc.element.body, ".//w:altChunk | .//w:del | .//w:ins"):
        raise ValueError("文档包含未接受的修订或外部内容块，请先确认并转换")
    if cache is not None and key is not None and len(cache) < 8:
        cache[key] = deepcopy(doc)
    return doc


def _modern_word_compatibility(doc):
    compat = doc.settings.element.find(qn("w:compat"))
    if compat is None:
        compat = OxmlElement("w:compat")
        doc.settings.element.append(compat)
    modes = [node for node in compat.findall(qn("w:compatSetting"))
             if node.get(qn("w:name")) == "compatibilityMode"]
    if len(modes) > 1:
        raise ValueError("Word兼容模式设置重复，不能安全生成报告。")
    mode = modes[0] if modes else OxmlElement("w:compatSetting")
    if not modes:
        compat.append(mode)
    mode.set(qn("w:name"), "compatibilityMode")
    mode.set(qn("w:uri"), "http://schemas.microsoft.com/office/word")
    mode.set(qn("w:val"), "15")


def _normalize_body(body):
    """Unwrap block SDTs and select a single supported compatibility branch."""
    for child in list(body):
        local = etree.QName(child).localname
        if local == "sdt":
            content = child.find(qn("w:sdtContent"))
            if content is None:
                raise ValueError("内容控件缺少 sdtContent")
            _normalize_body(content)
            index = body.index(child)
            for item in list(content):
                body.insert(index, item)
                index += 1
            body.remove(child)
        elif local == "AlternateContent":
            branches = list(child)
            branch = next((b for b in branches
                           if etree.QName(b).localname == "Choice"
                           and any(etree.QName(c).localname in ("p", "tbl", "sdt")
                                   for c in b)), None)
            if branch is None:
                branch = next((b for b in branches
                               if etree.QName(b).localname == "Fallback"), None)
            if branch is None:
                raise ValueError("不支持的大纲 AlternateContent 结构")
            _normalize_body(branch)
            index = body.index(child)
            for item in list(branch):
                body.insert(index, item)
                index += 1
            body.remove(child)
        elif local not in ("p", "tbl", "sectPr", "bookmarkStart", "bookmarkEnd",
                           "proofErr", "permStart", "permEnd"):
            raise ValueError("不支持的正文结构：" + local)


def _tables(doc):
    return list(_xp(doc.element.body, "./w:tbl"))


def _rows(table):
    return list(table.findall(qn("w:tr")))


def _cells(row):
    return list(row.findall(qn("w:tc")))


def _effective_rows(table):
    """Yield physical cells and their vertically merged owners."""
    above = {}
    for row in _rows(table):
        physical = _cells(row)
        effective = []
        column = 0
        before = row.find(qn("w:trPr"))
        if before is not None:
            skip = before.find(qn("w:gridBefore"))
            if skip is not None:
                column = int(skip.get(qn("w:val"), "0"))
        for cell in physical:
            span = cell.find("./" + qn("w:tcPr") + "/" + qn("w:gridSpan"))
            width = int(span.get(qn("w:val"), "1")) if span is not None else 1
            merge = cell.find("./" + qn("w:tcPr") + "/" + qn("w:vMerge"))
            if merge is not None and merge.get(qn("w:val")) != "restart":
                if column not in above:
                    raise ValueError("表格存在无起点的纵向合并单元格")
                owner = above[column]
            else:
                owner = cell
            effective.append(owner)
            for offset in range(width):
                above[column + offset] = owner
            column += width
        yield physical, effective


def _put(element, value):
    """Replace editable content, retaining paragraph/cell/run formatting."""
    if etree.QName(element).localname == "tc":
        paragraphs = element.findall(qn("w:p"))
        paragraph = paragraphs[0] if paragraphs else OxmlElement("w:p")
        if not paragraphs:
            element.append(paragraph)
        for child in list(element):
            if child is not paragraph and child.tag != qn("w:tcPr"):
                element.remove(child)
    else:
        paragraph = element
    runs = _xp(paragraph, ".//w:r")
    style = deepcopy(runs[0].find(qn("w:rPr"))) if runs and runs[0].find(qn("w:rPr")) is not None else None
    for child in list(paragraph):
        if child.tag != qn("w:pPr"):
            paragraph.remove(child)
    run = OxmlElement("w:r")
    if style is not None:
        run.append(style)
    for index, line in enumerate(str(value or "").split("\n")):
        if index:
            run.append(OxmlElement("w:br"))
        text = OxmlElement("w:t")
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        text.text = line
        run.append(text)
    paragraph.append(run)


def _has_text(value):
    return value is not None and bool(str(value).strip())


def _put_if_filled(element, value):
    if _has_text(value):
        _put(element, value)


def _replace_text_span(paragraph, start, end, replacement):
    """Change a text slot without rebuilding its labels, runs, tabs or formatting."""
    if not _has_text(replacement):
        return
    nodes = _xp(paragraph, ".//w:t | .//w:br[not(@w:type) or @w:type='textWrapping']")
    offset = 0
    inserted = False
    for node in nodes:
        if node.tag == qn("w:br"):
            if start < offset < end:
                node.getparent().remove(node)
            continue
        content = node.text or ""
        node_end = offset + len(content)
        if node_end > start and offset < end:
            left = content[:max(0, start - offset)]
            right = content[max(0, end - offset):]
            _write_slot_text(node, left + (str(replacement) if not inserted else "") + right)
            inserted = True
        offset = node_end
    if not inserted and replacement:
        runs = _xp(paragraph, "./w:r")
        run = OxmlElement("w:r")
        if runs and runs[-1].find(qn("w:rPr")) is not None:
            run.append(deepcopy(runs[-1].find(qn("w:rPr"))))
        text = OxmlElement("w:t")
        run.append(text)
        paragraph.append(run)
        _write_slot_text(text, str(replacement))


def _write_slot_text(node, value):
    lines = value.replace("\r\n", "\n").split("\n")
    node.text = lines[0]
    node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    parent = node.getparent()
    index = parent.index(node) + 1
    for line in lines[1:]:
        parent.insert(index, OxmlElement("w:br"))
        text = OxmlElement("w:t")
        text.text = line
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        parent.insert(index + 1, text)
        index += 2


def _replace_paragraph_value(paragraph, value, after_label=False, preserve_terminator=False):
    if not _has_text(value):
        return
    text = "".join(_xp(paragraph, ".//w:t/text()"))
    start = 0
    if after_label:
        colon = re.search(r"[：:]", text)
        if not colon:
            raise ValueError("未找到模板字段的冒号，无法保留格式替换：" + text)
        start = colon.end()
    while start < len(text) and text[start].isspace():
        start += 1
    end = max(start, len(text.rstrip()))
    if preserve_terminator and end > start and text[end - 1] in "；;":
        end = len(text[:end - 1].rstrip())
        value = value.rstrip(" \t\r\n；;")
    _replace_text_span(paragraph, start, end, value)


GAS_RESULT_FILENAME_ALIASES = {
    "19.功能性要求（密封性）-2 models.docx": "19.功能性要求-2 models.docx",
    "20.功能性要求（气流方向）-2 models.docx": "20.功能性要求-2 models.docx",
    "21.功能性要求（零流量读数）-2 models.docx": "21.功能性要求-2 models.docx",
    "39.机械环境（振动、跌落）-2 models.docx": "39.机械环境-2 models.docx",
    "41.机械环境（复测）-2 models.docx": "41.机械环境-2 models.docx",
}


def result_variant_matches(asset, variant):
    if variant == asset["variant"]:
        return True
    if asset["kind"] != "result" or asset.get("report_type") != "燃气表" or asset.get("folders"):
        return False
    original_name = GAS_RESULT_FILENAME_ALIASES.get(asset["name"] + ".docx")
    return original_name is not None and variant == Path(original_name).stem


def _asset_files(root):
    from signature_templates import SIGNATURE_FILENAMES

    base = _root(root) / "report model"
    cache = _ASSET_READ_CACHE.get()
    if cache is not None and base in cache:
        return deepcopy(cache[base])
    aliases_path = base / "result-filename-aliases.json"
    filename_aliases = json.loads(aliases_path.read_text(encoding="utf-8")) if aliases_path.exists() else {}
    if not isinstance(filename_aliases, dict) or any(
            not isinstance(current, str) or not isinstance(original, str)
            or not current.startswith("type evaluation results/")
            or not original.startswith("type evaluation results/")
            or ".." in Path(current).parts or ".." in Path(original).parts
            or Path(current).parent != Path(original).parent
            or not current.endswith(".docx") or not original.endswith(".docx")
            for current, original in filename_aliases.items()):
        raise ValueError("结果素材文件名别名配置无效，不能猜测素材身份。")
    kinds = {"items": "type evaluation items", "observe": "observe",
             "result": "type evaluation results"}
    result = []
    type_folders = {name: key for key, name in REPORT_TYPE_FOLDERS.items()}
    type_folders.update(dict.fromkeys(
        ("flow meter", "liquid flowmeter", "liquid flow meter"), "流量计（液体）"))
    for kind, folder in kinds.items():
        directory = base / folder
        if not directory.is_dir():
            raise ValueError("素材目录不存在：" + folder)
        for path in sorted(directory.rglob("*.docx")):
            if path.name.startswith("~$"):
                continue
            if kind == "result" and path.parent == directory and path.name in SIGNATURE_FILENAMES.values():
                continue
            relative = path.relative_to(base).as_posix()
            original_relative = filename_aliases.get(relative, relative) if kind == "result" else relative
            original_name = Path(original_relative).name
            folders = list(path.relative_to(directory).parts[:-1])
            report_type = type_folders.get(folders[0], "水表") if folders else "水表"
            typed = bool(folders and folders[0] in type_folders)
            if typed:
                folders = folders[1:]
            # Water materials keep their pre-category IDs so saved tasks still resolve.
            logical = "/".join([folder, *folders, original_name]) if report_type == "水表" else original_relative
            project_stem = Path(original_name).stem
            if kind == "result" and report_type == "燃气表" and not folders:
                legacy_name = GAS_RESULT_FILENAME_ALIASES.get(original_name)
                if legacy_name:
                    logical = Path(original_relative).with_name(legacy_name).as_posix()
                    project_stem = Path(legacy_name).stem
            project_folder = folders[0] if folders else path.stem
            if kind == "result" and report_type in FLOWMETER_TYPES and len(folders) > 1:
                project_folder = next((part for part in folders if re.match(r"^\d+[.．、]", part)),
                                      project_folder)
            project = _project_name(project_folder) if kind == "result" else ""
            if report_type in INDEPENDENT_METER_TYPES and kind == "result":
                if re.match(r"^\d+[.．、]", project_stem):
                    project = re.sub(r"-\d+\s*models?$", "", _project_name(project_stem), flags=re.I).strip()
                if report_type in FLOWMETER_TYPES:
                    for group in ("贮存环境性能", "电磁兼容"):
                        if project.startswith(group + "—"):
                            project = group
            result.append({
                "id": hashlib.sha256(logical.encode()).hexdigest()[:20],
                "_path_id": hashlib.sha256(relative.encode()).hexdigest()[:20],
                "_legacy_path_id": hashlib.sha256(original_relative.encode()).hexdigest()[:20],
                "kind": kind, "name": path.stem,
                "report_type": report_type,
                "project": project,
                "folders": folders,
                "variant": Path(original_name).stem, "_path": path,
            })
    identifiers = [asset["id"] for asset in result]
    if len(set(identifiers)) != len(identifiers):
        raise ValueError("素材存在重复的旧目录与类型目录路径，请保留唯一版本后刷新。")
    if cache is not None and len(cache) < 8:
        cache[base] = deepcopy(result)
    return result


def _asset(root, asset_id, kind=None):
    matches = [a for a in _asset_files(root) if asset_id in (a["id"], a["_path_id"], a["_legacy_path_id"])]
    if len(matches) != 1 or (kind and matches[0]["kind"] != kind):
        raise ValueError("素材不存在或类型不匹配：" + str(asset_id))
    return matches[0]


def _field_candidates(doc, source, record=False):
    fields = {}

    def add(key, value, location):
        value = value.strip()
        if value and value not in FIELDS.get(key, ()):
            entry = {"value": value, "source": source + " · " + location}
            if entry not in fields.setdefault(key, []):
                fields[key].append(entry)

    aliases = {_norm(label): key for key, labels in FIELDS.items() for label in labels}
    header_labels = set(aliases) | {
        "序号", "数量", "编号", "名称", "备注", "主要性能指标",
        "证书有效期", "要求", "技术要求", "样机编号", "试验结果",
    }
    for ti, table in enumerate(_tables(doc)):
        basic = "样品名称" in _text(table) and "制造厂" in _norm(_text(table))
        table_rows = _rows(table)
        first_cells = _cells(table_rows[0]) if table_rows else []
        headers = [_norm(_text(cell)) for cell in first_cells]
        canonical_headers = [
            "Q3" if re.fullmatch(r"Q3[（(].+[）)]", _text(cell)) else header
            for cell, header in zip(first_cells, headers)
        ]
        header_table = bool(headers and (
            headers[0] == "序号" or
            (len(headers) > 1 and all(header in header_labels for header in canonical_headers))
        ))
        if header_table:
            if "样机型号规格" in headers and "编号" in headers:
                for ri, row in enumerate(table_rows[1:], 1):
                    cells = _cells(row)
                    for column, key in ((headers.index("样机型号规格"), "model"),
                                        (headers.index("编号"), "sample_number"),
                                        (2, "q3"), (3, "q3_q1"), (4, "accuracy_class")):
                        if column < len(cells):
                            add(key, _text(cells[column]), "表%d/行%d/列%d" % (ti + 1, ri + 1, column + 1))
            elif "测量参数" in headers and "计量性能指标" in headers:
                for ri, row in enumerate(table_rows[1:], 1):
                    for column, cell in enumerate(_cells(row)):
                        if column < len(headers) and headers[column] in aliases:
                            add(aliases[headers[column]], _text(cell), "表%d/行%d/列%d" % (ti + 1, ri + 1, column + 1))
            elif (any(aliases.get(header) == "model" for header in canonical_headers)
                  and any(header in canonical_headers for header in ("Q3", "样机编号", "样品编号", "准确度等级"))
                  and "证书有效期" not in canonical_headers):
                for ri, row in enumerate(table_rows[1:], 1):
                    cells = _cells(row)
                    if len(cells) != len(first_cells):
                        continue
                    for column, header in enumerate(canonical_headers):
                        key = aliases.get(header)
                        if key:
                            add(key, _text(cells[column]), "表%d/行%d/列%d" % (ti + 1, ri + 1, column + 1))
            continue
        for ri, row in enumerate(table_rows):
            cells = _cells(row)
            for ci, cell in enumerate(cells[:-1]):
                label = _norm(_text(cell))
                key = aliases.get(label)
                if basic and label == "编号":
                    key = "sample_number"
                if record and basic and ri == 0 and ci == len(cells) - 2 and label == "证书编号":
                    key = "report_number"
                if key:
                    add(key, _text(cells[ci + 1]), "表%d/行%d/列%d" % (ti + 1, ri + 1, ci + 2))
    for pi, paragraph in enumerate(_xp(doc.element.body, "./w:p")):
        text = _text(paragraph)
        compact = re.sub(r"\s+", "", text)
        environment = re.match(r"^[1-5][.．、]?(温度|相对湿度|电源|地点|其他)[：:](.*)", compact)
        if environment:
            value = environment.group(2)
            if environment.group(1) == "电源":
                value = value.rstrip("；;")
            add(ENVIRONMENT_FIELDS[environment.group(1)], value, "段落%d" % (pi + 1))
        if "经审查，申请单位提交" in text:
            add("technical_review", text, "段落%d" % (pi + 1))
        if "本院已" in text and "封样" in text:
            add("sealing", text, "段落%d" % (pi + 1))
        if text.startswith("注：编号为"):
            sample = re.search(r"编号为(.+?)的样机", text)
            if sample:
                add("disassembly_sample", sample.group(1), "段落%d" % (pi + 1))
        if "显示型式：" in compact:
            add("display_type", compact.split("显示型式：", 1)[1], "段落%d" % (pi + 1))
        if "建议批准下列型号计量器具的型式" in _norm(text):
            add("evaluation_conclusion", text.split("建议批准")[0].rstrip("，"), "段落%d" % (pi + 1))
        for key, labels in FIELDS.items():
            if key in ("q3", "q3_q1"):
                continue
            for label in labels:
                clean = re.sub(r"\s+", "", label)
                if compact.startswith(clean):
                    tail = compact[len(clean):].lstrip("：:")
                    if tail and not tail.startswith(("及", "、", "（", "(")):
                        add(key, tail, "段落%d" % (pi + 1))
                    break
    if record:
        seen = set()
        for reference in _xp(doc.element, ".//w:headerReference"):
            identity = reference.get(qn("r:id"))
            if identity in seen:
                continue
            seen.add(identity)
            header = doc.part.related_parts[identity].element
            for pi, paragraph in enumerate(_xp(header, ".//w:p"), 1):
                text = re.sub(r"\s+", "", _text(paragraph))
                match = re.search(r"报告编号[：:]?([A-Za-z0-9][A-Za-z0-9\-–—－/]*)(?![A-Za-z0-9])", text)
                if match:
                    add("report_number", match.group(1), "页眉%s/段落%d" % (identity, pi))
            for ti, table in enumerate(_xp(header, ".//w:tbl"), 1):
                for ri, row in enumerate(_rows(table), 1):
                    cells = _cells(row)
                    for ci, cell in enumerate(cells[:-1]):
                        if _norm(_text(cell)) == "报告编号":
                            add("report_number", _text(cells[ci + 1]), "页眉%s/表%d/行%d" % (identity, ti, ri))
    return fields


def catalog(root):
    from basic_options import template_options
    from contracts import OBSERVATION_MANUAL_FIELDS
    from observation_documents import COPY_WARNING, LABELS, copy_warning

    assets = _asset_files(root)
    master = _load(_root(root) / "report model/report.docx")
    old = _field_candidates(master, "报告模板")
    old["test_basis"] = [{"value": "\n".join(_text(p) for p in _report_test_basis_paragraphs(master)
                                           if _text(p)), "source": "报告模板 · 三、型式评价的依据"}]
    overview = []
    for key, labels in FIELDS.items():
        if key in HIDDEN_OVERVIEW_FIELDS:
            continue
        values = list(dict.fromkeys(e["value"] for e in old.get(key, [])))
        overview.append({"key": key, "label": labels[0],
                         "old": "\n".join(values), "value": "", "confirmed": False,
                         "candidates": []})
    overview.extend({"key": key, "label": label, "old": "", "value": "", "confirmed": False,
                     "candidates": []} for key, label in OBSERVATION_MANUAL_FIELDS.items())
    return {"assets": [{k: v for k, v in a.items() if not k.startswith("_")} for a in assets],
            "report_type_folders": REPORT_TYPE_FOLDERS,
            "overview": overview,
            "basic_options": template_options(master),
            "sample_table_column_selection": True,
            "classification_options": CLASSIFICATIONS,
            "test_basis_from_outline": True,
            "observation_labels": LABELS,
            "observation_copy_warnings": {kind: copy_warning(kind) for kind in REPORT_TYPE_FOLDERS},
            "hidden_overview_fields": sorted(HIDDEN_OVERVIEW_FIELDS),
            "warnings": [TEMPLATE_FALLBACK_WARNING, COPY_WARNING]}


def _result_schema(doc, *, allow_empty_samples=False):
    result = []
    found_table = False
    for ti, table in enumerate(_tables(doc)):
        rows = _rows(table)
        if not rows or not all(word in _text(rows[0]) for word in ("技术要求", "样机编号", "试验结果")):
            continue
        found_table = True
        for ri, (physical, effective) in enumerate(_effective_rows(table)):
            if ri == 0 or (physical and _norm(_text(physical[-1])) == "试验结果"):
                continue
            if len(physical) < 5:
                raise ValueError("试验结果表列结构不支持：表%d行%d" % (ti + 1, ri + 1))
            sample = _text(effective[-2])
            if not sample:
                if allow_empty_samples and not _text(effective[-1]):
                    continue
                raise ValueError("试验结果行缺少样机标识")
            requirement = " | ".join(dict.fromkeys(_text(c) for c in effective[2:-2] if _text(c)))
            result.append({"id": "t%d:r%d:c%d" % (ti, ri, len(physical) - 1),
                           "requirement": requirement, "sample_label": sample,
                           "sample_id": "", "value": "", "source": "",
                           "_value_cell": physical[-1], "_sample_cell": physical[-2]})
    if not result and not (allow_empty_samples and found_table):
        raise ValueError("未找到可编辑的原生试验结果表")
    return result


def result_material_version(root):
    digest = hashlib.sha256()
    for asset in sorted(_asset_files(root), key=lambda item: item["id"]):
        if asset["kind"] != "result":
            continue
        metadata = {key: asset[key] for key in ("id", "report_type", "project", "variant")}
        digest.update(json.dumps(metadata, ensure_ascii=False, sort_keys=True).encode("utf-8"))
        digest.update(hashlib.sha256(asset["_path"].read_bytes()).digest())
    return digest.hexdigest()


def project_from_asset(root, asset_id):
    asset = _asset(root, asset_id, "result")
    rows = _result_schema(_load(asset["_path"]))
    return {"id": asset_id, "asset_id": asset_id, "name": asset["project"],
            "variant": asset["variant"],
            "rows": [{k: v for k, v in r.items() if not k.startswith("_")} for r in rows],
            "start": "", "end": "", "temperature": "", "humidity": "",
            "conclusion": "", "abnormal": "", "evaluators": [], "reviewers": [],
            "equipment_ids": [], "confirmed": False, "notes": [], "source": ""}


def reconcile_project_rows(root, project):
    asset = _asset(root, project["asset_id"], "result")
    if asset["report_type"] == "水表":
        return deepcopy(project)
    if (_project_name(project["name"]) != asset["project"]
            or not result_variant_matches(asset, project["variant"])):
        raise ValueError("项目名称或方案与素材不一致，请重新选择。")
    actual = _result_schema(_load(asset["_path"]))
    result = deepcopy(project)
    result["rows"] = reconcile_result_rows(actual, project.get("rows", []), project["name"])
    if [r["id"] for r in result["rows"]] != [r["id"] for r in project.get("rows", [])]:
        note = "素材表格位置已变化，已按相同顺序的样机和技术要求同步行标识；实测值及人工修改记录保留。"
        if note not in result.setdefault("notes", []):
            result["notes"].append(note)
    return result


def _section_schema(doc, kind):
    found = []
    for ti, table in enumerate(_tables(doc)):
        rows = _rows(table)
        if not rows:
            continue
        head = _text(rows[0])
        if "备注" not in head or not ("评价项目" in head or "要求" in head):
            continue
        for ri, (physical, effective) in enumerate(_effective_rows(table)):
            if ri == 0 or len(physical) < 3:
                continue
            # A merged "not applicable" result spans the three output columns.
            merged_output = _xp(physical[-1], "./w:tcPr/w:gridSpan/@w:val") == ["3"]
            outputs = physical[-1:] if merged_output else physical[-3:]
            labels = effective[:-len(outputs)]
            if kind == "items":
                label = " / ".join(dict.fromkeys(_text(c) for c in labels[1:] if _text(c)))
                requirement = ""
            else:
                label = " / ".join(dict.fromkeys(_text(c) for c in labels[:-1] if _text(c)))
                requirement = _text(labels[-1]) if labels else ""
            if not label:
                raise ValueError("无法定位观察/一览表项目标签")
            found.append({"id": "t%d:r%d" % (ti, ri), "label": label,
                          "requirement": requirement, "value": "", "note": "",
                          "_outputs": outputs})
    if not found:
        raise ValueError("未找到观察或项目一览表")
    return found


def section_rows(root, asset_id):
    asset = _asset(root, asset_id)
    if asset["kind"] not in ("items", "observe"):
        raise ValueError("非观察/项目一览表素材")
    return [{k: v for k, v in row.items() if not k.startswith("_")}
            for row in _section_schema(_load(asset["_path"]), asset["kind"])]


def _unique_anchor(body, phrase):
    matches = [e for e in body if e.tag == qn("w:p") and _norm(phrase) in _norm(_text(e))]
    if len(matches) != 1:
        raise ValueError("章节标记缺失或不唯一：%s（%d处）" % (phrase, len(matches)))
    return matches[0]


def _outline_spacing_paragraph(element):
    if element.tag != qn("w:p"):
        return False
    if _xp(element, "./w:pPr/w:sectPr | ./w:pPr/w:numPr | ./w:pPr/w:pageBreakBefore"):
        return False
    # Only discard known whitespace; unknown nodes may carry non-text content or references.
    for child in element:
        if child.tag == qn("w:pPr"):
            continue
        if child.tag != qn("w:r"):
            return False
        for content in child:
            if content.tag == qn("w:rPr"):
                continue
            if content.tag == qn("w:t") and not (content.text or "").strip():
                continue
            if content.tag in (qn("w:tab"), qn("w:cr")):
                continue
            if (content.tag == qn("w:br") and
                    content.get(qn("w:type"), "textWrapping") == "textWrapping" and
                    content.get(qn("w:clear"), "none") == "none"):
                continue
            return False
    return True


def _coverage_introduction(element):
    return element.tag == qn("w:p") and any(
        phrase in _norm(_text(element)) for phrase in (
            "故试验结果可覆盖", "技术参数如下",
        )
    )


def _outline_regions(doc):
    body = doc.element.body
    component_start = _unique_anchor(body, "关键零部件和材料表")
    sample_start = _unique_anchor(body, "提供试验样机")
    # Some source templates omit the leading 本 in the literal XML.
    end = _unique_anchor(body, "次型式评价项目")
    start_index, sample_index, end_index = [body.index(e) for e in (component_start, sample_start, end)]
    if not start_index < sample_index < end_index:
        raise ValueError("大纲第4、5、6项顺序不正确")
    components = [e for e in list(body)[start_index + 1:sample_index] if e.tag == qn("w:tbl")]
    if not components:
        raise ValueError("大纲第4项范围内未找到完整原生表格")
    component_other = [e for e in list(body)[start_index + 1:sample_index]
                       if e.tag != qn("w:tbl") and _text(e)]
    if component_other:
        raise ValueError("关键零部件区间含非表格正文，需确认完整边界")
    starts = [e for e in list(body)[sample_index + 1:end_index]
              if _coverage_introduction(e)]
    if len(starts) != 1:
        raise ValueError("大纲第5项覆盖型号引导语缺失或不唯一")
    coverage = list(body)[body.index(starts[0]) + 1:end_index]
    technical_parameters = "技术参数如下" in _norm(_text(starts[0]))
    if technical_parameters and not any(not _outline_spacing_paragraph(e) for e in coverage):
        raise ValueError("技术参数范围内未找到正文或完整表格")
    if not technical_parameters and not any(e.tag == qn("w:tbl") for e in coverage):
        raise ValueError("覆盖型号范围内未找到完整表格")
    for element in components + coverage:
        if element.tag in (qn("w:sectPr"), qn("w:altChunk")) or _xp(element, ".//w:sectPr | .//w:altChunk"):
            raise ValueError("大纲复制范围含不支持的节边界或外部内容")
    if not technical_parameters:
        coverage = [element for element in coverage if not _outline_spacing_paragraph(element)]
    return components, coverage


def _preview(elements):
    lines = []
    for element in elements:
        if element.tag == qn("w:tbl"):
            lines.extend(" | ".join(_text(c) for c in _cells(r)) for r in _rows(element))
        else:
            lines.append(_text(element))
    return "\n".join(lines)


def _sample_quantity_header(text):
    return re.sub(r"\s+", "", text) in ("数量", "数量（台）", "数量(台)")


def outline_samples(doc):
    body = doc.element.body
    start = _unique_anchor(body, "提供试验样机")
    end = _unique_anchor(body, "次型式评价项目")
    blocks = list(body)[body.index(start) + 1:body.index(end)]
    stop = next((i for i, e in enumerate(blocks)
                 if _coverage_introduction(e)), len(blocks))
    blocks = blocks[:stop]
    tables = [table for block in blocks
              for table in ([block] if block.tag == qn("w:tbl") else _xp(block, ".//w:tbl"))
              if _rows(table) and any(_sample_quantity_header(_text(c)) for c in _cells(_rows(table)[0]))]
    if len(tables) != 1:
        raise ValueError("大纲第5项中带「数量」列的样机表缺失或不唯一，请核对原生表格。")
    table = tables[0]
    if _xp(table, ".//w:gridSpan | .//w:sectPr | .//w:altChunk | .//w:tbl"):
        raise ValueError("样机表含不支持的横向合并、嵌套或节结构，请核对。")
    rows = list(_effective_rows(table))
    headers = [_text(c) for c in rows[0][1]]
    quantities = [i for i, text in enumerate(headers) if _sample_quantity_header(text)]
    if len(quantities) != 1 or len(rows) < 2:
        raise ValueError("样机表数量列不唯一或没有数据行。")
    column = quantities[0]
    groups = []
    for index, (_, cells) in enumerate(rows[1:]):
        if len(cells) != len(headers):
            raise ValueError("样机表数据列数与表头不一致。")
        values = [_text(c) for c in cells]
        count = re.fullmatch(r"\s*([1-9]\d*)\s*(?:台|个|只)?\s*", values[column])
        if not count:
            raise ValueError("样机表数量必须为正整数（可带台/个/只）：" + values[column])
        prefix, number = "", index + 1
        while number:
            number, digit = divmod(number - 1, 26)
            prefix = chr(65 + digit) + prefix
        groups.append({"values": values, "labels": [
            "%s%03d" % (prefix, n) for n in range(1, int(count.group(1)) + 1)]})
    schema = {"headers": headers, "quantity_column": column, "groups": groups,
              "note": "\n".join(_text(e) for e in blocks if e.tag != qn("w:tbl") and "拆" in _text(e))}
    return schema, table


def reconcile_samples(task, schema, strict=True):
    """Keep user mappings only when the source rows still have the same meaning."""
    previous = task.get("sample_schema", {})
    samples = task.get("samples", [])
    if previous and previous != schema:
        if any(s.get("number", "").strip() or s.get("disassembly") or s.get("sealed") for s in samples):
            raise ValueError("大纲样机表已改变。请先清空旧样机对应关系，再选择新大纲，避免编号对应错型号。")
        if any(r.get("sample_id", "").strip() for p in task.get("projects", []) for r in p["rows"]):
            raise ValueError("大纲样机表已改变，项目中仍有旧样机编号。请先清空样机对应关系。")
        samples = []
        task["samples_confirmed"] = False
    labels = [label for group in schema.get("groups", []) for label in group["labels"]]
    if samples and [s["label"] for s in samples] != labels:
        raise ValueError("样机清单与大纲数量或分组不一致，请重新读取大纲。")
    if not samples:
        samples = [{"label": label, "number": "", "disassembly": False, "sealed": False} for label in labels]
    if not schema:
        task.update(samples=[], sample_schema={}, sample_issues=[])
        return task
    by_label = {s["label"]: s for s in samples}
    mappings = {}
    issues = []
    for sample in samples:
        sample["number"] = sample.get("number", "").strip()
        if sample["number"]:
            mappings[sample["label"]] = sample["number"]
    for project in task.get("projects", []):
        for row in project["rows"]:
            label, number = row.get("sample_label", "").strip(), row.get("sample_id", "").strip()
            if label not in by_label:
                issues.append("结果素材样机代号不在大纲清单中：" + label + "，请选择匹配的素材方案。")
                continue
            if number:
                if label in mappings and mappings[label] != number:
                    issues.append("样机清单与项目编号映射冲突：" + label)
                else:
                    mappings[label] = number
    if len(set(mappings.values())) != len(mappings):
        issues.append("不同样机不能使用同一编号，请检查样机清单。")
    if sum(bool(s.get("disassembly")) for s in samples) > 1:
        issues.append("只能选择一台拆样样机。")
    effective_numbers = [mappings.get(s["label"], s["label"]) for s in samples]
    if len(set(effective_numbers)) != len(effective_numbers):
        issues.append("已填编号与其他样机的待填报告代号重复，请核对。")
    task.update(samples=samples, sample_schema=schema, sample_issues=list(dict.fromkeys(issues)))
    if issues:
        if strict:
            raise ValueError("；".join(task["sample_issues"]))
        task["samples_confirmed"] = False
        return task
    for sample in samples:
        value = mappings.get(sample["label"], "")
        if sample["number"] != value:
            task["samples_confirmed"] = False
        sample["number"] = value
    for project in task.get("projects", []):
        for row in project["rows"]:
            value = mappings.get(row.get("sample_label", "").strip(), "")
            if row.get("sample_id", "") != value:
                project["confirmed"] = False
            if row.get("matched_sample_id") and row["matched_sample_id"] != value:
                if row.get("auto_value") and row.get("value") == row["auto_value"]:
                    row.update(value="", source="", auto_value="", matched_sample_id="",
                               match_issue="样机对应已改变，请重新解析本次资料")
                elif _has_text(row.get("value", "")):
                    row["match_issue"] = "样机对应已改变，保留的人工结果属于旧编号，请核对后重新保存"
            row["sample_id"] = value
    return task


def outline_cover_model(doc, source):
    blocks = list(doc.element.body)
    endings = [index for index, block in enumerate(blocks)
               if block.tag == qn("w:p") and (
                   _norm(_text(block)).startswith("批准日期")
                   or re.fullmatch(r"\d*适用范围", _norm(_text(block))))]
    if not endings:
        raise ValueError("封样信息：无法确定大纲封面范围，请核对批准日期或适用范围标题。")
    candidates = []
    for index, block in enumerate(blocks[:min(endings)], 1):
        if block.tag == qn("w:p"):
            match = re.match(r"^\s*型\s*号(?!\s*规\s*格)\s*[：:]?\s*(.+?)\s*$", Paragraph(block, doc).text)
            if match and match.group(1).strip().strip("：:"):
                candidates.append({"value": match.group(1).strip(),
                                   "source": f"{source} · 封面型号 · 正文块{index}"})
        elif block.tag == qn("w:tbl"):
            for ri, row in enumerate(_rows(block), 1):
                cells = _cells(row)
                for ci, cell in enumerate(cells[:-1]):
                    if _norm(_text(cell)) == "型号" and _text(cells[ci + 1]):
                        candidates.append({"value": _text(cells[ci + 1]),
                                           "source": f"{source} · 封面型号 · 正文块{index}/行{ri}"})
    if len(candidates) != 1:
        raise ValueError("封样信息：大纲封面型号缺失或不唯一，不使用样机表中的型号代替。")
    prefix, marker, _suffix = candidates[0]["value"].partition("系列")
    candidates[0]["value"] = prefix + marker
    return candidates[0]


def sync_other_information(task, outline, source="当前大纲"):
    from supplementary_notes import clause_order, outline_note_ids, read_supplementary_notes, strip_number

    field = next((f for f in task.get("overview", []) if f["key"] == "sealing"), None)
    task["sealing_mode"] = task.get("sealing_mode", "manual" if field and _has_text(field.get("value")) else "auto")
    task["other_notes"] = task.get("other_notes") or []
    supplementary = read_supplementary_notes(outline, source)
    notes = [entry["text"] for entry in supplementary["entries"]]
    ids = outline_note_ids(task, supplementary["entries"])
    order = clause_order(task, ids)
    if field is not None and notes != task.get("supplementary_notes", []):
        field["confirmed"] = False
    task.update(supplementary_notes=notes, supplementary_source=supplementary["source"],
                supplementary_issues=supplementary["issues"],
                supplementary_note_ids=ids, other_information_order=order,
                other_sealing_number=order.index("sealing") + 1)
    for note in task.get("custom_supplementary_notes", []):
        if not strip_number(note["text"]).strip():
            task["supplementary_issues"].append(f"其他/第{order.index(note['id']) + 1}条新增补充说明正文为空，请填写或删除。")
    task.update(sealing_model="", sealing_model_source="", sealing_issues=[])
    issues = []
    if outline is None:
        issues.append("封样信息：请先选择本次大纲，读取封面型号。")
    else:
        try:
            model = outline_cover_model(outline, source)
            task.update(sealing_model=model["value"], sealing_model_source=model["source"])
        except ValueError as exc:
            issues.append(str(exc))
    instrument = next((f.get("value", "").strip() for f in task.get("overview", [])
                       if f["key"] == "instrument_name"), "")
    if not instrument:
        issues.append("封样信息：请先填写总览中的计量器具名称。")
    selected = [s for s in task.get("samples", []) if s.get("sealed")]
    if not selected:
        issues.append("封样信息：请在样机清单中勾选封样样机，不默认选择拆样样机。")
    missing = [s["label"] for s in selected if not s.get("number", "").strip()]
    if missing:
        issues.append("封样信息：已勾选封样样机缺少编号：" + "、".join(missing))
    if sum(bool(s.get("disassembly")) for s in selected) > 1:
        issues.append("封样信息：只能选择一台拆样样机。")
    text = ""
    if not issues:
        ordered = sorted(selected, key=lambda s: bool(s.get("disassembly")))
        numbers = [s["number"].strip() + ("（用于拆样，拍摄关键零部件）" if s.get("disassembly") else "")
                   for s in ordered]
        text = (f"{task['other_sealing_number']}.本院已对{task['sealing_model']}{instrument}（{'、'.join(numbers)}）"
                f"共{len(selected)}台样机进行了封样，并交付申请单位，"
                "申请单位应妥善保存封印和标记好的试验样机至停止生产该型式计量器具之日起不少于五年。")
    task["sealing_issues"] = issues if task["sealing_mode"] == "auto" else []
    if field is not None:
        field["label"] = "其他"
        field["candidates"] = ([{"value": text, "source": task["sealing_model_source"] +
                                "；总览计量器具名称；已勾选封样样机清单"}] if text else [])
        if task["sealing_mode"] == "auto":
            if field.get("value") != text:
                field["confirmed"] = False
            field["value"] = text
        elif _has_text(field.get("value")):
            numbered = f"{task['other_sealing_number']}." + strip_number(field["value"].strip())
            if field["value"] != numbered:
                field["confirmed"] = False
            field["value"] = numbered
    return task


def _sample_header_min_width(outline, table, cell, *, compress_punctuation=False):
    from math import ceil
    from word_fonts import text_width

    styles = outline.styles
    style_ids = _xp(table, "./w:tblPr/w:tblStyle/@w:val")
    style = next((item for item in styles if item.style_id in style_ids), None)
    if style is None:
        defaults = _xp(styles.element, "./w:style[@w:type='table'][@w:default='1']")
        style_properties = [defaults[0]] if defaults else []
    else:
        style_properties = []
        while style is not None:
            style_properties.append(style.element)
            style = style.base_style
    margins = 0
    for side in ("left", "right"):
        candidates = _xp(cell, f"./w:tcPr/w:tcMar/w:{side}/@w:w")
        candidates += _xp(table, f"./w:tblPr/w:tblCellMar/w:{side}/@w:w")
        for properties in style_properties:
            candidates += _xp(properties, f"./w:tblPr/w:tblCellMar/w:{side}/@w:w")
        margins += int(candidates[0]) if candidates else 108
    points, largest_size, inset = 0, 0, 0
    for element in _xp(cell, "./w:p"):
        paragraph = Paragraph(element, outline)
        for run in paragraph.runs:
            size = run.font.size
            style = run.style
            while size is None and style is not None:
                size = style.font.size
                style = style.base_style
            style = paragraph.style
            while size is None and style is not None:
                size = style.font.size
                style = style.base_style
            defaults = _xp(styles.element, "./w:docDefaults/w:rPrDefault/w:rPr/w:sz/@w:val")
            if size is None and not defaults:
                raise ValueError("样机表表头字号无法确定。")
            size = size.pt if size is not None else int(defaults[0]) / 2
            largest_size = max(largest_size, size)
            scale = _xp(run._r, "./w:rPr/w:w/@w:val")
            spacing = _xp(run._r, "./w:rPr/w:spacing/@w:val")
            text = run.text.replace("\n", "")
            width = text_width(text, size)
            if compress_punctuation:
                # Conservative allowance for Word's fullwidth punctuation spacing, not glyph scaling.
                width -= sum(text.count(char) for char in "：（）") * size / 8
            points += width * (int(scale[0]) / 100 if scale else 1)
            points += len(text) * (int(spacing[0]) / 20 if spacing else 0)
        indents = _xp(element, "./w:pPr/w:ind")
        if indents:
            inset = max(inset, sum(int(indents[0].get(qn("w:" + side), "0"))
                                   for side in ("left", "right", "firstLine")))
    if re.sub(r"\s+", "", _text(cell)).replace("³", "3") in {"Q3（m3/h）", "Q3(m3/h)"}:
        points = text_width("（m3/h）", largest_size)
    return ceil(points * 20) + margins + max(inset, 0) + 20


def _join_sample_header_lines(cell):
    label = re.sub(r"\s+", "", _text(cell)).replace("³", "3")
    if label not in {"序号", "型号", "规格", "型号规格", "Q3/Q1", "Q3Q1",
                     "准确度等级", "编号", "备注", "流量范围（m3/h）", "流量范围(m3/h)"}:
        return
    for line in _xp(cell, ".//w:br[not(@w:type) or @w:type='textWrapping'] | .//w:cr"):
        line.getparent().remove(line)
    paragraphs = _xp(cell, "./w:p")
    for paragraph in paragraphs[1:]:
        for child in list(paragraph):
            if child.tag != qn("w:pPr"):
                paragraphs[0].append(child)
        cell.remove(paragraph)


def sample_table_columns(schema, excluded_columns=()):
    headers = schema.get("headers", [])
    fixed = {index for index, header in enumerate(headers) if "准确度" in _norm(header)}
    if headers:
        fixed.add(schema["quantity_column"])
    if (any(type(index) is not int or index < 0 or index >= len(headers) for index in excluded_columns)
            or len(excluded_columns) != len(set(excluded_columns))):
        raise ValueError("样机信息表隐藏列无效，请重新核对当前大纲的列选择。")
    if fixed.intersection(excluded_columns):
        raise ValueError("样机信息表的准确度等级、编号为固定输出列，不能取消。")
    return [index for index in range(len(headers)) if index not in excluded_columns]


def _sample_table_fragment(outline, schema, table, samples, excluded_columns=()):
    included = sample_table_columns(schema, excluded_columns)
    table = deepcopy(table)
    rows = _rows(table)
    grid = _xp(table, "./w:tblGrid/w:gridCol")
    if len(grid) != len(schema["headers"]):
        raise ValueError("样机表列宽定义与表头不一致。")
    widths = [int(c.get(qn("w:w"))) for c in grid]
    total = sum(widths)
    serial_grid = deepcopy(grid[0])
    grid[0].addprevious(serial_grid)
    numbers = {s["label"]: s["number"] or s["label"] for s in samples}
    for index, row in enumerate(rows):
        cells = _cells(row)
        serial_cell = deepcopy(cells[0])
        for merge in _xp(serial_cell, "./w:tcPr/w:vMerge"):
            merge.getparent().remove(merge)
        _put(serial_cell, "序号" if index == 0 else str(index))
        cells[0].addprevious(serial_cell)
        value = "编号" if index == 0 else "\n".join(
            numbers[label] for label in schema["groups"][index - 1]["labels"])
        _put(cells[schema["quantity_column"]], value)
        if index:
            for height in _xp(row, "./w:trPr/w:trHeight"):
                if height.get(qn("w:hRule")) == "exact":
                    height.set(qn("w:hRule"), "atLeast")
    for index in excluded_columns:
        grid[index].getparent().remove(grid[index])
    for row in rows:
        cells = _cells(row)
        for index in excluded_columns:
            row.remove(cells[index + 1])
    grid = [grid[index] for index in included]
    widths = [widths[index] for index in included]
    headers = _cells(rows[0])
    for cell in headers:
        _join_sample_header_lines(cell)
    minimums = [_sample_header_min_width(outline, table, cell) for cell in headers]
    number_column = included.index(schema["quantity_column"]) + 1
    for row, group in zip(rows[1:], schema["groups"]):
        probe = deepcopy(_cells(row)[number_column])
        for label in group["labels"]:
            _put(probe, numbers[label])
            minimums[number_column] = max(
                minimums[number_column], _sample_header_min_width(outline, table, probe))
    if sum(minimums) > total:
        raise ValueError("样机表原总宽不足以保持表头及每台编号单行；"
                         "请在基本信息总览的「样机信息表」中减少勾选的列。"
                         "序号、准确度等级、编号固定保留；不应缩短真实编号。")
    # Each number keeps its own line; Q3 still only reserves the wider of its two lines.
    weights = [0] + [max(0, width - minimum) for width, minimum in zip(widths, minimums[1:])]
    if not sum(weights):
        weights = [1] * len(minimums)
    remaining = total - sum(minimums)
    adjusted = [minimum + int(remaining * weight / sum(weights))
                for minimum, weight in zip(minimums, weights)]
    adjusted[-1] += total - sum(adjusted)
    for col, width in zip([serial_grid] + grid, adjusted):
        col.set(qn("w:w"), str(width))
    table.tblPr.get_or_add_tblLayout().set(qn("w:type"), "fixed")
    table.tblPr.find(qn("w:tblW")).set(qn("w:type"), "dxa")
    table.tblPr.find(qn("w:tblW")).set(qn("w:w"), str(total))
    for row in rows:
        for cell, width in zip(_cells(row), adjusted):
            cell.get_or_add_tcPr().get_or_add_tcW().set(qn("w:w"), str(width))
            cell.tcPr.tcW.set(qn("w:type"), "dxa")
    return _source_fragment(outline, [table])


def _dates(text):
    match = re.search(r"检测日期\s*[：:]\s*(.+)", text)
    if not match:
        return "", ""
    parts = re.split(r"[~～至]", match.group(1).strip(), maxsplit=1)
    return parts[0].strip(), parts[1].strip() if len(parts) > 1 else parts[0].strip()


def outline_test_basis(doc, source):
    body = doc.element.body
    headings = (
        re.compile(r"^(?:2[.．、]?)?技术依据[：:]?$"),
        re.compile(r"^(?:3[.．、]?)?提供审查的技术文件[：:]?$"),
    )
    anchors = []
    for heading in headings:
        matches = [block for block in body if block.tag == qn("w:p")
                   and heading.fullmatch(re.sub(r"\s+", "", _text(block)))]
        if len(matches) != 1:
            raise ValueError("试验依据：未唯一定位大纲第2项技术依据及第3项提供审查的技术文件，请核对章节。")
        anchors.append(body.index(matches[0]))
    start, end = anchors
    if start >= end:
        raise ValueError("试验依据：大纲第2、3项顺序错误，不能确定完整依据范围。")
    lines = []
    for block in list(body)[start + 1:end]:
        if block.tag != qn("w:p") or _xp(
                block, ".//w:drawing | .//w:pict | .//w:fldChar | .//w:altChunk | .//w:del | .//w:ins"):
            raise ValueError("试验依据：大纲技术依据含不支持的非纯文字结构，不能截取部分内容，请核对。")
        lines.extend(line.strip() for line in Paragraph(block, doc).text.splitlines() if line.strip())
    if not lines:
        raise ValueError("试验依据：大纲第2项技术依据下未找到标准，请补录。")
    return {"value": "\n".join(lines),
            "source": f"{source} · 2.技术依据至3.提供审查的技术文件之前 · 正文块{start + 2}—{end}"}


def sync_test_basis(task, outline, source="当前大纲"):
    task = deepcopy(task)
    field = next((field for field in task.get("overview", []) if field["key"] == "test_basis"), None)
    if field is None:
        field = {"key": "test_basis", "label": "试验依据", "value": "", "candidates": []}
        task.setdefault("overview", []).append(field)
    task["test_basis_issues"] = []
    field["candidates"] = []
    if outline is None:
        task["test_basis_issues"].append("试验依据：请上传并选择本次大纲，以读取第2项技术依据。")
    else:
        try:
            field["candidates"] = [outline_test_basis(outline, source)]
        except ValueError as exc:
            task["test_basis_issues"].append(str(exc))
    previous = field.get("value", "")
    if not field.get("manual_override"):
        field["value"] = field["candidates"][0]["value"] if field["candidates"] else ""
    if previous != field.get("value", ""):
        field["confirmed"] = False
    return task


def outline_production_basis(doc, source):
    paragraphs = list(_xp(doc.element.body, ".//w:p"))
    heading = re.compile(r"^\s*(?:3\s*[.．]\s*2\s*)?产品标准\s*")
    matches = [(index, heading.match(Paragraph(p, doc).text))
               for index, p in enumerate(paragraphs)]
    matches = [(index, match) for index, match in matches if match]
    if len(matches) != 1:
        raise ValueError("生产依据：未唯一定位大纲3.2产品标准，请核对章节。")
    index, match = matches[0]
    text = Paragraph(paragraphs[index], doc).text[match.end():]
    for paragraph in paragraphs[index + 1:]:
        following = Paragraph(paragraph, doc).text
        if re.match(r"^\s*\d+\s*[.．、]\s*\d*\s*\S", following) or re.match(
                r"^\s*(总装图|使用说明书|关键零部件和材料)", following):
            break
        text += "\n" + following
    text = text.lstrip()
    if not text.startswith(("（", "(")):
        raise ValueError("生产依据：大纲3.2产品标准缺少外层括号，请核对原文。")
    stack = []
    closing = {"）": "（", ")": "("}
    end = None
    for position, character in enumerate(text):
        if character in "（(":
            stack.append(character)
        elif character in closing:
            if not stack or stack.pop() != closing[character]:
                raise ValueError("生产依据：大纲3.2产品标准括号不匹配，请核对原文。")
            if not stack:
                end = position
                break
    if end is None:
        raise ValueError("生产依据：大纲3.2产品标准括号未闭合，不能截取部分标准。")
    content = text[1:end]
    # Keep punctuation and inner parentheses; only introduce paragraph boundaries.
    lines, current, depth = [], "", 0
    for character in content:
        if character in "（(":
            depth += 1
        elif character in "）)":
            depth -= 1
        current += character
        if character == "\n" or (character in "；;" and depth == 0):
            if current.strip():
                lines.append(current.strip())
            current = ""
    if current.strip():
        lines.append(current.strip())
    if not lines:
        raise ValueError("生产依据：大纲3.2产品标准括号内为空，请补录。")
    return {"value": "\n".join(lines),
            "source": f"{source} · 3.2 产品标准 · 段落{index + 1}起（括号内原文）"}


def _record_sections(doc, source):
    """Group rows by record markers, retaining original table/row coordinates."""
    kind, rows = "unmapped", []
    table_index = 0
    for block_index, element in enumerate(doc.element.body, 1):
        if element.tag == qn("w:p"):
            start, end = _dates(_text(element))
            if start:
                if rows:
                    yield {"kind": kind, "rows": rows}
                kind, rows = "unmapped", []
                yield {"kind": "date", "start": start, "end": end,
                       "source": "%s · 正文块%d" % (source, block_index)}
            continue
        if element.tag != qn("w:tbl"):
            continue
        table_index += 1
        for ri, (cells, owners) in enumerate(_effective_rows(element), 1):
            texts = [_text(cell) for cell in cells]
            labels = {_norm(text) for text in texts}
            location = "%s · 表%d/行%d" % (source, table_index, ri)
            start, end = _dates(" ".join(texts))
            if start:
                if rows:
                    yield {"kind": kind, "rows": rows}
                kind, rows = "unmapped", []
                yield {"kind": "date", "start": start, "end": end, "source": location}
                continue
            marker = "sample" if "样品名称" in labels else "measurement" if "测量数据" in labels else ""
            if marker:
                if rows:
                    yield {"kind": kind, "rows": rows}
                kind, rows = marker, []
            if any(texts) or any(_text(cell) for cell in owners):
                rows.append({"cells": cells, "owners": owners, "source": location,
                             "table": table_index, "row": ri})
    if rows:
        yield {"kind": kind, "rows": rows}


def parse_upload(path, kind):
    from atmospheric_pressure import extract_pressure

    doc = _load(path)
    source = Path(path).name
    fields = _field_candidates(doc, source, record=kind == "record")
    result = {"fields": fields, "records": [], "warnings": [], "outline": {}}
    if kind == "outline":
        fields.pop("test_basis", None)
        try:
            fields["test_basis"] = [outline_test_basis(doc, source)]
        except ValueError as exc:
            result["warnings"].append(str(exc))
        fields.pop("production_basis", None)
        try:
            fields["production_basis"] = [outline_production_basis(doc, source)]
        except ValueError as exc:
            result["warnings"].append(str(exc))
        components, coverage = _outline_regions(doc)
        result["outline"] = {"components": _preview(components), "coverage": _preview(coverage)}
        result["sample_schema"] = outline_samples(doc)[0]
        return result
    if kind != "record":
        raise ValueError("导入类型必须为 record 或 outline")
    context = {}
    pending = []
    sheet = 0
    for section in _record_sections(doc, source):
        if section["kind"] == "date":
            for entry in pending:
                entry["start"], entry["end"] = section["start"], section["end"]
                entry.setdefault("provenance", {})["time"] = section["source"]
            if not pending and context:
                context["start"], context["end"] = section["start"], section["end"]
                context["provenance"]["time"] = section["source"]
            pending = []
            continue
        rows = section["rows"]
        if section["kind"] == "sample":
            sheet += 1
            context = {"sample_id": "", "model": "", "certificate": "", "temperature": "", "humidity": "",
                       "temperature_unit": "", "humidity_unit": "", "start": "", "end": "", "provenance": {}}
            pending = []
            for row in rows:
                cells = row["cells"]
                context.setdefault("atmospheric_pressure_candidates", []).extend(
                    extract_pressure(" ".join(_text(cell) for cell in row["owners"]), row["source"]))
                for ci, cell in enumerate(cells[:-1]):
                    label = _norm(_text(cell))
                    key = {"型号/规格": "model", "规格型号": "model", "编号": "sample_id",
                           "样品编号": "sample_id", "证书编号": "certificate"}.get(label)
                    if key:
                        context[key] = _text(cells[ci + 1])
                        context["provenance"][key] = "%s/列%d" % (row["source"], ci + 2)
                if "环境条件" in "".join(_text(cell) for cell in cells):
                    for ci, cell in enumerate(cells[:-1]):
                        key = {"温度": "temperature", "湿度": "humidity"}.get(_norm(_text(cell)))
                        if key:
                            context[key] = _text(cells[ci + 1])
                            if ci + 2 < len(cells):
                                context[key + "_unit"] = _text(cells[ci + 2])
                            context["provenance"][key] = "%s/列%d" % (row["source"], ci + 2)
            continue
        if section["kind"] != "measurement":
            result["warnings"].append("%s：未映射的记录内容，需人工核对" % rows[0]["source"])
            continue
        if not context.get("sample_id"):
            raise ValueError("测量数据表缺少前置的样品身份信息")
        title_cells = [_text(c) for c in rows[0]["cells"] if _text(c) and _norm(_text(c)) != "测量数据"]
        if len(set(title_cells)) != 1:
            raise ValueError("测量数据表项目名称不唯一")
        title = title_cells[0]
        names = (["准确度等级和最大允许误差", "示值误差曲线", "重复性"]
                 if title == "示值误差重复性" else [_project_name(title)])
        if len(rows) < 3:
            raise ValueError("测量数据表缺少表头或数据")
        headers = [_text(c) for c in rows[1]["owners"]]
        record_source = "%s · 第%d张记录单 / 表%d" % (source, sheet, rows[0]["table"])
        if rows[0]["row"] != 1:
            record_source += "/行%d" % rows[0]["row"]
        data_rows = [row for row in rows[2:] if [_text(c) for c in row["owners"]] != headers]
        requirement_columns = [i for i, h in enumerate(headers) if _norm(h) in ("技术要求", "试验条件", "子试验")]
        result_columns = [i for i, h in enumerate(headers) if _norm(h) in ("试验结果", "实测结果", "检测结果")]
        if len(requirement_columns) == len(result_columns) == 1:
            measurements = []
            for row in data_rows:
                owners = row["owners"]
                if len(owners) != len(headers):
                    raise ValueError("测量数据列数与表头不一致")
                requirement = _text(owners[requirement_columns[0]])
                value = _text(owners[result_columns[0]])
                if value.strip() and value != "/":
                    measurements.append({
                        "label": requirement, "requirement": requirement, "value": value,
                        "source": "%s/列%d" % (row["source"], result_columns[0] + 1),
                        "flow": "", "metric": headers[result_columns[0]],
                    })
            for name in names:
                entry = dict(deepcopy(context), project=name, rows=deepcopy(measurements),
                             source=record_source)
                result["records"].append(entry)
                pending.append(entry)
            continue
        if "流量点" not in headers or not any("示值误差" in h for h in headers):
            raise ValueError("不支持的测量数据表头")
        measurements = []
        repeats = {}
        seen_summary = set()
        for row in data_rows:
            owners = row["owners"]
            if len(owners) != len(headers):
                raise ValueError("测量数据列数与表头不一致")
            flow = _text(owners[0])
            repeats[flow] = repeats.get(flow, 0) + 1
            repeat = repeats[flow]
            conditions = {header: _text(cell) for header, cell in zip(headers, owners)
                          if "工作压力" in header or "工作温度" in header}
            for ci, (header, cell) in enumerate(zip(headers, owners)):
                value = _text(cell)
                if ci == 0 or not value or value == "/":
                    continue
                summary = "平均示值误差" in header or "重复性" in header
                identity = (header, id(cell))
                if summary and identity in seen_summary:
                    continue
                if summary:
                    seen_summary.add(identity)
                location = "%s/列%d" % (row["source"], ci + 1)
                label = "%s | %s%s" % (flow, "" if summary else "测量%d | " % repeat, header)
                measurements.append({"label": label, "value": value, "source": location,
                                     "flow": flow, "metric": header, "repeat": repeat,
                                     "conditions": conditions})
        for name in names:
            entry = dict(deepcopy(context), project=name, rows=deepcopy(measurements),
                         source=record_source)
            result["records"].append(entry)
            pending.append(entry)
    if not result["records"]:
        raise ValueError("未找到支持的样品测量记录")
    missing = sum(not entry["start"] for entry in result["records"])
    if missing:
        result["warnings"].append("%d组项目/样机记录没有检测日期，待补录" % missing)
    if not fields.get("report_number"):
        result["warnings"].append("未识别到原始记录的报告编号，请核对记录右上方并手动填写。")
    elif len({entry["value"] for entry in fields["report_number"]}) > 1:
        result["warnings"].append("原始记录各处报告编号不一致，请在总览核对全部候选。")
    result["warnings"].append("报告编号读取记录的报告编号栏或样品信息首行右侧证书编号栏，不使用样品/设备编号；工作温度不等同环境温度。")
    return result


def _flow_for_requirement(requirement):
    pieces = [piece.strip() for piece in requirement.split("|")]
    return next((p for p in reversed(pieces) if re.fullmatch(r"Q[1234]|0[.．](?:7|35)[（(]Q3\+Q2[）)]", p)), "")


def sample_record_candidates(schema, parsed, match_model=False):
    """Propose within-specification record order, without assigning identities."""
    candidates, issues = {}, []
    groups = schema.get("groups", [])
    headers = schema.get("headers", [])
    specifications = []
    for group in groups:
        text = " ".join(v for h, v in zip(headers, group["values"]) if "规格" in h or "型号" in h)
        specifications.append(set(re.findall(r"DN\s*(\d+)", text, re.I)))
    distinct = {}
    for data in parsed:
        for record in data.get("records", []):
            key = (record["sample_id"], record.get("model", ""), record.get("certificate", ""))
            if key not in distinct:
                distinct[key] = {k: record.get(k, "") for k in ("sample_id", "model", "certificate", "source")}
    for record in distinct.values():
        dn = set(re.findall(r"DN\s*(\d+)", record["model"], re.I))
        matches = [i for i, spec in enumerate(specifications) if dn and dn == spec]
        if not dn:
            matches = [i for i, g in enumerate(groups) if any(
                _norm(v) == _norm(record["model"]) for h, v in zip(headers, g["values"]) if "型号" in h)]
        if match_model:
            model = _norm(re.sub(r"DN\s*\d+", "", record["model"], flags=re.I))
            matches = [i for i in matches if any(
                _norm(re.sub(r"DN\s*\d+", "", value, flags=re.I)) == model
                for header, value in zip(headers, groups[i]["values"]) if "型号" in header)
                or not any("型号" in header for header in headers)]
        if len(matches) != 1:
            issues.append("原始编号 %s（%s）不能唯一对应大纲规格组，请手工核对编号。" %
                          (record["sample_id"], record["model"]))
            continue
        for label in groups[matches[0]]["labels"]:
            candidates.setdefault(label, []).append(record)
    scopes = {r["certificate"] for r in distinct.values() if r["certificate"]}
    if len(scopes) > 1:
        issues.append("所选原始记录包含不同证书编号，不能自动混为同批样机；请核对资料范围后逐项选择。")
    return candidates, list(dict.fromkeys(issues))


def fill_sample_numbers(task, parsed, *, match_model=False, roster_issues=None):
    """Fill unedited slots from current evidence; never re-sort manual mappings."""
    candidates, issues = sample_record_candidates(task.get("sample_schema", {}), parsed, match_model=match_model)
    samples = task.get("samples", [])
    by_label = {sample["label"]: sample for sample in samples}
    previous = {sample["label"]: sample.get("number", "").strip() for sample in samples}
    flags = {sample["number"]: (sample.get("disassembly", False), sample.get("sealed", False))
             for sample in samples if sample.get("number", "").strip()}
    manual = set()
    for sample in samples:
        value = sample.get("number", "").strip()
        origin = sample.get("number_origin", "")
        automatic = origin == "auto" and value == sample.get("auto_number", "")
        if origin == "manual" or (value and not automatic and value != sample["label"]):
            manual.add(sample["label"])
            if origin == "auto" and not automatic:
                task["samples_confirmed"] = False
            if origin != "manual" and value != sample.get("auto_number", ""):
                sample["number_source"] = ""
            sample.update(number_origin="manual", auto_number="")
    reserved = {sample["number"].strip() for sample in samples
                if sample["label"] in manual and sample.get("number", "").strip()}
    blocked = bool(issues)
    for group in task.get("sample_schema", {}).get("groups", []):
        labels = group["labels"]
        choices = candidates.get(labels[0], []) if labels else []
        old_choices = task.get("sample_candidates", {}).get(labels[0], []) if labels else []
        same_candidates = bool(old_choices) and [
            (c["sample_id"], c.get("model", ""), c.get("certificate", "")) for c in old_choices
        ] == [(c["sample_id"], c.get("model", ""), c.get("certificate", "")) for c in choices]
        by_number = {choice["sample_id"]: choice for choice in choices}
        stable = {
            label for label in labels if label not in manual and not blocked
            and previous[label] in by_number and previous[label] not in reserved
        }
        used = reserved | {previous[label] for label in stable}
        available = [choice for choice in choices if choice["sample_id"] not in used]
        editable = [label for label in labels if label not in manual and label not in stable]
        overflow = len(available) > len(editable)
        if overflow:
            issues.append("规格组%s的原始编号多于可用位置，不能自动取舍，请在匹配窗口核对。" %
                          (labels[0] if labels else ""))
        next_choice = iter(available if not blocked and not overflow else [])
        for label in labels:
            if label in manual:
                continue
            sample = by_label[label]
            if label in stable:
                choice = by_number[previous[label]]
            elif same_candidates and manual.intersection(labels) and sample.get("number_origin") == "auto" and not previous[label]:
                choice = None
            else:
                choice = next(next_choice, None)
            number = choice["sample_id"] if choice else ""
            sample.update(number=number, number_origin="auto", auto_number=number,
                          number_source=choice["source"] if choice else "")
            if previous[label] and previous[label] != label and previous[label] != number:
                sample["disassembly"], sample["sealed"] = flags.get(number, (False, False))
        if parsed:
            for label in labels:
                if not by_label[label].get("number", "").strip():
                    message = "%s：未识别到可自动填入的样机编号，请核对原始记录或手工填写。" % label
                    if roster_issues is not None and message not in issues:
                        roster_issues.append({"kind": "missing_number", "label": label,
                                              "issue_index": len(dict.fromkeys(issues))})
                    issues.append(message)
    changes = {sample["label"]: sample["number"] for sample in samples
               if previous[sample["label"]] != sample["number"]}
    if changes:
        task["samples_confirmed"] = False
        for field in task.get("overview", []):
            if field["key"] == "sealing":
                field["confirmed"] = False
        for project in task.get("projects", []):
            for row in project["rows"]:
                label = row.get("sample_label")
                if label not in changes:
                    continue
                if _has_text(row.get("value", "")):
                    if row.get("auto_value") and row["value"] == row["auto_value"]:
                        row.update(value="", source="", auto_value="", matched_sample_id="",
                                   match_issue="样机对应已改变，请重新解析本次资料")
                    else:
                        row["matched_sample_id"] = row.get("matched_sample_id") or previous[label]
                        row["match_issue"] = "样机对应已改变，保留的人工结果属于旧编号，请核对后重新保存"
                row["sample_id"] = changes[label]
                project["confirmed"] = False
    task.update(sample_candidates=candidates, sample_mapping_issues=list(dict.fromkeys(issues)))
    return task


def _measured_value(value):
    return value["value"] + ("%" if "%" in value.get("metric", "") and "%" not in value["value"] else "")


def _match_measurement(name, row, records):
    sample = row.get("sample_id", "")
    if not sample:
        return "", "", "待匹配：请先核对报告代号与原始样机编号"
    records = [r for r in records if r["sample_id"] == sample]
    if not records:
        return "", "", "缺少该样机的本项目原始记录：" + sample
    requirement = row.get("requirement", "")
    flow = _flow_for_requirement(requirement)
    values = [(r, v) for r in records for v in r["rows"]]
    averages = [(r, v) for r, v in values if "平均示值误差" in v.get("metric", "")]
    errors = [(r, v) for r, v in values if _norm(v.get("metric", "")) == "示值误差%"]
    if name == "重复性" and flow:
        chosen = [(r, v) for r, v in values if "重复性" in v.get("metric", "")
                  and _norm(v.get("flow", "")) == _norm(flow)]
    elif name in ("准确度等级和最大允许误差", "反向流") and flow:
        chosen = [(r, v) for r, v in averages if _norm(v.get("flow", "")) == _norm(flow)]
        if not chosen:
            chosen = [(r, v) for r, v in errors if _norm(v.get("flow", "")) == _norm(flow)]
    elif name == "示值误差曲线":
        if len(records) != 1:
            return "", "", "同一样机存在多个测量批次，不能混合读取误差曲线"
        flows = [_norm(v.get("flow", "")) for _, v in averages]
        expected = {_norm(f) for f in ("Q1", "Q2", "Q3", "Q4", "0.7（Q3+Q2）", "0.35（Q3+Q2）")}
        if len(averages) != 6 or set(flows) != expected or any(
                not re.fullmatch(r"[+-]?\d+(?:\.\d+)?%?", v["value"]) for _, v in averages):
            return "", "", "需六个流量点各一项明确的平均误差，才能读取误差符号情况"
        numbers = [Decimal(v["value"].rstrip("%")) for _, v in averages]
        if min(numbers) > 0 or max(numbers) < 0:
            winner = averages[min(range(len(numbers)), key=lambda index: abs(numbers[index]))][1]
            return (_measured_value(winner),
                    "；".join(v["source"] for _, v in averages)
                    + "；各流量点平均示值误差同号，取绝对值最小的流量点"
                    + winner["flow"] + "的平均示值误差，保留正负号；选取来源："
                    + winner["source"] + "；不作合格判定", "")
        value = "示值误差符号不相同" if min(numbers) < 0 < max(numbers) else "示值误差含零值，请核对"
        return value, "；".join(v["source"] for _, v in averages) + "；仅汇总误差符号，不作合格判定", ""
    elif name == "水温":
        if len(records) != 1 or len(errors) != 2:
            return "", "", "水温低温/最高温度需唯一记录中的前后两次误差，当前数量或批次不唯一"
        index = 1 if "最高允许温度" in requirement else 0 if "10℃" in requirement else None
        if index is None:
            return "", "", "未识别水温子试验条件，需人工核对"
        r, v = errors[index]
        return _measured_value(v), v["source"] + "；按已确认前低温/后最高温度顺序；" + str(v.get("conditions", {})), ""
    elif name == "水压":
        low = "0.03MPa" in requirement
        maximum = re.search(r"最大允许压力[：:]\s*(\d+(?:\.\d+)?)\s*MPa", requirement)
        if not low and not maximum:
            return "", "", "未识别水压子试验压力，需人工核对"
        target = Decimal("0.03") if low else Decimal(maximum[1])
        lower, upper = (target * Decimal("0.95"), target * Decimal("1.05")) if low else (target * Decimal("0.9"), target)
        chosen = []
        for r, v in errors:
            pressures = [text for header, text in v.get("conditions", {}).items() if "工作压力" in header and "MPa" in header]
            if len(pressures) == 1 and re.fullmatch(r"[+-]?\d+(?:\.\d+)?", pressures[0]):
                if lower <= Decimal(pressures[0]) <= upper:
                    chosen.append((r, v))
    elif name in ("静磁场", "安装条件") and "示值误差" in requirement:
        if len(records) != 1:
            return "", "", "同一样机存在多个测量批次，不能混合取三次最大绝对误差"
        if name == "安装条件":
            if len(errors) != 6 or not any(k in requirement for k in ("上游", "下游")):
                return "", "", "安装条件需前后各三次误差分别对应上游/下游，当前无法唯一分组"
            chosen = errors[:3] if "上游" in requirement else errors[3:]
        else:
            chosen = errors
        numeric = [
            re.sub(r"[（(](?:左旋|右旋|不对称速度剖面)[）)]$", "", v["value"]).strip()
            if name == "安装条件" else v["value"]
            for _, v in chosen
        ]
        if len(chosen) != 3 or any(not re.fullmatch(r"[+-]?\d+(?:\.\d+)?%?", value) for value in numeric):
            return "", "", "需要三次明确的示值误差，当前记录不足或格式无法识别"
        winner = max(numeric, key=lambda value: abs(Decimal(value.rstrip("%"))))
        if len({value for value in numeric if abs(Decimal(value.rstrip("%"))) ==
                abs(Decimal(winner.rstrip("%")))}) > 1:
            return "", "", "最大绝对误差并列且符号/数值表达不同，请人工核对"
        measurement = dict(chosen[numeric.index(winner)][1], value=winner)
        return _measured_value(measurement), "；".join(v["source"] for _, v in chosen) + "；三次误差取绝对值最大的一次，保留正负号", ""
    else:
        # Explicit requirement/result records can be used for any project, never inferred pass/fail.
        chosen = [(r, v) for r, v in values if v.get("requirement") and
                  _norm(v["requirement"]) == _norm(requirement)]
        if not chosen:
            return "", "", "记录中未找到可直接对应的结果/定性结论，需补充该子试验记录或人工核对"
    if len(chosen) == 1:
        r, v = chosen[0]
        return _measured_value(v), v.get("source", r["source"]), ""
    if not chosen:
        return "", "", "缺少该样机对应流量点/试验条件的实测结果"
    return "", "", "匹配多个测量批次或结果，不能唯一录入：" + "；".join(v.get("source", r["source"]) for r, v in chosen)


def _record_timestamp(value, end=False):
    day = parse_date(value)
    clock = re.search(r"[ T](\d{2}):(\d{2})(?::(\d{2}))?$", value.strip())
    if clock:
        instant = datetime.combine(day, time(*(int(v or 0) for v in clock.groups())))
        return instant, instant.isoformat(sep=" ", timespec="seconds" if clock[3] else "minutes")
    return datetime.combine(day, time.max if end else time.min), day.isoformat()


def _record_conditions(records):
    values, issues = {}, []
    for key, title in (("start", "开始时间"), ("end", "结束时间"),
                       ("temperature", "环境温度"), ("humidity", "环境湿度")):
        bounds = []
        for record in records:
            try:
                if not record.get(key, "").strip():
                    raise ValueError("原记录未填写")
                if key in ("start", "end"):
                    start = _record_timestamp(record.get("start", ""))
                    end = _record_timestamp(record.get("end", ""), end=True)
                    if start[0] > end[0]:
                        raise ValueError("开始时间晚于结束时间")
                    bounds.append(start if key == "start" else end)
                else:
                    bounds.append(_environment_bounds(record[key], record.get(key + "_unit", ""), key))
            except ValueError as exc:
                issues.append(f"{record['source']}：{title}无法完整汇总（{exc}），请核对原记录。")
        if len(bounds) != len(records):
            continue
        if key in ("start", "end"):
            values[key] = (min if key == "start" else max)(bounds, key=lambda item: item[0])[1]
        else:
            low, high = min(v[0] for v in bounds), max(v[1] for v in bounds)
            def display(value):
                return format(value.normalize(), "f") if value else "0"
            values[key] = display(low) + ("～" + display(high) if high != low else "") + (
                " ℃" if key == "temperature" else " %RH")
    return values, list(dict.fromkeys(issues))


def apply_records(project, parsed):
    result = deepcopy(project)
    result["confirmed"] = False
    records = [r for data in parsed for r in data.get("records", [])
               if _project_name(r.get("project", "")) == _project_name(project["name"])]
    old_summary = result.get("record_summary", {})
    result["record_summary"] = {"values": {}, "records": [], "issues": []}
    notes = result.setdefault("notes", [])
    notes[:] = [note for note in notes if note != "未找到原始记录，待手工录入"
                and not note.startswith(("缺失原始记录：", "请确认样机代号映射", "原始逐次测量已保留"))
                and "存在逐样机/批次差异" not in note and "匹配多个测量批次，待确认" not in note]
    for row in result["rows"]:
        if row.get("auto_value") and row.get("value") == row["auto_value"]:
            row.update(value="", source="", auto_value="", matched_sample_id="")
        row["match_issue"] = ""
    for key, value in old_summary.get("values", {}).items():
        if result.get(key) == value:
            result[key] = ""
    if not records:
        note = "缺失原始记录：" + project["name"] + "（本次所选文件中没有该项目）"
        if note not in notes:
            notes.append(note)
        result["source"] = ""
        for row in result["rows"]:
            row["match_issue"] = note
        return result
    sample_ids = list(dict.fromkeys(r["sample_id"] for r in records))
    if any(not row.get("sample_id") for row in result["rows"]):
        note = "请确认样机代号映射；原始编号：" + "、".join(sample_ids)
        if note not in notes:
            notes.append(note)
    values, issues = _record_conditions(records)
    result.update(values)
    result["record_summary"] = {
        "values": values, "issues": issues,
        "records": [{key: record.get(key, "") for key in
                     ("source", "sample_id", "model", "start", "end", "temperature", "humidity")}
                    for record in records],
    }
    for evidence, record in zip(result["record_summary"]["records"], records):
        for key in ("temperature", "humidity"):
            evidence[key] = " ".join(filter(None, (record.get(key, ""), record.get(key + "_unit", ""))))
    for row in result["rows"]:
        if _has_text(row.get("value", "")):
            if row.get("matched_sample_id") and row["matched_sample_id"] != row.get("sample_id"):
                row["match_issue"] = "样机对应已改变，保留的人工结果属于旧编号，请核对后重新保存"
            continue
        value, source, issue = _match_measurement(_project_name(project["name"]), row, records)
        row["match_issue"] = issue
        if value:
            row.update(value=value, source=source, auto_value=value, matched_sample_id=row["sample_id"])
    result["source"] = "；".join(dict.fromkeys(r["source"] for r in records))
    if any(not row.get("value") for row in result["rows"]):
        note = "原始逐次测量已保留；未录入原因见实测结果窗口逐行提示。静磁场/安装条件按已确认规则取三次绝对值最大的一次并保留符号，不自动推断定性结论。"
        if note not in notes:
            notes.append(note)
    return result


def order_projects(root, items_asset_id, projects):
    """Order by semantic items rows; preserve selection order among variants."""
    rows = section_rows(root, items_asset_id)
    report_type = _asset(root, items_asset_id, "items")["report_type"]
    if report_type == "燃气表":
        from gas_records import order_projects as order_gas_projects

        return order_gas_projects(rows, projects)
    positions = {}
    gas_groups = {
        "低温贮存、高温贮存、恒定湿热贮存、振动": "贮存环境性能",
        "静电放电抗扰度、射频电磁场辐射抗扰度": "电磁兼容",
    } if report_type == "流量计（气体）" else {}
    for project in projects:
        original_name = _project_name(project["name"])
        if original_name in positions:
            continue
        name = gas_groups.get(original_name, original_name)
        matches = [i for i, row in enumerate(rows)
                   if _norm(name) in [_norm(part) for part in row["label"].split(" / ")]]
        if not matches:
            raise ValueError("项目未在本次一览表中找到：" + name)
        if original_name != name and {
                _norm(rows[i]["label"].split(" / ")[-1]) for i in matches
        } != {_norm(part) for part in original_name.split("、")}:
            raise ValueError("一览表项目匹配歧义：" + original_name)
        if len(matches) > 1:
            subtrials = {rows[i]["label"].split(" / ")[-1] for i in matches}
            flow_group = False
            if name in ("贮存环境性能", "电磁兼容") and report_type in FLOWMETER_TYPES:
                labels = [rows[i]["label"].split(" / ") for i in matches]
                children = [_norm(" / ".join(parts[1:])) for parts in labels]
                flow_group = (
                    matches == list(range(matches[0], matches[-1] + 1))
                    and all(_norm(parts[0]) == _norm(name) for parts in labels)
                    and all(children) and len(set(children)) == len(children)
                )
            if not (flow_group or (name == "耐久性" and len(matches) == 2
                    and subtrials == {"断续流量试验", "连续流量试验"})):
                raise ValueError("一览表项目匹配歧义：" + name)
        positions[original_name] = min(matches)
    return sorted(deepcopy(projects), key=lambda p: positions[_project_name(p["name"])])


def _validate_rows(actual, submitted, kind):
    expected = {row["id"] for row in actual}
    identifiers = [row["id"] for row in submitted]
    if len(set(identifiers)) != len(identifiers) or set(identifiers) != expected:
        raise ValueError(kind + "行标识/数量与素材不一致，请重新载入并确认")
    return {row["id"]: row for row in submitted}


def reconcile_result_rows(actual, submitted, kind):
    """Remap physical addresses only when the full ordered semantic schema agrees."""
    identifiers = [row["id"] for row in submitted]
    expected_ids = [row["id"] for row in actual]
    if (len(actual) != len(submitted) or len(set(identifiers)) != len(identifiers)
            or len(set(expected_ids)) != len(expected_ids)
            or any(not re.fullmatch(r"t\d+:r\d+:c\d+", value) for value in identifiers)):
        raise ValueError(kind + "：结果行与素材不一致，不能截断、追加或丢弃未映射数据。")
    for source, previous in zip(actual, submitted):
        if (source["sample_label"] != previous.get("sample_label")
                or " ".join(source["requirement"].split())
                != " ".join(previous.get("requirement", "").split())):
            raise ValueError(kind + "：结果行与素材的样机、技术要求或顺序不一致，不能自动迁移实测值。")
    return [{**deepcopy(previous), "id": source["id"]}
            for source, previous in zip(actual, submitted)]


def _replace_result_field_text(element, value):
    """Keep unchanged characters in their original runs, including separators."""
    from difflib import SequenceMatcher

    original = "".join(_xp(element, ".//w:t/text()"))
    for operation, start, end, new_start, new_end in reversed(
            SequenceMatcher(None, original, value, autojunk=False).get_opcodes()):
        if operation == "equal":
            continue
        replacement = value[new_start:new_end]
        nodes = _xp(element, ".//w:t")
        offset = 0
        inserted = False
        for node in nodes:
            text = node.text or ""
            stop = offset + len(text)
            if start == end and offset <= start <= stop and not inserted:
                _write_slot_text(node, text[:start - offset] + replacement + text[start - offset:])
                inserted = True
            elif stop > start and offset < end:
                left = text[:max(0, start - offset)]
                right = text[max(0, end - offset):]
                _write_slot_text(node, left + (replacement if not inserted else "") + right)
                inserted = True
            offset = stop
        if not nodes:
            _put_if_filled(element, value)


def _result_date_parts(value):
    date = re.fullmatch(
        r"(?P<year>\d{4})(?P<first>\s*[-/.年]\s*)(?P<month>\d{1,2})"
        r"(?P<second>\s*[-/.月]\s*)(?P<day>\d{1,2})(?P<tail>\s*日)?"
        r"(?:(?P<gap>(?(tail)[ T\s]*|[ T\s]+))(?P<time>.+))?", value.strip())
    if not date:
        return None
    clock = None
    if date["time"]:
        clock = re.fullmatch(
            r"(?P<hour>\d{1,2})(?P<separator>\s*[时:：]\s*)(?P<minute>\d{1,2})"
            r"(?P<tail>\s*分)?(?P<seconds>(?P<second_separator>\s*[:：]?\s*)"
            r"(?P<second>\d{1,2})(?P<second_tail>\s*秒)?)?", date["time"])
        if not clock:
            return None
        chinese = "时" in clock["separator"]
        if chinese != bool(clock["tail"]):
            return None
        if clock["second"] and (
                chinese != bool(clock["second_tail"]) or
                (not chinese and not clock["second_separator"].strip())):
            return None
    try:
        datetime(int(date["year"]), int(date["month"]), int(date["day"]),
                 int(clock["hour"]) if clock else 0,
                 int(clock["minute"]) if clock else 0,
                 int(clock["second"] or 0) if clock else 0)
    except ValueError:
        return None
    return date, clock


def _format_result_date(source, value):
    entered = _result_date_parts(value)
    template = _result_date_parts(source)
    if not entered or not template:
        return value
    date, clock = entered
    old_date, old_clock = template
    rendered = "".join((
        str(int(date["year"])).zfill(len(old_date["year"])), old_date["first"],
        str(int(date["month"])).zfill(len(old_date["month"])), old_date["second"],
        str(int(date["day"])).zfill(len(old_date["day"])), old_date["tail"] or "",
    ))
    if not clock:
        return rendered
    if old_clock:
        rendered += old_date["gap"]
        rendered += str(int(clock["hour"])).zfill(len(old_clock["hour"]))
        rendered += old_clock["separator"]
        rendered += str(int(clock["minute"])).zfill(len(old_clock["minute"]))
        rendered += old_clock["tail"] or ""
        if clock["second"] is not None:
            chinese = "时" in old_clock["separator"]
            rendered += (old_clock["second_separator"] if old_clock["second"] is not None
                         else "" if chinese else old_clock["separator"])
            rendered += str(int(clock["second"])).zfill(len(old_clock["second"] or "00"))
            rendered += (old_clock["second_tail"] or "") if old_clock["second"] else (
                "秒" if chinese else "")
    else:
        # No source time is invented; retain all time precision actually supplied.
        rendered += date["gap"] + date["time"]
    return rendered


def _replace_result_date(paragraph, value):
    if not _has_text(value):
        return
    text = "".join(_xp(paragraph, ".//w:t/text()"))
    marker = re.search(r"试验的(?:开始|结束)时间", text)
    if not marker:
        raise ValueError("未找到模板日期字段，无法保留格式替换：" + text)
    separator = re.match(r"\s*[：:]\s*", text[marker.end():])
    spacing = re.match(r"\s*", text[marker.end():])
    start = marker.end() + (separator.end() if separator else spacing.end())
    end = max(start, len(text.rstrip()))
    replacement = _format_result_date(text[start:end], str(value).strip())
    _replace_result_field_text(paragraph, text[:start] + replacement + text[end:])


def _result_environment_parts(value, label):
    number = r"[+\-−]?(?:\d+(?:\.\d+)?|\.\d+)"
    unit = r"(?:℃|°C)" if label == "环境温度" else r"(?:%RH|%)"
    match = re.fullmatch(
        r"(?P<leading>\s*)(?P<open>[（(]?)(?P<inner>\s*)(?P<first>" + number + r")"
        r"(?:(?P<separator>\s*(?:~|～|至|—|–|-|－)\s*)(?P<second>" + number + r"))?"
        r"(?P<inner_end>\s*)(?P<close>[）)]?)(?P<unit_space>\s*)(?P<unit>" + unit
        + r")?(?P<trailing>\s*)", value, flags=re.I)
    if not match or (match["open"], match["close"]) not in (("", ""), ("（", "）"), ("(", ")")):
        return None
    return match


def _format_result_environment(source, value, label, separate_unit):
    entered = _result_environment_parts(value, label)
    if not entered:
        return value
    template = _result_environment_parts(source, label)
    is_range = entered["second"] is not None
    if not template and not is_range:
        template = entered
    if not template:
        if separate_unit and entered["unit"]:
            return value[:entered.start("unit")].rstrip() + value[entered.end("unit"):]
        return value
    rendered = template["leading"] + (template["open"] if is_range else "") + template["inner"] + entered["first"]
    if is_range:
        rendered += (template["separator"] or entered["separator"]) + entered["second"]
    rendered += template["inner_end"] + (template["close"] if is_range else "")
    unit = "" if separate_unit else (template["unit"] or entered["unit"] or "")
    rendered += template["unit_space"] + unit + template["trailing"]
    return rendered


def _number_ranges(numbers, separator="～"):
    numbers = sorted(set(numbers))
    if not numbers:
        return ""
    runs = []
    first = last = numbers[0]
    for number in numbers[1:]:
        if number == last + 1:
            last = number
        else:
            runs.append(str(first) if first == last else f"{first}{separator}{last}")
            first = last = number
    runs.append(str(first) if first == last else f"{first}{separator}{last}")
    return "、".join(runs)


def _fill_project(doc, project, mapping, equipment_numbers):
    from water_results import output_schema, project_schema
    from gas_flowmeter_excel import fill_pressure_requirements

    fill_pressure_requirements(doc, project)
    schema = project_schema(doc, project)
    rows = _validate_rows(schema, project.get("rows", []), project["name"])
    if project.get("water_output_labels") is not None:
        schema = output_schema(doc, project)
    for row in schema:
        submitted = rows[row["id"]]
        if submitted.get("sample_label") != row["sample_label"]:
            raise ValueError("样机代号标签被改动：" + row["id"])
        _put_if_filled(row["_value_cell"], submitted.get("value", ""))
        sample = submitted.get("sample_id", "")
        if not _has_text(sample):
            sample = mapping.get(row["sample_label"], "")
        _put_if_filled(row["_sample_cell"], sample)
    references = []
    for equipment_id in project.get("equipment_ids", []):
        if equipment_id not in equipment_numbers:
            raise ValueError("项目使用的设备未列入本次设备清单")
        references.append(equipment_numbers[equipment_id])
    for paragraph in _xp(doc.element.body, ".//w:p"):
        text = _text(paragraph)
        for marker, value in (
            ("试验的开始时间", project.get("start", "")),
            ("试验的结束时间", project.get("end", "")),
            ("所用试验设备信息见", _number_ranges(references)),
        ):
            if marker in text:
                if marker in ("试验的开始时间", "试验的结束时间"):
                    _replace_result_date(paragraph, value)
                else:
                    _replace_paragraph_value(paragraph, value, after_label=True)
                break
    replacements = {"环境温度": project.get("temperature", ""),
                    "相对湿度": project.get("humidity", ""),
                    "评价人员": "、".join(p for p in project.get("evaluators", []) if _has_text(p)),
                    "复核人员": "、".join(p for p in project.get("reviewers", []) if _has_text(p))}
    for table in _tables(doc):
        for row in _rows(table):
            cells = _cells(row)
            for i, cell in enumerate(cells[:-1]):
                if _norm(_text(cell)) in replacements:
                    label = _norm(_text(cell))
                    value = replacements[label]
                    if label in ("环境温度", "相对湿度") and _has_text(value):
                        unit = _norm(_text(cells[i + 2])).upper() if i + 2 < len(cells) else ""
                        separate_unit = unit in (("℃", "°C") if label == "环境温度" else ("%", "%RH"))
                        target = cells[i + 1]
                        source = "".join(_xp(target, ".//w:t/text()"))
                        _replace_result_field_text(
                            target, _format_result_environment(source, str(value), label, separate_unit))
                    else:
                        _put_if_filled(cells[i + 1], value)
    for paragraph in list(_xp(doc.element.body, "./w:p")):
        if _norm(_text(paragraph)) == _norm("四、型式评价结果摘要"):
            doc.element.body.remove(paragraph)
    _flow_project_layout(doc)


def _paragraph_setting(paragraph, name, enabled):
    properties = paragraph.find(qn("w:pPr"))
    if properties is None:
        properties = OxmlElement("w:pPr")
        paragraph.insert(0, properties)
    # python-docx inserts these properties in the required OOXML child order.
    setting = getattr(properties, "get_or_add_" + name)()
    setting.set(qn("w:val"), "true" if enabled else "false")


def _start_continuation_table(table):
    previous = table.getprevious()
    blank = previous
    while blank is not None and blank.tag == qn("w:p") and not _text(blank):
        if _xp(blank, ".//w:drawing | .//w:pict | .//w:object | .//w:fldChar | "
                      ".//w:fldSimple | .//w:instrText | .//w:sectPr"):
            break
        if _xp(blank, ".//w:br[@w:type='page'] | "
                      "./w:pPr/w:pageBreakBefore[not(@w:val) or @w:val='true' or "
                      "@w:val='1' or @w:val='on']"):
            return
        blank = blank.getprevious()
    if previous is not None and _outline_spacing_paragraph(previous):
        page = previous
    else:
        page = OxmlElement("w:p")
        table.addprevious(page)
    _page_start(page)
    spacing = page.get_or_add_pPr().get_or_add_spacing()
    for key, value in (("before", "0"), ("after", "0"), ("line", "1"), ("lineRule", "exact")):
        spacing.set(qn("w:" + key), value)
    return page


def _result_header_row(row):
    return [_norm(_text(cell)) for cell in _cells(row)[:2]] == ["序号", "项目名称"]


def _split_repeated_project_headers(doc):
    """Return only separators generated here, excluding existing source breaks."""
    generated_separators = []
    for table in list(_tables(doc)):
        rows = _rows(table)
        starts = [0]
        for index, row in enumerate(rows[1:], 1):
            if _result_header_row(row):
                starts.append(index)
        if len(starts) == 1:
            continue
        chunks = []
        for start, end in zip(starts, [*starts[1:], len(rows)]):
            chunk = deepcopy(table)
            for row in _rows(chunk):
                chunk.remove(row)
            for row in rows[start:end]:
                chunk.append(deepcopy(row))
            chunks.append(chunk)
        for index, chunk in enumerate(chunks):
            table.addprevious(chunk)
            if index:
                separator = _start_continuation_table(chunk)
                if separator is not None:
                    generated_separators.append(separator)
        table.getparent().remove(table)
    seen_header = False
    for block in list(doc.element.body):
        if block.tag == qn("w:p") and re.match(
                r"^型式评价(?:结果|原始记录)\s*[-—－–]", _text(block)):
            seen_header = False
        elif block.tag == qn("w:tbl"):
            rows = _rows(block)
            if rows and _result_header_row(rows[0]):
                if seen_header:
                    separator = _start_continuation_table(block)
                    if separator is not None:
                        generated_separators.append(separator)
                seen_header = True
    return generated_separators


def _unpack_metadata_groups(doc):
    groups = []
    for outer in _tables(doc):
        if _xp(outer, "./w:tblPr/w:tblCaption/@w:val") != ["report-project-metadata-group"]:
            continue
        rows = _rows(outer)
        cells = _cells(rows[0]) if len(rows) == 1 else []
        inner = _xp(outer, "./w:tr/w:tc/w:tbl")
        floating = outer.tblPr.find(qn("w:tblpPr"))
        if (len(cells) != 1 or len(inner) != 1 or len(_rows(inner[0])) != 3
                or floating is None or
                _xp(inner[0], "./w:tblPr/w:tblCaption/@w:val") != ["report-project-metadata"]):
            raise ValueError("项目页底包装结构不明确，无法安全读取独立的设备、环境及人员行。")
        if inner[0].tblPr.find(qn("w:tblpPr")) is not None:
            raise ValueError("项目页底存在重复浮动定位，不能安全解包。")
        paragraphs = [deepcopy(p) for p in _xp(cells[0], "./w:p")]
        for paragraph in paragraphs:
            for setting in _xp(paragraph, "./w:pPr/w:pageBreakBefore[@w:val='false' or @w:val='0' or @w:val='off']"):
                setting.getparent().remove(setting)
        if (any(child.tag not in {qn("w:tcPr"), qn("w:tbl"), qn("w:p")} for child in cells[0])
                or any(not _outline_spacing_paragraph(p) for p in paragraphs)):
            raise ValueError("项目页底包装含额外内容，不能解包时丢弃。")
        groups.append((outer, inner[0], floating))
    for outer, table, floating in groups:
        table.tblPr.insert(0, deepcopy(floating))
        outer.addprevious(table)
        outer.getparent().remove(outer)


def _unpack_metadata_equipment(table):
    rows = _rows(table)
    if rows and "所用试验设备信息见" in _text(rows[0]):
        paragraphs = _xp(rows[0], "./w:tc/w:p")
        if len(paragraphs) != 1 or _xp(rows[0], ".//w:tbl"):
            raise ValueError("页底设备信息行结构不明确，不能重新排版。")
        table.addprevious(paragraphs[0])
        table.remove(rows[0])


def _flow_project_layout(doc):
    """Anchor each project's final metadata block above the page footer."""
    _unpack_metadata_groups(doc)
    body = doc.element.body
    for paragraph in list(_xp(body, "./w:p")):
        if _text(paragraph):
            continue
        protected = _xp(
            paragraph,
            ".//w:drawing | .//w:pict | .//w:object | .//w:fldChar | "
            ".//w:fldSimple | .//w:instrText | .//w:sectPr | .//w:br | "
            ".//w:bookmarkStart | .//w:bookmarkEnd | .//w:sdt | "
            ".//w:commentRangeStart | .//w:commentRangeEnd | .//w:commentReference | "
            ".//w:footnoteReference | .//w:endnoteReference | ./w:pPr/w:pageBreakBefore",
        )
        if not protected:
            body.remove(paragraph)
    metadata_tables = [
        table for table in _tables(doc)
        if "环境温度" in _text(table) and "评价人员" in _text(table) and "复核人员" in _text(table)
    ]
    for table in metadata_tables:
        _unpack_metadata_equipment(table)
        if _xp(table, ".//w:br[@w:type='page'] | .//w:sectPr"):
            raise ValueError("环境及人员表含内部强制分页/分节，须先确认素材结构")
        equipment = table.getprevious()
        if equipment is None or equipment.tag != qn("w:p") or "所用试验设备信息见" not in _text(equipment):
            raise ValueError("环境及人员表前未找到设备序号段落，无法定位项目页底信息块")
        if table is not metadata_tables[-1]:
            body.remove(equipment)
            body.remove(table)
            continue
        _anchor_metadata(table, equipment)
        caption = table.tblPr.find(qn("w:tblCaption"))
        if caption is None:
            caption = OxmlElement("w:tblCaption")
            table.tblPr.append(caption)
        caption.set(qn("w:val"), "report-project-metadata")
        rows = _rows(table)
        for index, row in enumerate(rows):
            properties = row.find(qn("w:trPr"))
            if properties is None:
                properties = OxmlElement("w:trPr")
                row.insert(0, properties)
            no_split = properties.find(qn("w:cantSplit"))
            if no_split is None:
                no_split = OxmlElement("w:cantSplit")
                properties.insert_element_before(
                    no_split, "w:trHeight", "w:tblHeader", "w:tblCellSpacing",
                    "w:jc", "w:hidden", "w:ins", "w:del", "w:trPrChange",
                )
            no_split.set(qn("w:val"), "true")
            height = properties.get_or_add_trHeight()
            height.set(qn("w:val"), str(PROJECT_METADATA_ROW_HEIGHT_TWIPS))
            height.set(qn("w:hRule"), "exact")
            for paragraph in _xp(row, ".//w:p"):
                _paragraph_setting(paragraph, "pageBreakBefore", False)
                _paragraph_setting(paragraph, "keepNext", index < len(rows) - 1)
        previous = table.getprevious()
        while previous is not None and previous.tag == qn("w:p"):
            text = _text(previous)
            if not any(marker in text for marker in (
                "所用试验设备信息见", "试验过程中的异常情况记录", "本试验项目的结论"
            )):
                break
            _paragraph_setting(previous, "keepNext", True)
            previous = previous.getprevious()
    if (len(body) > 1 and body[-1].tag == qn("w:sectPr") and body[-2].tag == qn("w:tbl")
            and not _xp(body[-2], "./w:tblPr/w:tblpPr")):
        # Word's paragraph after a final table is retained as a minimal terminator,
        # not a stack of fixed-height padding lines.
        paragraph = OxmlElement("w:p")
        properties = OxmlElement("w:pPr")
        spacing = OxmlElement("w:spacing")
        for key, value in (("before", "0"), ("after", "0"), ("line", "20"), ("lineRule", "exact")):
            spacing.set(qn("w:" + key), value)
        properties.append(spacing)
        paragraph.append(properties)
        body.insert(len(body) - 1, paragraph)


def _anchor_metadata(table, equipment):
    grid = _xp(table, "./w:tblGrid/w:gridCol")
    width = sum(int(column.get(qn("w:w"))) for column in grid)
    if not grid or width <= 0:
        raise ValueError("环境及人员表缺少有效列宽，无法固定页底版式")
    properties = table.tblPr
    floating = OxmlElement("w:tblpPr")
    for name, value in (("vertAnchor", "margin"), ("horzAnchor", "margin"),
                        ("tblpYSpec", "bottom"), ("tblpX", "0"),
                        ("topFromText", "120"), ("bottomFromText", "0"),
                        ("leftFromText", "0"), ("rightFromText", "0")):
        floating.set(qn("w:" + name), value)
    for old in list(properties.findall(qn("w:tblpPr"))):
        properties.remove(old)
    properties.insert(1 if len(properties) and properties[0].tag == qn("w:tblStyle") else 0, floating)
    overlap = properties.find(qn("w:tblOverlap"))
    if overlap is None:
        overlap = OxmlElement("w:tblOverlap")
        properties.insert(properties.index(floating) + 1, overlap)
    overlap.set(qn("w:val"), "never")
    table_width = properties.find(qn("w:tblW"))
    if table_width is None:
        table_width = OxmlElement("w:tblW")
        properties.insert_element_before(table_width, "w:jc", "w:tblCellSpacing", "w:tblInd",
                                         "w:tblBorders", "w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook")
    table_width.set(qn("w:type"), "dxa")
    table_width.set(qn("w:w"), str(width))
    labels = {"环境温度", "相对湿度", "评价人员", "复核人员"}
    for row in _rows(table):
        cells = _cells(row)
        for index, cell in enumerate(cells):
            label = _norm(_text(cell))
            _fit_label_cell(cell, label in labels)
            if label in labels or (index and _norm(_text(cells[index - 1])) in {"评价人员", "复核人员"}):
                _left_metadata_cell(cell)
    row = OxmlElement("w:tr")
    cell = OxmlElement("w:tc")
    prop = OxmlElement("w:tcPr")
    cell_width = OxmlElement("w:tcW")
    cell_width.set(qn("w:w"), str(width))
    cell_width.set(qn("w:type"), "dxa")
    span = OxmlElement("w:gridSpan")
    span.set(qn("w:val"), str(len(grid)))
    prop.extend([cell_width, span])
    cell.append(prop)
    cell.append(equipment)
    _fit_label_cell(cell, False)
    _left_metadata_cell(cell)
    row.append(cell)
    table.insert(table.index(_rows(table)[0]), row)


def _left_metadata_cell(cell):
    for paragraph in _xp(cell, "./w:p"):
        properties = paragraph.get_or_add_pPr()
        properties.get_or_add_jc().set(qn("w:val"), "left")
        indent = properties.get_or_add_ind()
        indent.attrib.clear()
        for name in ("left", "right", "firstLine"):
            indent.set(qn("w:" + name), "0")


def _format_liquid_metadata(doc):
    tables = [table for table in _tables(doc)
              if all(label in _text(table) for label in ("环境温度", "评价人员", "复核人员"))]
    if not tables:
        raise ValueError("水表/流量计结果素材缺少页底温湿度及人员表")
    for table in tables:
        rows = _rows(table)
        if len(rows) == 3 and "所用试验设备信息见" in _text(rows[0]):
            equipment = _xp(rows[0], ".//w:p")
            content = rows[1:]
        elif len(rows) == 2:
            previous = table.getprevious()
            if previous is None or previous.tag != qn("w:p") or "所用试验设备信息见" not in _text(previous):
                raise ValueError("水表/流量计页底表前未找到设备引用段落")
            equipment, content = [previous], rows
        else:
            raise ValueError("水表/流量计页底表结构不明确，不能调整文字排版")
        if len(equipment) != 1:
            raise ValueError("水表/流量计设备引用段落不唯一")
        spacing = equipment[0].get_or_add_pPr().get_or_add_spacing()
        spacing.set(qn("w:line"), "360")
        spacing.set(qn("w:lineRule"), "auto")
        for row in content:
            for cell in _cells(row):
                cell.get_or_add_tcPr().get_or_add_vAlign().set(qn("w:val"), "center")
            for paragraph in _xp(row, ".//w:p"):
                properties = paragraph.get_or_add_pPr()
                properties.get_or_add_jc().set(qn("w:val"), "center")
                spacing = properties.get_or_add_spacing()
                for key, value in (("line", "240"), ("lineRule", "auto"),
                                   ("before", "0"), ("after", "0"),
                                   ("beforeLines", "0"), ("afterLines", "0"),
                                   ("beforeAutospacing", "0"), ("afterAutospacing", "0")):
                    spacing.set(qn("w:" + key), value)


def _fill_production_basis(doc, value):
    if not _has_text(value):
        return
    body = doc.element.body
    label = _unique_anchor(body, "申请型式评价的计量器具其生产依据为")
    ending = _unique_anchor(body, "样机型号、规格、准确度等级/最大允许误差/不确定度及编号")
    start, stop = body.index(label), body.index(ending)
    if stop <= start:
        raise ValueError("生产依据目标区域顺序错误")
    previous = list(body)[start + 1:stop]
    if any(p.tag != qn("w:p") or _xp(p, ".//w:sectPr | .//w:drawing | .//w:fldChar")
           for p in previous):
        raise ValueError("生产依据目标区域包含不支持的结构，不能覆盖")
    prototype = previous[0] if previous else label
    raw = "".join(_xp(label, ".//w:t/text()"))
    colon = re.search(r"[：:]", raw)
    if not colon:
        raise ValueError("生产依据引导句缺少冒号")
    for old in previous:
        body.remove(old)
    # Clear legacy inline values without changing the template's label or its runs.
    for node in _xp(label, ".//w:br"):
        node.getparent().remove(node)
    nodes = _xp(label, ".//w:t")
    offset = 0
    for node in nodes:
        text = node.text or ""
        node.text = text[:max(0, colon.end() - offset)]
        offset += len(text)
    for line in (line.strip() for line in value.splitlines() if line.strip()):
        paragraph = OxmlElement("w:p")
        if prototype.find(qn("w:pPr")) is not None:
            paragraph.append(deepcopy(prototype.find(qn("w:pPr"))))
        _put(paragraph, line)
        run_style = prototype.find("./" + qn("w:pPr") + "/" + qn("w:rPr"))
        if run_style is None:
            styles = _xp(label, "./w:r/w:rPr")
            run_style = styles[0] if styles else None
        if run_style is not None:
            paragraph.find(qn("w:r")).insert(0, deepcopy(run_style))
        properties = paragraph.get_or_add_pPr()
        indent = properties.get_or_add_ind()
        for name in ("firstLine", "hanging", "hangingChars"):
            indent.attrib.pop(qn("w:" + name), None)
        indent.set(qn("w:firstLineChars"), "200")
        body.insert(body.index(ending), paragraph)


def _report_test_basis_paragraphs(doc):
    body = doc.element.body
    start = _unique_anchor(body, "型式评价的依据")
    end = _unique_anchor(body, "型式评价所用的仪器设备一览表")
    if body.index(start) >= body.index(end):
        raise ValueError("试验依据：报告第三部分的目标边界顺序错误。")
    paragraphs = list(body)[body.index(start) + 1:body.index(end)]
    if not any(_text(p) for p in paragraphs) or any(
            p.tag != qn("w:p") or _xp(p, ".//w:sectPr | .//w:drawing | .//w:pict | .//w:fldChar | .//w:bookmarkStart")
            for p in paragraphs):
        raise ValueError("试验依据：报告第三部分缺少依据段落或含不支持的结构，不能覆盖。")
    return paragraphs


def _fill_test_basis(doc, value):
    if not _has_text(value):
        return
    lines = [line.strip() for line in value.splitlines() if line.strip()]
    body = doc.element.body
    previous = _report_test_basis_paragraphs(doc)
    start = _unique_anchor(body, "审查的技术资料及结论")
    end = _unique_anchor(body, "型式评价结论及建议")
    candidates = [p for p in list(body)[body.index(start) + 1:body.index(end)]
                  if p.tag == qn("w:p") and "申请单位提交的" in _text(p)]
    if len(candidates) != 1:
        raise ValueError("试验依据：技术资料审查段落缺失或不唯一。")
    review = candidates[0]
    raw = "".join(_xp(review, ".//w:t/text()"))
    matches = list(re.finditer(r"符合(.+?)的要求", raw))
    if len(matches) != 1:
        raise ValueError("试验依据：审查段落的“符合……的要求”引用位置缺失或不唯一。")
    prototype = next(p for p in previous if _text(p))
    position = body.index(previous[0])
    for index, line in enumerate(lines):
        paragraph = deepcopy(prototype)
        _replace_paragraph_value(paragraph, line)
        body.insert(position + index, paragraph)
    for paragraph in previous:
        body.remove(paragraph)
    _replace_text_span(review, *matches[0].span(1), "、".join(lines))


def _fill_review_model(doc, model, *, instrument_name=None):
    issues = []
    if not _has_text(model):
        issues.append("技术资料审查段落：本次大纲封面型号缺失或不唯一，系列保留原文，请核对。")
    if instrument_name is not None and not _has_text(instrument_name):
        issues.append("技术资料审查段落：本次计量器具名称未填写，名称保留原文，请核对。")
    if not _has_text(model) and not _has_text(instrument_name):
        return issues
    body = doc.element.body
    start = _unique_anchor(body, "审查的技术资料及结论")
    end = _unique_anchor(body, "型式评价结论及建议")
    candidates = [p for p in list(body)[body.index(start) + 1:body.index(end)]
                  if p.tag == qn("w:p") and "申请单位提交的" in _text(p)]
    if len(candidates) != 1:
        raise ValueError("技术资料审查段落缺失或不唯一，无法替换本次系列和计量器具名称。")
    paragraph = candidates[0]
    text = "".join(_xp(paragraph, ".//w:t/text()"))
    matches = list(re.finditer(
        r"申请单位提交的\s*(.+?系列)(\s*)(.+?)(\s*)(?=的技术(?:材料|资料))", text))
    if len(matches) != 1:
        raise ValueError("技术资料审查段落的系列和计量器具名称位置缺失或不唯一，请核对模板。")
    match = matches[0]
    if _has_text(instrument_name):
        _replace_text_span(paragraph, *match.span(3), instrument_name.strip())
    if _has_text(model):
        _replace_text_span(paragraph, *match.span(1), model.strip())
    return issues


def _fill_other_information(doc, task, outline=None):
    from supplementary_notes import (
        NUMBER, clause_number, clause_order, outline_note_ids, read_supplementary_notes, renumber_note, strip_number,
    )

    first = next((f.get("value", "") for f in task.get("overview", []) if f["key"] == "sealing"), "")
    body = doc.element.body
    heading = _unique_anchor(body, "其他说明")
    ending = _unique_anchor(body, "九、签发")
    if body.index(ending) <= body.index(heading):
        raise ValueError("其他说明章节边界顺序错误。")
    blocks = list(body)[body.index(heading) + 1:body.index(ending)]
    existing = {}
    for block in blocks:
        if block.tag != qn("w:p") or _xp(block, ".//w:sectPr | .//w:drawing | .//w:fldChar"):
            raise ValueError("其他说明区域包含不支持的结构，不能覆盖。")
        text = _text(block)
        if not text:
            continue
        match = NUMBER.match(text)
        if not match or int(match[1] or match[2]) in existing:
            raise ValueError("其他说明条目编号缺失或重复，不能可靠替换。")
        existing[int(match[1] or match[2])] = block
    templates = [block for block in existing.values() if "本院已对" in _text(block) and "封样" in _text(block)]
    if not templates and _has_text(first):
        templates = [block for block in existing.values() if strip_number(_text(block)) == strip_number(first.strip())]
    if len(templates) != 1:
        raise ValueError("其他说明缺少唯一封样信息模板。")
    prototype = deepcopy(templates[0])
    seal_text = strip_number(first.strip() if _has_text(first) else _text(prototype))
    parsed = read_supplementary_notes(outline) if outline is not None else None
    source_ids = outline_note_ids(task, parsed["entries"]) if parsed is not None else task.get("supplementary_note_ids", [])
    if parsed is None and not source_ids:
        source_ids = [f"outline:current:{index}" for index, _ in enumerate(task.get("supplementary_notes", []), 1)]
    order = clause_order(task, source_ids)
    source_entries = dict(zip(source_ids, parsed["entries"])) if parsed is not None else {}
    source_texts = dict(zip(source_ids, task.get("supplementary_notes", [])))
    custom = {note["id"]: note["text"] for note in task.get("custom_supplementary_notes", [])}
    for block in blocks:
        body.remove(block)
    composer = Composer(doc)
    for number, key in enumerate(order, 1):
        if key in source_entries:
            entry = source_entries[key]
            fragment = _source_fragment(outline, [entry["element"]])
            paragraph = list(fragment.element.body)[0]
            if entry["prefix"]:
                for numbering in _xp(paragraph, "./w:pPr/w:numPr"):
                    numbering.getparent().remove(numbering)
                properties = paragraph.get_or_add_pPr()
                numbering = properties.get_or_add_numPr()
                numbering.get_or_add_numId().val = 0
                run = OxmlElement("w:r")
                styles = _xp(paragraph, "./w:r/w:rPr")
                if styles:
                    run.append(deepcopy(styles[0]))
                Run(run, None).text = renumber_note(entry["prefix"], number)
                paragraph.insert(list(paragraph).index(properties) + 1, run)
            else:
                match = NUMBER.match(Paragraph(paragraph, None).text)
                if match is None:
                    raise ValueError("补充说明条款号缺失，无法重新编号。")
                _replace_text_span(paragraph, *match.span(1 if match[1] is not None else 2),
                                   clause_number(match[1] or match[2], number))
            _insert(composer, body.index(ending), fragment)
        else:
            paragraph = deepcopy(prototype)
            text = seal_text if key == "sealing" else source_texts.get(key, custom.get(key, ""))
            if key in custom:
                text = strip_number(text).strip() or "待补录"
            _put(paragraph, renumber_note(text, number))
            body.insert(body.index(ending), paragraph)


def _fill_basic(doc, overview, mapping, master=False, classification_code=""):
    values = {field["key"]: field.get("value", "") for field in overview
              if field["key"] not in TEMPLATE_DEFAULT_FIELDS
              and not (master and field["key"] in SAMPLE_OVERVIEW_FIELDS)}
    values["applicant"] = values.get("manufacturer", "")
    if master:
        values["classification_code"] = resolve_classification({
            "overview": overview, "classification_code": classification_code,
        }) or "待补录"
        _fill_production_basis(doc, values.get("production_basis", ""))
    aliases = {_norm(label): key for key, labels in FIELDS.items() for label in labels}
    for table in _tables(doc):
        rows = _rows(table)
        if not rows:
            continue
        head = _norm(_text(rows[0]))
        if master and "样机型号" in head and "编号" in head:
            continue
        if master and "测量参数" in head and "计量性能" in head:
            if len(rows) > 1:
                cells = _cells(rows[1])
                if len(cells) != 6:
                    raise ValueError("计量参数表结构不支持")
                parameter_values = [values.get(key, "") for key in (
                    "measurement_parameter", "parameter_unit", "measurement_range",
                    "display_digits", "metrological_performance")]
                for cell, value in zip(cells[1:], parameter_values):
                    _put_if_filled(cell, value)
            continue
        if master and _text(_cells(rows[0])[0]) == "序号":
            # Equipment column labels are not label/value form fields.
            continue
        for row in rows:
            cells = _cells(row)
            for ci, cell in enumerate(cells[:-1]):
                key = aliases.get(_norm(_text(cell)))
                if key and key not in TEMPLATE_DEFAULT_FIELDS:
                    value = values.get(key, "")
                    if key == "sample_number" and not _has_text(value):
                        old_labels = re.findall(r"[A-Z]\d{3}", _text(cells[ci + 1]))
                        value = "、".join(mapping[label] for label in old_labels if label in mapping)
                    _put_if_filled(cells[ci + 1], value)
    for paragraph in _xp(doc.element.body, "./w:p"):
        text = _text(paragraph)
        compact = re.sub(r"\s+", "", text)
        for key, labels in FIELDS.items():
            if key in TEMPLATE_DEFAULT_FIELDS:
                continue
            label = next((label for label in labels
                          if compact.startswith(label + "：") or compact.startswith(label + ":")), None)
            if label:
                _replace_paragraph_value(paragraph, values.get(key, ""), after_label=True)
                break
        if not master:
            continue
        if text.startswith("注：编号为"):
            sample = re.search(r"编号为(.+?)的样机", "".join(_xp(paragraph, ".//w:t/text()")))
            if sample:
                _replace_text_span(paragraph, *sample.span(1), values.get("disassembly_sample", ""))
        elif "经审查，申请单位提交" in text:
            _replace_paragraph_value(paragraph, values.get("technical_review", ""))
        elif "建议批准下列型号计量器具的型式" in _norm(text):
            conclusion = values.get("evaluation_conclusion", "")
            raw = "".join(_xp(paragraph, ".//w:t/text()"))
            marker = raw.index("建议批准")
            end = len(raw[:marker].rstrip("，, \t")) if conclusion else marker
            _replace_text_span(paragraph, len(raw) - len(raw.lstrip()), end, conclusion)
        elif "本院已" in text and "封样" in text:
            _replace_paragraph_value(paragraph, values.get("sealing", ""))
        elif compact.startswith("型式评价时间"):
            _replace_paragraph_value(paragraph, values.get("evaluation_period", ""), after_label=True)
        elif "显示型式" in text:
            _replace_paragraph_value(paragraph, values.get("display_type", ""), after_label=True)
        elif compact.startswith("申请型式评价的计量器具其生产依据为"):
            continue
        elif re.match(r"^[1-5][.．、]?\s*(温度|相对湿度|电源|地点|其他)", compact):
            environment = re.match(r"^[1-5][.．、]?(温度|相对湿度|电源|地点|其他)", compact)
            _replace_paragraph_value(paragraph, values.get(ENVIRONMENT_FIELDS[environment.group(1)], ""),
                                     after_label=True, preserve_terminator=environment.group(1) == "电源")
    if master:
        seen_parts = set()
        for section in doc.sections:
            for part in (section.header, section.first_page_header, section.even_page_header,
                         section.footer, section.first_page_footer, section.even_page_footer):
                if part.is_linked_to_previous:
                    continue
                if id(part.part) in seen_parts:
                    continue
                seen_parts.add(id(part.part))
                for paragraph in _xp(part._element, ".//w:p"):
                    text = _text(paragraph)
                    if re.search(r"\d{4}E\d+[-－]\d+[-－]\d+", text):
                        # Only header/footer report-number slots; attachment labels remain.
                        _replace_text_nodes(paragraph, values.get("report_number", ""))
                for field in _xp(part._element, ".//w:fldSimple"):
                    if re.search(r"\b(?:PAGE|NUMPAGES)\b", field.get(qn("w:instr"), "")):
                        field.set(qn("w:dirty"), "true")
                        for text in _xp(field, ".//w:t"):
                            text.text = ""


def _replace_text_nodes(paragraph, replacement):
    nodes = _xp(paragraph, ".//w:t")
    # Header paragraphs contain tabs/positioned runs; replace only number runs.
    old_number = re.search(r"\d{4}E\d+[-－]\d+[-－]\d+", "".join(node.text or "" for node in nodes))
    if not old_number:
        return
    _replace_text_span(paragraph, *old_number.span(), replacement)


def _put_equipment_runs(cell, runs):
    _put(cell, "")
    paragraph = cell.find(qn("w:p"))
    template = paragraph.find(qn("w:r"))
    paragraph.remove(template)
    for content in runs:
        run = deepcopy(template)
        for child in list(run):
            if child.tag != qn("w:rPr"):
                run.remove(child)
        properties = run.get_or_add_rPr()
        properties.get_or_add_i().val = content["italic"]
        properties.get_or_add_iCs().val = content["italic"]
        properties.get_or_add_vertAlign().val = {
            "normal": "baseline", "sup": "superscript", "sub": "subscript",
        }[content["script"]]
        Run(run, None).text = content["text"]
        paragraph.append(run)


def _equipment_performance_layout(cell):
    """Keep the existing 15.6pt equipment line pitch independent of script-font metrics."""
    from negative_numbers import protect_negative_paragraph

    for paragraph in _xp(cell, "./w:p"):
        properties = paragraph.get_or_add_pPr()
        spacing = properties.get_or_add_spacing()
        spacing.attrib.clear()
        for key, value in (("before", "0"), ("after", "0"), ("line", "312"), ("lineRule", "exact")):
            spacing.set(qn("w:" + key), value)
        snap = properties.find(qn("w:snapToGrid"))
        if snap is None:
            snap = OxmlElement("w:snapToGrid")
            properties.insert_element_before(snap, "w:spacing", "w:ind", "w:jc", "w:rPr")
        snap.set(qn("w:val"), "false")
        protect_negative_paragraph(paragraph)


def _set_font_size(element, points):
    for run in _xp(element, ".//w:r"):
        properties = run.get_or_add_rPr()
        properties.get_or_add_sz().val = Pt(points)
        complex_size = properties.find(qn("w:szCs"))
        if complex_size is None:
            complex_size = OxmlElement("w:szCs")
            properties.insert_element_before(
                complex_size, "w:highlight", "w:u", "w:effect", "w:bdr", "w:shd",
                "w:fitText", "w:vertAlign", "w:rtl", "w:cs", "w:em", "w:lang")
        complex_size.set(qn("w:val"), str(int(points * 2)))


def _format_equipment_dates(table):
    """Change calendar labels only; stored dates and validity checks stay intact."""
    for row in _rows(table)[1:]:
        cell = _cells(row)[-1]
        original = _text(cell)
        match = re.fullmatch(r"(\d{4})[-/.年](\d{1,2})[-/.月](\d{1,2})日?", original.strip())
        if match:
            year, month, day = match.groups()
            formatted = f"{year}.{int(month):02d}.{int(day):02d}"
            if original != formatted:
                _put(cell, formatted)


def _fill_equipment(doc, equipment):
    from equipment_typography import normalize_equipment_typography

    equipment = [
        normalize_equipment_typography(
            Equipment.model_validate({"name": "", "number": "", "valid_until": "", **item}).model_dump()
            if item.get("formatting") is not None else item)
        for item in equipment
    ]
    matches = []
    for table in _tables(doc):
        rows = _rows(table)
        if not rows:
            continue
        headers = [_norm(_text(c)) for c in _cells(rows[0])]
        if len(headers) in (4, 6) and headers[0] == "序号" and headers[-1] == "证书有效期" and "编号" in headers:
            matches.append(table)
    if len(matches) != 2:
        raise ValueError("主文/附件设备表缺失或不唯一")
    for table in matches:
        rows = _rows(table)
        if len(rows) < 2:
            raise ValueError("设备表没有可复制的数据行")
        _set_font_size(table, 12)
        if not equipment:
            _format_equipment_dates(table)
            continue
        template = deepcopy(rows[1])
        for row in rows[1:]:
            table.remove(row)
        for index, equipment_row in enumerate(equipment, 1):
            row = deepcopy(rows[index] if index < len(rows) else template)
            cells = _cells(row)
            if index >= len(rows):
                for cell in cells:
                    _put(cell, "")
            fields = ["name"]
            if len(cells) == 6:
                fields += ["model", "performance"]
            fields += ["number", "valid_until"]
            _put(cells[0], str(index))
            formatting = equipment_row.get("formatting") or {}
            for cell, field in zip(cells[1:], fields):
                value = equipment_row.get(field, "")
                if field in formatting and _has_text(value):
                    _put_equipment_runs(cell, formatting[field])
                else:
                    _put_if_filled(cell, value)
                if field == "performance" and _has_text(value):
                    _equipment_performance_layout(cell)
            table.append(row)
        _format_equipment_dates(table)
        _set_font_size(table, 12)


def _source_fragment(doc, elements):
    fragment = deepcopy(doc)
    body = fragment.element.body
    for element in list(body):
        body.remove(element)
    for element in elements:
        clone = deepcopy(element)
        for sect in _xp(clone, ".//w:sectPr"):
            sect.getparent().remove(sect)
        body.append(clone)
    body.append(OxmlElement("w:sectPr"))
    return fragment


def _page_start(paragraph, enabled=True):
    for name, value in (("pageBreakBefore", enabled), ("keepNext", True)):
        _paragraph_setting(paragraph, name, value)


def _explicit_page_start(paragraph):
    _paragraph_setting(paragraph, "pageBreakBefore", False)
    content = _xp(paragraph, ".//w:t[normalize-space(.) != ''] | .//w:br | "
                  ".//w:drawing | .//w:pict | .//w:object | .//w:fldChar")
    if content and content[0].tag == qn("w:br") and content[0].get(qn("w:type")) == "page":
        return
    run, page_break = OxmlElement("w:r"), OxmlElement("w:br")
    page_break.set(qn("w:type"), "page")
    run.append(page_break)
    paragraph.insert(1, run)


def _editable_page_breaks(doc):
    blocks = list(doc.element.body)
    for index, block in enumerate(blocks):
        if block.tag != qn("w:p"):
            continue
        paragraph = Paragraph(block, doc)
        enabled = paragraph.paragraph_format.page_break_before
        style = paragraph.style
        while enabled is None and style is not None:
            enabled = style.paragraph_format.page_break_before
            style = style.base_style
        if not enabled:
            continue
        _paragraph_setting(block, "pageBreakBefore", False)
        if index == 0:
            continue
        previous = blocks[index - 1]
        prior = previous
        while prior is not None and prior.tag == qn("w:p") and not _xp(prior, "./w:pPr/w:sectPr"):
            content = _xp(prior, ".//w:t[normalize-space(.) != ''] | .//w:br | "
                          ".//w:drawing | .//w:pict | .//w:object | .//w:fldChar")
            if content:
                break
            prior = prior.getprevious()
        else:
            content = []
        if content and content[-1].tag == qn("w:br") and content[-1].get(qn("w:type")) == "page":
            continue
        if _xp(previous, "./w:pPr/w:sectPr"):
            # The following section's properties define how that section starts.
            section = next((s for b in blocks[index:] for s in
                            ([b] if b.tag == qn("w:sectPr") else _xp(b, "./w:pPr/w:sectPr"))), None)
            if section is not None:
                start = _xp(section, "./w:type/@w:val")
                if not start or start[0] in ("nextPage", "oddPage", "evenPage"):
                    continue
        _explicit_page_start(block)


PAGE_START_HEADINGS = (
    "样机型号、规格、准确度等级/最大允许误差/不确定度及编号",
    "关键零部件和材料", "型式评价的依据", "型式评价项目及评价结果一览表",
    "审查的技术资料及结论", "九、签发", "三、试验项目记录",
)


def _start_required_sections(doc):
    for label in PAGE_START_HEADINGS:
        heading = _unique_anchor(doc.element.body, label)
        if heading.tag != qn("w:p"):
            raise ValueError("须另起一页的标题不是独立段落：" + label)
        previous = heading.getprevious()
        while previous is not None and previous.tag == qn("w:p") and not _text(previous):
            if _xp(previous, ".//w:sectPr | .//w:drawing | .//w:pict | .//w:object | "
                   ".//w:fldChar | .//w:fldSimple | .//w:instrText | .//w:bookmarkStart | "
                   ".//w:bookmarkEnd | .//w:sdt | .//w:commentReference | "
                   ".//w:commentRangeStart | .//w:commentRangeEnd | "
                   ".//w:footnoteReference | .//w:endnoteReference"):
                break
            earlier = previous.getprevious()
            doc.element.body.remove(previous)
            previous = earlier
        _page_start(heading)
        spacing = heading.get_or_add_pPr().get_or_add_spacing()
        for name in ("before", "beforeLines", "beforeAutospacing"):
            spacing.set(qn("w:" + name), "0")


def _reference_underline_width(cell):
    text = "".join(_xp(cell, ".//w:t/text()"))
    sizes = set(_xp(cell, ".//w:r/w:rPr/w:sz/@w:val"))
    if len(sizes) != 1 or not text or any(c != " " and not "\u4e00" <= c <= "\u9fff" for c in text):
        raise ValueError("委托信息参考下划线须为同字号中文及半角空格，请核对模板")
    # These fixed template lines render CJK glyphs full-width and spaces half-width.
    return sum(1 if c == " " else 2 for c in text) * int(sizes.pop()) * 5


def _template_run_width(run):
    text = "".join(_xp(run, "./w:t/text()"))
    sizes = _xp(run, "./w:rPr/w:sz/@w:val")
    if len(sizes) != 1 or any(c != " " and not c.isascii() and
                              not "\u4e00" <= c <= "\u9fff" and c != "：" for c in text):
        raise ValueError("固定横线参考文字的字号或字符不支持，请核对报告模板")
    if any(c.isascii() and c != " " and c not in "0123456789" for c in text):
        raise ValueError("固定横线参考区须为中文、空格或数字")
    # The template balances spaces to half an em; Arial digits are 1139/2048 em.
    return sum(0.5 if c == " " else 1139 / 2048 if c.isascii() else 1
               for c in text) * int(sizes[0]) * 10


def _reference_line_geometry(paragraph):
    start = width = 0
    found = False
    for run in _xp(paragraph, "./w:r"):
        underlined = _xp(run, "./w:rPr/w:u/@w:val") == ["single"]
        if underlined:
            found = True
            width += _template_run_width(run)
        elif not found:
            start += _template_run_width(run)
        elif _text(run):
            raise ValueError("固定横线参考区后存在额外文字")
    if not found or width <= 0:
        raise ValueError("报告模板缺少参考下划线")
    return round(start), round(width)


def _center_underlined_paragraph(paragraph, width, available):
    properties = paragraph.get_or_add_pPr()
    properties.get_or_add_jc().set(qn("w:val"), "center")
    indent = properties.get_or_add_ind()
    indent.attrib.clear()
    # Word extends paragraph rules 1.5 pt past each text indent.
    for name, value in (("left", 30), ("right", available - width + 30), ("firstLine", 0)):
        indent.set(qn("w:" + name), str(value))
    for underline in _xp(paragraph, ".//w:rPr/w:u"):
        underline.set(qn("w:val"), "none")
    border = properties.find(qn("w:pBdr"))
    if border is None:
        border = OxmlElement("w:pBdr")
        properties.insert_element_before(
            border, "w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku",
            "w:wordWrap", "w:overflowPunct", "w:topLinePunct", "w:autoSpaceDE",
            "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid",
            "w:spacing", "w:ind", "w:contextualSpacing", "w:jc", "w:rPr",
        )
    bottom = border.find(qn("w:bottom"))
    if bottom is None:
        bottom = OxmlElement("w:bottom")
        border.append(bottom)
    for name, value in (("val", "single"), ("sz", "6"), ("space", "0"), ("color", "auto")):
        bottom.set(qn("w:" + name), value)


def _fixed_value_line(doc, paragraph, start, width):
    text = Paragraph(paragraph, None).text
    match = re.search(r"[：:]", text)
    if not match:
        raise ValueError("固定横线字段缺少标签分隔符")
    label, value = text[:match.end()], text[match.end():].strip()
    sizes = set(_xp(paragraph, "./w:r/w:rPr/w:sz/@w:val"))
    pitches = set(_xp(doc.element.body, ".//w:sectPr/w:docGrid/@w:linePitch"))
    if len(sizes) != 1 or len(pitches) != 1:
        raise ValueError("固定横线字段的字号或行网格不一致，请核对模板")
    size, pitch = int(sizes.pop()), int(pitches.pop())
    spacing = paragraph.pPr.find(qn("w:spacing"))
    line_height = size * 12
    if spacing is not None:
        line = int(spacing.get(qn("w:line"), "240"))
        line_height = max(line_height, round(size * 10 * line / 240)
                          if spacing.get(qn("w:lineRule"), "auto") == "auto" else line)
    row_height = ((line_height + pitch - 1) // pitch) * pitch
    table = doc.add_table(rows=1, cols=2)
    table.autofit = False
    properties = table._tbl.tblPr
    properties.find(qn("w:tblW")).set(qn("w:type"), "dxa")
    properties.find(qn("w:tblW")).set(qn("w:w"), str(start + width))
    margins = OxmlElement("w:tblCellMar")
    for side in ("top", "left", "bottom", "right"):
        edge = OxmlElement("w:" + side)
        edge.set(qn("w:w"), "0")
        edge.set(qn("w:type"), "dxa")
        margins.append(edge)
    properties.append(margins)
    borders = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        edge = OxmlElement("w:" + side)
        edge.set(qn("w:val"), "nil")
        borders.append(edge)
    properties.insert_element_before(borders, "w:tblLayout", "w:tblCellMar", "w:tblLook")
    for index, (content, cell_width) in enumerate(((label, start), (value, width))):
        table._tbl.tblGrid.gridCol_lst[index].set(qn("w:w"), str(cell_width))
        cell = table.rows[0].cells[index]._tc
        cell.tcPr.tcW.set(qn("w:w"), str(cell_width))
        cell.remove(cell.find(qn("w:p")))
        p = deepcopy(paragraph)
        _put(p, content)
        cell.append(p)
        cell.tcPr.get_or_add_vAlign().set(qn("w:val"), "center")
        snap = p.pPr.find(qn("w:snapToGrid"))
        if snap is None:
            snap = OxmlElement("w:snapToGrid")
            p.pPr.insert_element_before(snap, "w:spacing", "w:ind", "w:contextualSpacing", "w:jc", "w:rPr")
        snap.set(qn("w:val"), "false")
        spacing = p.get_or_add_pPr().get_or_add_spacing()
        for name, setting in (("line", size * 10), ("lineRule", "exact"),
                              ("before", 0), ("after", 0)):
            spacing.set(qn("w:" + name), str(setting))
        if index:
            for number in _xp(p, "./w:pPr/w:numPr"):
                number.getparent().remove(number)
            _center_underlined_paragraph(p, width, width)
    row_properties = table.rows[0]._tr.get_or_add_trPr()
    row_properties.append(OxmlElement("w:cantSplit"))
    height = OxmlElement("w:trHeight")
    height.set(qn("w:val"), str(row_height))
    height.set(qn("w:hRule"), "atLeast")
    row_properties.append(height)
    paragraph.addprevious(table._tbl)
    paragraph.getparent().remove(paragraph)


def _align_cover_and_signing(doc, reference):
    originals = _xp(reference.element.body, "./w:p")
    cover_end = next(i for i, p in enumerate(originals) if _xp(p, "./w:pPr/w:sectPr"))
    classification = next(p for p in originals[:cover_end] if _norm(_text(p)).startswith("分类编码"))
    start, width = _reference_line_geometry(classification)
    for paragraph in list(doc.element.body):
        if _xp(paragraph, "./w:pPr/w:sectPr"):
            break
        if paragraph.tag == qn("w:p") and any(
                _norm(_text(paragraph)).startswith(label) for label in ("器具名称", "分类编码", "报告编号")):
            _fixed_value_line(doc, paragraph, start, width)
    personnel = next(p for p in originals if _norm(_text(p)).startswith("型式评价人员"))
    start, width = _reference_line_geometry(personnel)
    period = next(p for p in doc.element.body if _norm(_text(p)).startswith("型式评价时间"))
    tabs = _xp(period, "./w:pPr/w:tabs/w:tab[@w:val='left']/@w:pos")
    if len(tabs) != 1:
        raise ValueError("型式评价时间的编号缩进无法确定，请核对模板")
    _fixed_value_line(doc, period, start + int(tabs[0]), width)


def _fit_label_cell(cell, no_wrap=True):
    properties = cell.get_or_add_tcPr()
    if no_wrap:
        setting = properties.find(qn("w:noWrap"))
        if setting is None:
            setting = OxmlElement("w:noWrap")
            properties.insert_element_before(setting, "w:tcMar", "w:textDirection", "w:tcFitText", "w:vAlign")
        setting.set(qn("w:val"), "true")
    margins = properties.find(qn("w:tcMar"))
    if margins is None:
        margins = OxmlElement("w:tcMar")
        properties.insert_element_before(margins, "w:textDirection", "w:tcFitText", "w:vAlign")
    for side in ("left", "right"):
        edge = margins.find(qn("w:" + side))
        if edge is None:
            edge = OxmlElement("w:" + side)
            margins.append(edge)
        edge.set(qn("w:w"), "0")
        edge.set(qn("w:type"), "dxa")


def _align_basic_information(doc):
    tables = [t for t in _tables(doc) if "委托单位" in _text(t) and "委托负责人" in _text(t)]
    if len(tables) != 1:
        raise ValueError("主报告申请及委托信息表缺失或不唯一")
    table = tables[0]
    cells = {}
    for row in _rows(table):
        row_cells = _cells(row)
        for index, cell in enumerate(row_cells[:-1]):
            label = _norm(_text(cell))
            if label in ("制造单位", "申请单位", "代理人", "委托日期", "委托单位", "委托负责人"):
                if label in cells:
                    raise ValueError("基本信息字段重复：" + label)
                cells[label] = row_cells[index + 1]
    missing = {"制造单位", "申请单位", "代理人", "委托日期", "委托单位", "委托负责人"} - cells.keys()
    if missing:
        raise ValueError("基本信息字段缺失：" + "、".join(sorted(missing)))
    widths = {_reference_underline_width(cells[label]) for label in ("委托单位", "委托负责人")}
    if len(widths) != 1:
        raise ValueError("委托单位与委托负责人参考下划线宽度不一致")
    width = widths.pop()
    for label in ("制造单位", "申请单位", "代理人", "委托日期"):
        cell = cells[label]
        paragraphs = _xp(cell, "./w:p")
        if len(paragraphs) != 1:
            raise ValueError("基本信息填写区不是单段落：" + label)
        margins = []
        for side in ("left", "right"):
            values = (_xp(cell, f"./w:tcPr/w:tcMar/w:{side}/@w:w")
                      or _xp(table, f"./w:tblPr/w:tblCellMar/w:{side}/@w:w"))
            margins.append(int(values[0]) if values else 108)
        available = int(cell.tcPr.tcW.w) - sum(margins)
        if width > available:
            raise ValueError("基本信息填写区窄于参考下划线：" + label)
        _put(cell, Paragraph(paragraphs[0], None).text.strip())
        _center_underlined_paragraph(paragraphs[0], width, available)


def _isolated_styles(source, namespace, preserve_table_defaults=False):
    source = deepcopy(source)
    styles = source.styles.element
    mapping = {s.get(qn("w:styleId")): namespace + s.get(qn("w:styleId"))
               for s in styles.findall(qn("w:style"))}
    defaults = {s.get(qn("w:type")): s.get(qn("w:styleId"))
                for s in styles.findall(qn("w:style")) if s.get(qn("w:default")) == "1"}
    roots = [source.element, styles]
    roots.extend(rel.target_part.element for rel in source.part.rels.values()
                 if rel.reltype == RT.NUMBERING)
    roots.extend(part.element for part in source.part.related_parts.values() if isinstance(part, HeaderPart))
    for root in roots:
        for paragraph in _xp(root, ".//w:p"):
            props = paragraph.get_or_add_pPr()
            if props.find(qn("w:pStyle")) is None and defaults.get("paragraph"):
                props.get_or_add_pStyle().val = defaults["paragraph"]
        if preserve_table_defaults and defaults.get("table"):
            for table in _xp(root, ".//w:tbl"):
                props = table.tblPr
                if props.find(qn("w:tblStyle")) is None:
                    props.get_or_add_tblStyle().val = defaults["table"]
        for ref in _xp(root, ".//w:pStyle | .//w:rStyle | .//w:tblStyle | .//w:basedOn | .//w:next | .//w:link"):
            old = ref.get(qn("w:val"))
            if old in mapping:
                ref.set(qn("w:val"), mapping[old])
    for style in styles.findall(qn("w:style")):
        style.set(qn("w:styleId"), mapping[style.get(qn("w:styleId"))])
        style.attrib.pop(qn("w:default"), None)
        name = style.find(qn("w:name"))
        if name is not None:
            name.set(qn("w:val"), namespace + name.get(qn("w:val")))
    return source


def _insert(composer, index, source):
    """Import relationships/styles/numbering without touching master sections."""
    composer.reset_reference_mapping()
    composer._create_style_id_mapping(source)
    for element in source.element.body:
        if element.tag == qn("w:sectPr"):
            continue
        clone = deepcopy(element)
        for sect in _xp(clone, ".//w:sectPr"):
            sect.getparent().remove(sect)
        composer.doc.element.body.insert(index, clone)
        composer.add_referenced_parts(source.part, composer.doc.part, clone)
        composer.add_styles(source, clone)
        composer.add_numberings(source, clone)
        composer.add_images(source, clone)
        composer.add_diagrams(source, clone)
        composer.add_shapes(source, clone)
        composer.add_footnotes(source, clone)
        composer.remove_header_and_footer_references(source, clone)
        index += 1
    composer.add_styles_from_other_parts(source)
    composer.renumber_bookmarks()
    composer.renumber_docpr_ids()
    composer.renumber_nvpicpr_ids()
    return index


def _replace_region(composer, start, end, fragments, include_start=False):
    body = composer.doc.element.body
    left = body.index(start) + (0 if include_start else 1)
    right = body.index(end)
    if right < left:
        raise ValueError("报告替换区域边界顺序错误")
    preserved = []
    for element in list(body)[left:right]:
        sections = _xp(element, ".//w:sectPr")
        if sections:
            if element.tag != qn("w:p") or len(sections) != 1:
                raise ValueError("区域边界包含不支持的分节结构")
            anchor = deepcopy(element)
            _put(anchor, "")
            preserved.append(anchor)
        body.remove(element)
    cursor = left
    for fragment in fragments:
        cursor = _insert(composer, cursor, fragment)
    for anchor in preserved:
        body.insert(cursor, anchor)
        cursor += 1


def _normalize_font_tree(root):
    if root.tag == qn("w:styles"):
        defaults = root.find(qn("w:docDefaults"))
        if defaults is None:
            defaults = OxmlElement("w:docDefaults")
            root.insert(0, defaults)
        run_default = defaults.find(qn("w:rPrDefault"))
        if run_default is None:
            run_default = OxmlElement("w:rPrDefault")
            defaults.insert(0, run_default)
        if run_default.find(qn("w:rPr")) is None:
            run_default.append(OxmlElement("w:rPr"))
    for element in _xp(root, ".//w:r | .//w:lvl"):
        if element.find(qn("w:rPr")) is None:
            properties = OxmlElement("w:rPr")
            if element.tag == qn("w:r"):
                element.insert(0, properties)
            else:
                element.append(properties)
    properties = _xp(root, ".//w:rPr")
    for prop in properties:
        fonts = prop.find(qn("w:rFonts"))
        if fonts is None:
            fonts = OxmlElement("w:rFonts")
            prop.insert(1 if len(prop) and prop[0].tag == qn("w:rStyle") else 0, fonts)
        for attribute in ("asciiTheme", "hAnsiTheme", "eastAsiaTheme", "cstheme"):
            fonts.attrib.pop(qn("w:" + attribute), None)
        for attribute, name in (("ascii", "Arial"), ("hAnsi", "Arial"),
                                ("cs", "Arial"), ("eastAsia", "宋体")):
            fonts.set(qn("w:" + attribute), name)
    return bool(properties)


def _normalize_page_numbers(root):
    if root.tag != qn("w:ftr"):
        return
    for paragraph in _xp(root, ".//w:p"):
        codes = _xp(paragraph, ".//w:instrText/text() | .//w:fldSimple/@w:instr")
        if not re.search(r"\b(?:PAGE|NUMPAGES)\b", " ".join(codes), re.I):
            continue
        # Word for Mac can regenerate a simple field's empty result at 9 pt.
        # Use the same explicit instruction/result runs as the working PAGE field.
        for field in _xp(paragraph, ".//w:fldSimple"):
            code = field.get(qn("w:instr"), "")
            if not re.search(r"\b(?:PAGE|NUMPAGES)\b", code, re.I):
                continue
            parent = field.getparent()
            index = parent.index(field)
            elements = []
            for field_type in ("begin", "instruction", "separate"):
                run = OxmlElement("w:r")
                if field_type == "instruction":
                    instruction = OxmlElement("w:instrText")
                    instruction.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
                    instruction.text = code
                    run.append(instruction)
                else:
                    marker = OxmlElement("w:fldChar")
                    marker.set(qn("w:fldCharType"), field_type)
                    if field_type == "begin":
                        marker.set(qn("w:dirty"), "true")
                    run.append(marker)
                elements.append(run)
            elements.extend(deepcopy(child) for child in field)
            end = OxmlElement("w:r")
            marker = OxmlElement("w:fldChar")
            marker.set(qn("w:fldCharType"), "end")
            end.append(marker)
            elements.append(end)
            parent.remove(field)
            for offset, element in enumerate(elements):
                parent.insert(index + offset, element)
        properties = paragraph.get_or_add_pPr()
        if properties.find(qn("w:rPr")) is None:
            properties.append(OxmlElement("w:rPr"))
        for run in _xp(paragraph, ".//w:r"):
            run.get_or_add_rPr()
        for prop in _xp(paragraph, ".//w:rPr"):
            for name in ("sz", "szCs"):
                size = prop.find(qn("w:" + name))
                if size is None:
                    size = OxmlElement("w:" + name)
                    prop.insert_element_before(
                        size, "w:szCs", "w:highlight", "w:u", "w:effect", "w:bdr",
                        "w:shd", "w:fitText", "w:vertAlign", "w:rtl", "w:cs",
                        "w:em", "w:lang", "w:eastAsianLayout", "w:rPrChange",
                    )
                size.set(qn("w:val"), "21")
        for node in _xp(paragraph, ".//w:instrText | .//w:fldSimple"):
            simple = node.tag == qn("w:fldSimple")
            code = node.get(qn("w:instr"), "") if simple else (node.text or "")
            if not re.search(r"\b(?:PAGE|NUMPAGES)\b", code, re.I):
                continue
            code = re.sub(r"\\\*\s+MERGEFORMAT\b", r"\\* CHARFORMAT", code, flags=re.I)
            if not re.search(r"\\\*\s+CHARFORMAT\b", code, re.I):
                code = code.rstrip() + " \\* CHARFORMAT "
            if simple:
                node.set(qn("w:instr"), code)
            else:
                node.text = code


def _validate_attachment_headers(doc):
    headers = {}
    for section in doc.sections:
        for header in (section.header, section.first_page_header, section.even_page_header):
            if header.is_linked_to_previous:
                continue
            identity = str(header.part.partname)
            headers[identity] = header
    pattern = r"^(附件\s*[12])\s+(型式评价记录|型式评价试验样机照片)\s+(报告编号[：:].+?)\s*$"
    for header in headers.values():
        for element in _xp(header._element, ".//w:p"):
            paragraph = Paragraph(element, header)
            text = paragraph.text.strip()
            if re.match(r"附件\s*[12]", text):
                match = re.fullmatch(pattern, text)
                if not match or _xp(element, ".//w:fldChar | .//w:fldSimple | .//w:drawing | .//w:pict"):
                    raise ValueError("附件页眉结构无法可靠识别，未擅自改写附件标识或报告编号。")


def _check_report_font_tree(root, part_name):
    expected = {"ascii": "Arial", "hAnsi": "Arial", "cs": "Arial", "eastAsia": "宋体"}
    for element in _xp(root, ".//w:r | .//w:lvl"):
        if len(_xp(element, "./w:rPr/w:rFonts")) != 1:
            raise ValueError("最终Word字体检查失败（%s）：文字或编号缺少唯一字体设置。" % part_name)
    for fonts in _xp(root, ".//w:rFonts"):
        if (any(fonts.get(qn("w:" + key)) != value for key, value in expected.items())
                or any("theme" in etree.QName(key).localname.lower() for key in fonts.attrib)):
            raise ValueError("最终Word字体检查失败（%s）：须为中文宋体、西文Arial，且不得被主题字体覆盖。" % part_name)


def _normalize_report_fonts(path):
    # Normalize every Word XML part, including generic footnote/endnote parts.
    temporary = path.with_suffix(".fonts.tmp")
    try:
        with ZipFile(path) as source, ZipFile(temporary, "w") as target:
            target.comment = source.comment
            for entry in source.infolist():
                data = source.read(entry.filename)
                if entry.filename.startswith("word/") and entry.filename.endswith(".xml"):
                    root = etree.fromstring(data)
                    if root.tag == qn("w:ftr"):
                        root = parse_xml(data)
                        _normalize_page_numbers(root)
                    if _normalize_font_tree(root):
                        data = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
                    _check_report_font_tree(root, entry.filename)
                target.writestr(entry, data)
        temporary.replace(path)
    finally:
        temporary.unlink(missing_ok=True)


def compose_report(root, task, outline_path: Optional[Path], output, draft, *, observation_builder=None,
                   signature_width_twips=None):
    root = _root(root)
    output = Path(output).resolve()
    readonly = [a["_path"].resolve() for a in _asset_files(root)] + [(root / "report model/report.docx").resolve()]
    if outline_path:
        readonly.append(Path(outline_path).resolve())
    if output in readonly or "report model" in output.parts or "original record" in output.parts:
        raise ValueError("不能覆盖模板、素材或导入原件")
    if not outline_path:
        raise ValueError("生成完整报告必须提供本次大纲")
    outline = _load(outline_path)
    from observation_documents import outline_values
    from result_requirements import fill_result_requirements, sync_project_requirements

    parameters = outline_values(outline, Path(outline_path).name, report_type=task.get("report_type", "水表"))
    components, coverage = _outline_regions(outline)
    schema, sample_table = outline_samples(outline)
    independent_meter = task.get("report_type") in INDEPENDENT_METER_TYPES
    if independent_meter:
        from flowmeter_tasks import reconcile_samples as reconcile_flowmeter_samples

        task = reconcile_flowmeter_samples(task, schema, report_type=task["report_type"])
    else:
        if task.get("report_type") == "燃气表":
            parameters = outline_values(outline, Path(outline_path).name)
        from water_results import reconcile_task

        task = reconcile_task(deepcopy(task), schema)
        schema = task["sample_schema"]
    task = sync_other_information(task, outline, Path(outline_path).name)
    task = sync_test_basis(task, outline, Path(outline_path).name)
    environment = aggregate_project_environment(task.get("projects", []))
    from evaluation_summary import summarize_evaluation

    evaluation = summarize_evaluation(task)
    automatic_fields = REPORT_ENVIRONMENT_FIELDS | AUTOMATIC_OVERVIEW_FIELDS
    task["overview"] = [field for field in task.get("overview", []) if field["key"] not in automatic_fields]
    task["overview"].extend({"key": key, "value": environment["values"].get(key, "")}
                            for key in REPORT_ENVIRONMENT_FIELDS)
    task["overview"].extend({"key": key, "value": evaluation["values"].get(key, "")}
                            for key in AUTOMATIC_OVERVIEW_FIELDS)
    doc = _load(root / "report model/report.docx")
    original_sections = [etree.tostring(s) for s in _xp(doc.element, ".//w:sectPr")]
    projects = order_projects(root, task["items_asset_id"], task.get("projects", []))
    if independent_meter:
        projects = [reconcile_project_rows(root, project) for project in projects]
    if not independent_meter:
        projects = [sync_project_requirements(root, project, parameters) for project in projects]
    main_overview = {field["key"]: field for field in task.get("overview", [])}
    missing_main = [FIELDS[key][0] for key in MAIN_FIELDS
                    if not _has_text(main_overview.get(key, {}).get("value"))]
    if not resolve_classification(task):
        missing_main.append("分类编码")
    if not draft:
        missing = list(task.get("sealing_issues", []))
        missing.extend(task.get("supplementary_issues", []))
        missing.extend(task.get("test_basis_issues", []))
        if not task["samples"] or any(not _has_text(s["number"]) for s in task["samples"]):
            missing.append("样机编号未填写")
        if not any(s["disassembly"] for s in task["samples"]):
            missing.append("未选择拆样样机")
        if task.get("legacy_sample_fields"):
            missing.append("旧样机字段尚未核对")
        missing.extend("主文/" + label for label in missing_main)
        if not projects:
            missing.append("未选择试验项目")
        for project in projects:
            for key in ("start", "end", "temperature", "humidity"):
                if not _has_text(project.get(key)):
                    missing.append(project["name"] + "/" + key)
            if any(not _has_text(row.get("value")) for row in project.get("rows", [])):
                missing.append(project["name"] + "/缺少试验结果")
            evaluators, reviewers = project.get("evaluators", []), project.get("reviewers", [])
            count = len(required_staff_groups(task, project))
            if len(set(evaluators)) != count or len(set(reviewers)) != count or set(evaluators) & set(reviewers):
                missing.append(project["name"] + "/人员配置未完成或角色冲突")
        if missing:
            raise ValueError("正式输出仍有缺项：" + "；".join(missing))
    mapping = {s["label"]: s["number"] for s in task["samples"] if s["number"]}
    for project in projects:
        for row in project.get("rows", []):
            label, sample = row.get("sample_label", ""), row.get("sample_id", "")
            if _has_text(sample):
                if label in mapping and mapping[label] != sample:
                    raise ValueError("同一样机代号映射冲突：" + label)
                mapping[label] = sample
    reverse = {}
    for label, sample in mapping.items():
        if sample in reverse and reverse[sample] != label:
            if not (independent_meter and (label == sample or reverse[sample] == sample)):
                raise ValueError("同一真实样机对应多个报告代号，请确认：" + sample)
        reverse[sample] = label
    equipment = task.get("equipment", [])
    equipment_numbers = {e["id"]: i for i, e in enumerate(equipment, 1)}
    if len(equipment_numbers) != len(equipment):
        raise ValueError("本次设备清单包含重复 ID")
    warnings = [TEMPLATE_FALLBACK_WARNING, *environment["issues"], *evaluation["issues"], *task.get("sealing_issues", [])]
    warnings.extend(note for project in projects for note in project.get("notes", [])
                    if note.startswith("大纲技术参数/"))
    warnings.extend(task.get("supplementary_issues", []))
    warnings.extend(task.get("test_basis_issues", []))
    basis_field = main_overview.get("test_basis", {})
    if (_has_text(basis_field.get("value")) and basis_field.get("candidates")
            and not any(candidate["value"] == basis_field["value"] for candidate in basis_field["candidates"])):
        warnings.append("试验依据与本次大纲第2项技术依据不同，已保留人工填写值，请核对。")
    if not resolve_classification(task):
        warnings.append(CLASSIFICATION_ISSUE)
    if draft:
        warnings.append("本文件为待核对草稿；未填写字段保留原文，不代表本次通过或正式签发。")
    if missing_main:
        warnings.append("主文尚有未填写字段：" + "、".join(missing_main) +
                        "；未填写部分保留原文，须核对；不适用项请明确填写“不适用”。")
    if any(not _has_text(r.get("sample_id")) and r.get("sample_label") not in mapping
           for p in projects for r in p["rows"]):
        if not draft:
            raise ValueError("正式输出存在缺失的样机代号映射")
        warnings.append("存在缺失的样机代号映射，未填写的对应编号保留素材原文，须人工核对。")
    outline_fields = _field_candidates(outline, Path(outline_path).name)
    for field in task.get("overview", []):
        if field["key"] not in ("instrument_name", "manufacturer", "model"):
            continue
        source_values = {entry["value"] for entry in outline_fields.get(field["key"], [])}
        if field.get("value") and source_values and field["value"] not in source_values:
            conflict = "大纲原文%s（%s）与总览（%s）不同；原文复制区域未改写，须核对。" % (
                field.get("label", field["key"]), "；".join(sorted(source_values)), field["value"])
            if not draft:
                raise ValueError(conflict)
            warnings.append(conflict)
    reference = deepcopy(doc)
    _fill_basic(doc, [f for f in task.get("overview", []) if f["key"] != "sealing"], mapping,
                master=True, classification_code=task.get("classification_code", ""))
    from basic_options import align_measurement_table, fill_basic_options
    fill_basic_options(doc, task)
    warnings.extend(_fill_review_model(
        doc, task["sealing_model"], instrument_name=main_overview.get("instrument_name", {}).get("value", "")))
    _fill_test_basis(doc, main_overview.get("test_basis", {}).get("value", ""))
    _fill_other_information(doc, task, outline)
    selected_sample = next((s for s in task["samples"] if s["disassembly"]), None)
    if selected_sample:
        note = _unique_anchor(doc.element.body, "注：编号为")
        match = re.search(r"编号为(.+?)的样机", _text(note))
        if match is None:
            raise ValueError("报告拆样样机说明格式无法定位。")
        _replace_text_span(note, *match.span(1), selected_sample["number"] or selected_sample["label"])
    else:
        warnings.append("未选择拆样样机，报告拆样说明保留模板原文，请核对。")
    if any(not s["number"] for s in task["samples"]):
        warnings.append("大纲样机表展开后，未填写编号处使用对应报告代号，未视为已确认的真实编号。")
    _align_basic_information(doc)
    _fill_equipment(doc, equipment)
    _align_cover_and_signing(doc, reference)
    items = _load(_asset(root, task["items_asset_id"], "items")["_path"])
    observe = _load(_asset(root, task["observe_asset_id"], "observe")["_path"])
    from observation_documents import (copy_warning, fill_observation_identities,
                                       fill_observation_remarks, prepare_observation_export)
    from observation_certificates import fill_certificates

    if observation_builder is None:
        warnings.extend(fill_observation_remarks(observe, outline, task)["issues"])
        if independent_meter:
            warnings.extend(fill_observation_identities(observe, task))
        prepare_observation_export(observe)
    else:
        warnings.extend(observation_builder(observe, outline, task)["warnings"])
    warnings.extend(fill_certificates(observe, outline, task))
    warnings.extend(fill_certificates(items, outline, task, section="项目一览表"))
    observe = _isolated_styles(observe, "reportObserve_", preserve_table_defaults=True)
    if observation_builder is None:
        warnings.append(copy_warning(task.get("report_type", "水表")))
    fragments = []
    for project in projects:
        asset = _asset(root, project["asset_id"], "result")
        if _project_name(project["name"]) != asset["project"]:
            raise ValueError("项目名称与素材不匹配")
        fragment = _load(asset["_path"])
        if not independent_meter:
            warnings.extend(fill_result_requirements(fragment, project["name"], parameters))
        if task.get("report_type") == "燃气表":
            from gas_records import fill_metadata

            fill_metadata(fragment, project)
        _fill_project(fragment, project, mapping, equipment_numbers)
        from signature_templates import apply_signature_templates

        apply_signature_templates(fragment, root, task.get("include_atmospheric_pressure", False),
                                  project=project, equipment_numbers=equipment_numbers,
                                  layout_settings=doc.settings.element, width_twips=signature_width_twips)
        if task.get("report_type") == "燃气表":
            from gas_curves import fill_curves

            warnings.extend(fill_curves(fragment, project, task["projects"]))
        _split_repeated_project_headers(fragment)
        from result_pagination import mark_result_tables

        mark_result_tables(fragment)
        first_heading = next((p for p in _xp(fragment.element.body, "./w:p") if _text(p)), None)
        if first_heading is None:
            raise ValueError("结果素材缺少项目起始标题")
        _page_start(first_heading, enabled=bool(fragments))
        fragments.append(fragment)
    composer = Composer(doc)
    body = doc.element.body
    sample_heading = _unique_anchor(body, "样机型号、规格、准确度等级")
    note = _unique_anchor(body, "注：编号为")
    _replace_region(composer, sample_heading, note,
                    [_sample_table_fragment(outline, schema, sample_table, task["samples"],
                                            task.get("sample_table_excluded_columns", []))])
    inserted_samples = [
        block for block in list(body)[body.index(sample_heading) + 1:body.index(note)]
        if block.tag == qn("w:tbl")]
    if len(inserted_samples) != 1:
        raise ValueError("主文样机表未唯一合成，无法对齐计量器具测量参数表。")
    align_measurement_table(doc, inserted_samples[0])
    result_heading = _unique_anchor(body, "四、型式评价结果摘要")
    _page_start(result_heading)
    regions = [
        (_unique_anchor(body, "关键零部件和材料"), _unique_anchor(body, "型式评价的依据"),
         [_source_fragment(outline, components)], False),
        (_unique_anchor(body, "型式评价项目及评价结果一览表"), _unique_anchor(body, "审查的技术资料及结论"),
         [items], True),
        (_unique_anchor(body, "建议批准下列型号计量器具的型式"), _unique_anchor(body, "其他说明"),
         [_source_fragment(outline, coverage)], False),
        (_unique_anchor(body, "样机的基本信息"), _unique_anchor(body, "三、试验项目记录"),
         [observe], True),
        (result_heading, _unique_anchor(body, "样机的整体外形照片"),
         fragments, False),
    ]
    for start, end, source_docs, include in regions:
        _replace_region(composer, start, end, source_docs, include)
    from photo_documents import extract_component_names, fill_attachment_photos
    photo_components = [
        {"id": f"component:{task.get('outline_id', '')}:{index}", "name": name}
        for index, name in enumerate(extract_component_names(components), 1)
    ]
    warnings.extend(fill_attachment_photos(
        doc, task.get("photos", {}), photo_components, task.get("_photo_paths", {}),
        task.get("photo_edits", {}), display_types=task.get("display_types"),
        photo_order=task.get("photo_order", []), photo_layout=task.get("photo_layout", {}),
        report_type=task.get("report_type", "水表")
    ))
    _start_required_sections(doc)
    if original_sections != [etree.tostring(s) for s in _xp(doc.element, ".//w:sectPr")]:
        raise ValueError("合成改变了模板分节设置，已停止输出")
    _validate_attachment_headers(doc)
    settings = doc.settings.element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        settings.append(update)
    update.set(qn("w:val"), "true")
    _editable_page_breaks(doc)
    output.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output)
    _normalize_report_fonts(output)
    _load(output)
    return {"warnings": warnings + ["Word 页码域已标记更新；需本机 Word 重分页核对。"]}
