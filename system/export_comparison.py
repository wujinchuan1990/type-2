"""Read-only comparison of saved exports, not a visual/pagination or approval check.

``unavailable`` means at least one required snapshot/document could not be
verified; successfully compared sections are still returned. Locations are XML
paragraph/table positions, never claimed to be Word physical page numbers.
"""

import hashlib
import io
import json
import posixpath
from difflib import SequenceMatcher
from pathlib import Path
from urllib.parse import unquote, urlsplit
from zipfile import BadZipFile, ZipFile
from zlib import error as CompressionError

from lxml import etree


W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS = {"w": W}
NOISE = {
    "revision", "version", "updated_at", "created_at", "confirmed",
    "samples_confirmed", "source", "sources", "candidates", "sample_candidates",
    "manual_override", "observation_manual_override", "auto_value", "auto_number",
    "number_origin", "number_source", "matched_sample_id", "match_issue",
    "record_summary", "report_environment", "evaluation_summary",
    "observation_summary", "legacy_sample_fields", "observations", "items",
    "pressure_sources", "pressure_issues",
    "pressure_auto_value",
}
LABELS = {
    "report_type": "报告对象类型",
    "include_atmospheric_pressure": "显示并输出大气压力",
    "atmospheric_pressure": "大气压力",
    "pressure_manual_override": "大气压力手工填写",
    "custom_supplementary_notes": "新增补充说明",
    "other_information_order": "说明条款顺序",
    "observation_certificates": "观察证书备注",
    "explosion": "防爆证书", "protection": "防护证书",
    "classification_code": "分类编码选择",
    "test_basis": "试验依据",
    "overview": "基本信息", "projects": "试验项目", "samples": "样机清单",
    "equipment": "设备清单", "photos": "照片关联", "photo_edits": "照片编辑", "photo_order": "照片顺序",
    "photo_layout": "照片与标题排版",
    "photo_components": "零部件照片", "record_ids": "原始记录及顺序",
    "outline_id": "大纲", "items_asset_id": "项目一览表素材",
    "observe_asset_id": "观察素材", "measurement_parameters": "测量参数",
    "include_volume_display": "输出体积量显示列", "volume_display_label": "体积量显示列名",
    "volume_display": "体积量显示",
    "sample_table_excluded_columns": "样机信息表隐藏列",
    "display_types": "显示型式", "is_new_type": "新型申请",
    "value": "填写值", "name": "名称", "number": "编号", "label": "标签",
    "rows": "实测结果", "start": "开始时间", "end": "结束时间",
    "temperature": "环境温度", "humidity": "相对湿度", "evaluators": "评价人员",
    "reviewers": "复核人员", "equipment_ids": "所用设备", "formatting": "字符格式",
    "sealed": "封样", "disassembly": "拆样", "rotation": "旋转角度",
    "crop": "裁剪", "width_cm": "宽度（厘米）", "report_number": "报告编号",
    "manufacturer": "制造单位", "agent": "代理人", "instrument_name": "器具名称",
    "sample_delivery_date": "送样日期", "supplementary_notes": "补充说明",
}
SCOPE_WARNING = (
    "比较基于任务快照与可编辑 Word XML（含页眉页脚、图片、尺寸、裁剪及格式）；"
    "位置为文档结构位置，不是物理页码，不替代完整视觉分页核对或人工审版。"
)
MISSING = object()


def _json(value):
    return json.dumps(value, ensure_ascii=False, sort_keys=True, separators=(",", ":"))


def _display(value):
    if value is MISSING:
        return "（不存在）"
    return value if isinstance(value, str) else _json(value)


def _clean_task(value):
    if isinstance(value, dict):
        return {
            key: _clean_task(item) for key, item in value.items()
            if key not in NOISE and not key.endswith(("_issues", "_source", "_confirmed"))
            and key not in {"issues", "notes", "old"}
            and not (key == "sample_table_excluded_columns" and item == [])
            and not (key == "classification_code" and item == "")
        }
    if isinstance(value, list):
        return [_clean_task(item) for item in value]
    return value


def _task_changes(before, after, location="任务"):
    changes = []
    if isinstance(before, dict) and isinstance(after, dict):
        for key in sorted(before.keys() | after.keys()):
            changes.extend(_task_changes(
                before.get(key, MISSING), after.get(key, MISSING),
                location + " / " + LABELS.get(key, key),
            ))
    elif isinstance(before, list) and isinstance(after, list):
        for index in range(max(len(before), len(after))):
            old = before[index] if index < len(before) else MISSING
            new = after[index] if index < len(after) else MISSING
            item = new if isinstance(new, dict) else old
            label = ""
            if isinstance(item, dict):
                label = item.get("label") or item.get("name") or item.get("key") or ""
            changes.extend(_task_changes(old, new, f"{location} / 第{index + 1}项"
                                         + (f"（{label}）" if label else "")))
    elif type(before) is not type(after) or before != after:
        changes.append({"location": location, "before": _display(before), "after": _display(after)})
    return changes


def _local(tag):
    return etree.QName(tag).localname


XML_LABELS = {
    "p": "段落", "pPr": "段落格式", "r": "文字段", "rPr": "字符格式",
    "b": "加粗", "i": "斜体", "u": "下划线", "sz": "字号（半磅）",
    "szCs": "复杂文字字号（半磅）", "rFonts": "字体", "eastAsia": "中文字体",
    "ascii": "ASCII西文字体", "hAnsi": "高ANSI西文字体", "color": "颜色", "val": "值",
    "jc": "对齐", "spacing": "间距", "before": "段前", "after": "段后",
    "line": "行距", "ind": "缩进", "left": "左侧", "right": "右侧",
    "firstLine": "首行", "hanging": "悬挂", "keepNext": "与下段同页",
    "keepLines": "段中不分页", "pageBreakBefore": "段前分页",
    "t": "文字范围", "instrText": "域代码", "delText": "删除文字",
    "tblPr": "表格格式", "tblGrid": "列网格", "gridCol": "列",
    "tcPr": "单元格格式", "trPr": "行格式", "tblW": "表格宽度",
    "tcW": "单元格宽度", "trHeight": "行高", "vMerge": "纵向合并",
    "gridSpan": "跨列数", "tblBorders": "表格边框", "tcBorders": "单元格边框",
    "tblCellMar": "单元格边距", "vAlign": "垂直对齐", "w": "宽度",
    "h": "高度", "type": "类型", "sectPr": "分节设置", "pgSz": "纸张尺寸",
    "pgMar": "页边距", "orient": "方向", "headerReference": "页眉关联",
    "footerReference": "页脚关联", "id": "标识", "pStyle": "段落样式",
    "rStyle": "字符样式", "tblStyle": "表格样式", "style": "样式",
    "basedOn": "继承样式", "docDefaults": "默认格式", "vertAlign": "上下标",
    "drawing": "插图", "inline": "嵌入布局", "anchor": "浮动布局",
    "extent": "显示尺寸", "ext": "图形尺寸", "cx": "宽度（EMU）", "cy": "高度（EMU）",
    "graphic": "图形", "graphicData": "图形数据", "pic": "图片",
    "blipFill": "图片填充", "blip": "图片内容", "embed": "嵌入",
    "link": "链接", "srcRect": "裁剪（十万分比）", "l": "左", "spPr": "形状格式",
    "xfrm": "变换", "rot": "旋转（六万分之一度）", "off": "偏移",
    "x": "横坐标", "y": "纵坐标", "docPr": "图片说明", "descr": "替代说明",
}
BOOL_PROPERTIES = {"b", "i", "keepNext", "keepLines", "pageBreakBefore", "vanish", "strike"}


def _format_fields(tree):
    fields = {}

    def visit(node, prefix):
        if node is None:
            return
        tag, attrs, text, children = node
        name = _local(tag)
        if name in BOOL_PROPERTIES:
            fields[prefix] = "否" if attrs.get("{" + W + "}val") in {"0", "false", "off"} else "是"
            return
        for key, value in attrs.items():
            local = _local(key)
            label = ({"l": "左", "r": "右", "t": "上", "b": "下"}.get(local, local)
                     if name == "srcRect" else XML_LABELS.get(local, local))
            location = prefix + " / " + label
            if isinstance(value, dict) and "样式定义" in value:
                visit(value["样式定义"], location + " / 样式定义")
            elif isinstance(value, dict):
                for item, detail in value.items():
                    fields[location + " / " + item] = str(detail)
            else:
                display = str(value)
                try:
                    number = float(value)
                except (TypeError, ValueError):
                    pass
                else:
                    if local in {"cx", "cy"}:
                        display = f"{number / 360000:g} 厘米（{value} EMU）"
                    elif name == "srcRect" and local in {"l", "r", "t", "b"}:
                        display = f"{number / 1000:g}%（原值{value}）"
                    elif name in {"sz", "szCs"} and local == "val":
                        display = f"{number / 2:g} 磅（原值{value}）"
                    elif local == "rot":
                        display = f"{number / 60000:g}°（原值{value}）"
                fields[location] = display
        if text is not None:
            if isinstance(text, dict):
                for key, value in text.items():
                    fields[prefix + " / " + key] = str(value)
            else:
                fields[prefix + " / 文字"] = str(text)
        counts = {}
        for child in children:
            local = _local(child[0])
            counts[local] = counts.get(local, 0) + 1
            visit(child, prefix + " / " + XML_LABELS.get(local, local) + str(counts[local]))
        if not attrs and text is None and not children:
            fields[prefix] = "存在"

    if tree is not None:
        visit(tree, XML_LABELS.get(_local(tree[0]), _local(tree[0])))
    return fields


def _word_changes(before, after):
    """Align sibling XML blocks before comparing their content and properties."""
    changes = []
    indexes = {}
    for snapshot in (before, after):
        nodes = set(snapshot["groups"])
        nodes.update(child for group in snapshot["groups"].values() for child in group)
        index = {}
        for key, value in snapshot["values"].items():
            owner = key.rsplit(" / ", 1)[0]
            while owner not in nodes:
                parent = owner.rsplit(" / ", 1)[0]
                if parent == owner:
                    raise ValueError(f"无法定位比较内容的文档结构：{key}")
                owner = parent
            index.setdefault(owner, {})[key[len(owner):]] = value
        indexes[id(snapshot)] = index

    def properties(snapshot, path):
        return indexes[id(snapshot)].get(path, {})

    caches = ({}, {})

    def signature(snapshot, path, side):
        if path not in caches[side]:
            children = snapshot["groups"].get(path, [])
            value = [properties(snapshot, path),
                     [signature(snapshot, child, side) for child in children]]
            caches[side][path] = hashlib.sha256(_json(value).encode()).hexdigest()
        return caches[side][path]

    def emit(location, old, new):
        if old != new:
            changes.append({"location": location, "before": _display(old), "after": _display(new)})

    def fields(old, new, location):
        if isinstance(old, list) or isinstance(new, list):
            a = _format_fields(old) if isinstance(old, list) else {}
            b = _format_fields(new) if isinstance(new, list) else {}
            for key in sorted(a.keys() | b.keys()):
                emit(location + " / " + key, a.get(key, MISSING), b.get(key, MISSING))
        else:
            emit(location, old, new)

    def summary(snapshot, path):
        entries = []
        for key, value in snapshot["values"].items():
            if not key.startswith(path + " / ") or key.endswith(" / 结构"):
                continue
            relative = key[len(path) + 3:]
            if isinstance(value, list):
                entries.extend(relative + " / " + field + "：" + text
                               for field, text in _format_fields(value).items())
            else:
                entries.append(relative + "：" + str(value or "（空段落）"))
        return "\n".join(entries) or "（空结构）"

    def added(path):
        emit(path + " / 新增", MISSING, summary(after, path))

    def removed(path):
        emit(path + " / 删除", summary(before, path), MISSING)

    def compare_node(old, new):
        a, b = properties(before, old), properties(after, new)
        for key in sorted(a.keys() | b.keys()):
            fields(a.get(key, MISSING), b.get(key, MISSING), new + key)
        align(before["groups"].get(old, []), after["groups"].get(new, []))

    def kind(path):
        return path.rsplit(" / ", 1)[-1].rstrip("0123456789")

    def align(old, new):
        matcher = SequenceMatcher(
            None, [signature(before, path, 0) for path in old],
            [signature(after, path, 1) for path in new], autojunk=False,
        )
        for operation, a, b, c, d in matcher.get_opcodes():
            if operation == "equal":
                continue
            if operation == "delete":
                for path in old[a:b]:
                    removed(path)
            elif operation == "insert":
                for path in new[c:d]:
                    added(path)
            else:
                # Exact unchanged siblings are already anchored. Pair only like
                # structural kinds within the remaining local replacement span.
                old_span, new_span = old[a:b], new[c:d]
                common = min(len(old_span), len(new_span))
                for index in range(common):
                    if kind(old_span[index]) == kind(new_span[index]):
                        compare_node(old_span[index], new_span[index])
                    else:
                        removed(old_span[index])
                        added(new_span[index])
                for path in old_span[common:]:
                    removed(path)
                for path in new_span[common:]:
                    added(path)

    def roots(snapshot):
        children = {child for group in snapshot["groups"].values() for child in group}
        return set(snapshot["groups"]) - children

    for root in sorted(roots(before) | roots(after)):
        compare_node(root, root)
    return changes


class _WordSnapshot:
    def __init__(self, content):
        self.archive = ZipFile(io.BytesIO(content))
        self.parts = {}
        self.relationships = {}
        self.fingerprints = {}
        self.styles = None
        self.style_cache = {}

    def style(self, identity, stack):
        if self.styles is None:
            self.styles = {}
            for kind, target, external in self.rels("word/document.xml").values():
                if kind.rsplit("/", 1)[-1] == "styles" and not external:
                    self.styles = {node.get("{" + W + "}styleId"): (node, target)
                                   for node in self.xml(target).findall("w:style", NS)}
        if identity not in self.styles:
            raise ValueError(f"无法解析样式引用：{identity}")
        marker = "style:" + identity
        if marker in stack:
            raise ValueError("样式继承存在循环，无法可靠比较")
        if identity not in self.style_cache:
            node, target = self.styles[identity]
            self.style_cache[identity] = self.canonical(node, target, stack + (marker,))
        return self.style_cache[identity]

    def xml(self, part):
        if part not in self.parts:
            self.parts[part] = etree.fromstring(
                self.archive.read(part),
                etree.XMLParser(resolve_entities=False, no_network=True, remove_comments=True),
            )
        return self.parts[part]

    def rels(self, part):
        if part not in self.relationships:
            folder, name = posixpath.split(part)
            relpath = posixpath.join(folder, "_rels", name + ".rels")
            result = {}
            if relpath in self.archive.namelist():
                for rel in self.xml(relpath):
                    target = rel.get("Target", "")
                    external = rel.get("TargetMode") == "External"
                    if not external:
                        target = posixpath.normpath(
                            target.lstrip("/") if target.startswith("/")
                            else posixpath.join(folder, target)
                        )
                    result[rel.get("Id")] = (rel.get("Type", ""), target, external)
            self.relationships[part] = result
        return self.relationships[part]

    def relationship(self, part, identity, stack):
        if identity not in self.rels(part):
            raise ValueError(f"{part} 存在无法解析的关系 {identity}")
        kind, target, external = self.rels(part)[identity]
        if external:
            if kind.rsplit("/", 1)[-1] != "hyperlink":
                raise ValueError(f"外部关联内容无法只读核验：{target}")
            return {"关系": kind.rsplit("/", 1)[-1], "外部目标": target}
        if target not in self.fingerprints:
            if target in stack:
                raise ValueError(f"无法完整比较循环文档关系：{target}")
            content = self.archive.read(target)
            if target.endswith(".xml"):
                content = _json(self.canonical(self.xml(target), target, stack + (target,))).encode()
            self.fingerprints[target] = hashlib.sha256(content).hexdigest()
        return {"关系": kind.rsplit("/", 1)[-1], "内容SHA256": self.fingerprints[target]}

    def canonical(self, element, part, stack=(), mode="full"):
        name = _local(element.tag)
        if name in {"proofErr", "lastRenderedPageBreak"}:
            return None
        parent = element.getparent()
        parent_tag = parent.tag if parent is not None else None
        if (parent_tag == "{" + W + "}abstractNum"
                and element.tag in {"{" + W + "}nsid", "{" + W + "}tmpl"}):
            return None
        if (parent_tag == "{" + W + "}numbering"
                and element.tag == "{" + W + "}numIdMacAtCleanup"):
            return None
        if (element.getparent() is not None and _local(element.getparent().tag) == "style"
                and name in {"name", "aliases", "next", "link", "uiPriority", "rsid"}):
            return None
        if mode == "format" and name in {"drawing", "pict", "object"}:
            return None
        attrs = {}
        for key, value in element.attrib.items():
            local = _local(key)
            # Template allocation codes are not list appearance or counter
            # semantics; retain numId/abstractNumId, levels and all overrides.
            if element.tag == "{" + W + "}lvl" and key == "{" + W + "}tplc":
                continue
            if (local.startswith("rsid") or local in {"paraId", "textId", "anchorId", "editId"}
                    or (name == "style" and local == "styleId")
                    or (name in {"docPr", "cNvPr"} and local in {"id", "name"})):
                continue
            if name in {"pStyle", "rStyle", "tblStyle", "basedOn"} and local == "val":
                value = {"样式定义": self.style(value, stack)}
            if key.startswith("{" + R + "}") or (
                    local == "relid" and "office:office" in key):
                value = self.relationship(part, value, stack)
            attrs[key] = value
        text = element.text
        if text is not None and name not in {"t", "instrText", "delText"} and not text.strip():
            text = None
        if mode == "format" and name in {"t", "instrText", "delText"}:
            # Keep span boundaries: moving bold from "A" to "AB" is a format
            # change even when the paragraph still reads "ABC".
            text = {"字符数": len(text or "")}
        children = [self.canonical(child, part, stack, mode) for child in element
                    if isinstance(child.tag, str)]
        return [element.tag, attrs, text, [child for child in children if child is not None]]

    def snapshot(self):
        result = {}
        groups = {}
        main = "word/document.xml"
        root = self.xml(main)
        body = root.find("w:body", NS)
        if body is None:
            raise ValueError("Word 缺少正文结构")

        def walk(element, part, location):
            counters = {}
            groups[location] = []
            for child in element:
                if not isinstance(child.tag, str):
                    continue
                name = _local(child.tag)
                counters[name] = counters.get(name, 0) + 1
                label = {"p": "段落", "tbl": "表格", "tr": "行", "tc": "单元格"}.get(name, name)
                path = f"{location} / {label}{counters[name]}"
                groups[location].append(path)
                if name == "p":
                    text = []
                    for node in child.iter():
                        tag = _local(node.tag)
                        if tag in {"t", "instrText", "delText"}:
                            text.append(node.text or "")
                        elif tag in {"tab", "br", "cr"}:
                            text.append("\t" if tag == "tab" else "\n")
                    result[path + " / 文字"] = "".join(text)
                    result[path + " / 格式"] = self.canonical(child, part, mode="format")
                    for index, drawing in enumerate(child.xpath(
                            ".//w:drawing | .//w:pict | .//w:object", namespaces=NS), 1):
                        result[f"{path} / 插图{index}（内容/尺寸/裁剪/格式）"] = self.canonical(
                            drawing, part)
                elif name in {"tbl", "tr", "tc", "sdt", "sdtContent", "customXml"}:
                    result[path + " / 结构"] = name
                    walk(child, part, path)
                else:
                    result[path + " / 格式或结构"] = self.canonical(child, part)

        walk(body, main, "正文")
        for index, section in enumerate(root.findall(".//w:sectPr", NS), 1):
            for reference in section:
                name = _local(reference.tag)
                if name not in {"headerReference", "footerReference"}:
                    continue
                identity = reference.get("{" + R + "}id")
                if identity not in self.rels(main):
                    raise ValueError("页眉页脚关系缺失")
                _, target, external = self.rels(main)[identity]
                if external:
                    raise ValueError("不支持外部页眉页脚")
                label = "页眉" if name == "headerReference" else "页脚"
                walk(self.xml(target), target,
                     f"第{index}节{label}（{reference.get('{' + W + '}type', 'default')}）")
        # Global styles, numbering, theme and settings can alter visible formatting
        # without changing any paragraph's direct properties.
        for kind, target, external in self.rels(main).values():
            short = kind.rsplit("/", 1)[-1]
            if short not in {"styles", "stylesWithEffects", "numbering", "theme", "settings",
                             "fontTable", "footnotes", "endnotes", "comments"}:
                continue
            if external:
                raise ValueError(f"不支持外部格式或文字部件：{short}")
            part_root = self.xml(target)
            location = {
                "styles": "全局格式 / 样式", "stylesWithEffects": "全局格式 / 效果样式",
                "numbering": "全局格式 / 编号", "theme": "全局格式 / 主题",
                "settings": "全局格式 / 文档设置", "fontTable": "全局格式 / 字体表",
                "footnotes": "脚注", "endnotes": "尾注", "comments": "批注",
            }[short]
            groups[location] = []
            counts = {}
            nodes = [(child, self.canonical(child, target)) for child in part_root]
            nodes = [(child, value) for child, value in nodes if value is not None]
            if short in {"styles", "stylesWithEffects"}:
                # IDs and names of imported styles are allocation details. Sort
                # and deduplicate semantic definitions, not randomly named IDs.
                unique = {_json(value): (child, value) for child, value in nodes}
                nodes = [unique[key] for key in sorted(unique)]
            for child, value in nodes:
                name = _local(child.tag)
                counts[name] = counts.get(name, 0) + 1
                path = f"{location} / {name}{counts[name]}"
                groups[location].append(path)
                result[path + " / 格式或附注"] = value
        return {"values": result, "groups": groups}


def _read_word(store, manifest, url_key, hash_key):
    identity = manifest.get("id")
    url = manifest.get(url_key)
    if not isinstance(identity, str) or not identity or Path(identity).name != identity:
        raise ValueError("导出标识缺失或无效")
    if not isinstance(url, str) or not url:
        raise ValueError("文件下载地址缺失")
    name = unquote(urlsplit(url).path.rsplit("/", 1)[-1])
    if not name or Path(name).name != name or name in {".", ".."}:
        raise ValueError("文件名无效")
    exports = (Path(store.directory) / "exports").resolve()
    folder = (exports / identity).resolve()
    path = (folder / name).resolve()
    if folder.parent != exports or path.parent != folder:
        raise ValueError("文件不在指定导出目录内")
    expected = manifest.get(hash_key)
    if not isinstance(expected, str) or not expected:
        raise ValueError("缺少文件 SHA256，无法核验历史文件")
    content = path.read_bytes()
    if hashlib.sha256(content).hexdigest().lower() != expected.lower():
        raise ValueError("SHA256 不符，文件可能已被修改，未将其作为可信历史比较")
    word = _WordSnapshot(content)
    try:
        return word.snapshot()
    finally:
        word.archive.close()


def compare_exports(store, current, previous):
    """Return JSON-ready differences; never read live tasks or modify any files.

    The caller selects adjacent exports of one task. Missing history/integrity
    failures yield ``unavailable`` with explicit warnings, not “no changes”.
    Empty ``changes`` is meaningful only with ``status == 'compared'``.
    """
    result = {
        "status": "first" if previous is None else "compared",
        "previous_id": previous.get("id") if previous is not None else None,
        "previous_revision": previous.get("revision") if previous is not None else None,
        "sections": [], "warnings": [],
    }
    if previous is None:
        return result
    result["warnings"].append(SCOPE_WARNING)
    if (not current.get("task_id") or not previous.get("task_id")
            or current["task_id"] != previous["task_id"]):
        result["status"] = "unavailable"
        result["warnings"].append("两次导出缺少相同任务标识，不能作为同一任务的相邻生成比较。")
        return result
    snapshots = [previous.get("task"), current.get("task")]
    if not all(isinstance(task, dict) and task for task in snapshots):
        result["status"] = "unavailable"
        result["warnings"].append("历史或本次导出缺少完整任务快照，任务字段无法比较；不会使用当前任务代替。")
    else:
        snapshots = [{"report_type": "水表", **task} for task in snapshots]
        result["sections"].append({
            "title": "任务字段",
            "changes": _task_changes(_clean_task(snapshots[0]), _clean_task(snapshots[1])),
        })
    for title, url_key, hash_key in (
            ("型式评价报告", "url", "sha256"),
            ("配套原始记录", "record_url", "record_sha256")):
        documents = []
        for label, manifest in (("上次", previous), ("本次", current)):
            try:
                documents.append(_read_word(store, manifest, url_key, hash_key))
            except (OSError, ValueError, KeyError, BadZipFile, etree.XMLSyntaxError,
                    RuntimeError, NotImplementedError, RecursionError,
                    EOFError, CompressionError) as exc:
                documents.append(None)
                result["status"] = "unavailable"
                result["warnings"].append(f"{label}{title}无法比较：{type(exc).__name__}: {exc}")
        if any(document is None for document in documents):
            continue
        before, after = documents
        changes = _word_changes(before, after)
        result["sections"].append({"title": title, "changes": changes})
    return result
