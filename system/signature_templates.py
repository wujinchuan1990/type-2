"""Apply the pressure-selected editable project footer without restyling its cells."""

from copy import deepcopy
import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docxcompose.composer import Composer

from atmospheric_pressure import pressure_output_value


SIGNATURE_FILENAMES = {
    False: "signature-non atmosphere pressure.docx",
    True: "signature- atmospheric pressrue.docx",
}
DESCRIPTION = "report-signature-template:"
LABELS = ("序号", "环境温度", "相对湿度", "大气压力", "评价人员", "复核人员")
REFERENCE_WIDTH_TWIPS = 8522


def is_signature_table(table):
    node = table.tblPr.find(qn("w:tblDescription"))
    return node is not None and node.get(qn("w:val"), "").startswith(DESCRIPTION)


def _tables(doc):
    import documents as d

    return [table for table in d._tables(doc)
            if all(label in d._text(table) for label in ("环境温度", "评价人员", "复核人员"))]


def _slots(table):
    import documents as d

    slots = {}
    pattern = re.compile(r"(" + "|".join(LABELS) + r")\s*[：:]")
    for row in d._rows(table):
        cells = d._cells(row)
        for index, cell in enumerate(cells):
            paragraphs = d._xp(cell, "./w:p")
            if d._norm(d._text(cell)) in LABELS:
                label = d._norm(d._text(cell))
                if index + 1 >= len(cells):
                    raise ValueError("页底标签缺少填写单元格：" + label)
                paragraphs = d._xp(cells[index + 1], "./w:p")
                if len(paragraphs) != 1:
                    raise ValueError("页底填写单元格须为单段：" + label)
                text = "".join(d._xp(paragraphs[0], ".//w:t/text()"))
                matches = [(label, 0, len(text), text, True)]
            else:
                matches = []
                for paragraph in paragraphs:
                    text = "".join(d._xp(paragraph, ".//w:t/text()"))
                    found = list(pattern.finditer(text))
                    for offset, match in enumerate(found):
                        end = found[offset + 1].start() if offset + 1 < len(found) else len(text)
                        matches.append((match[1], match.end(), end, text[match.end():end], False))
                    if found and len(paragraphs) != 1:
                        raise ValueError("页底标签所在单元格须为单段")
            for label, start, end, source, separate in matches:
                if label in slots:
                    raise ValueError("页底填写标签重复：" + label)
                slots[label] = (paragraphs[0], start, end, source, separate)
    return slots


def _replace_slot(paragraph, start, end, value):
    import documents as d

    if start == end:
        offset = 0
        for node in d._xp(paragraph, ".//w:t"):
            text = node.text or ""
            if offset == start or offset < start < offset + len(text):
                position = start - offset
                d._write_slot_text(node, text[:position] + value + text[position:])
                return
            offset += len(text)
    d._replace_text_span(paragraph, start, end, value)


def _set_columns(table, widths):
    import documents as d

    for column, width in zip(table.tblGrid, widths):
        column.set(qn("w:w"), str(width))
    for row in d._rows(table):
        start = 0
        for cell in d._cells(row):
            end = start + cell.grid_span
            cell.get_or_add_tcPr().get_or_add_tcW().set(qn("w:w"), str(sum(widths[start:end])))
            cell.tcPr.tcW.set(qn("w:type"), "dxa")
            start = end
    table.tblPr.get_or_add_tblLayout().set(qn("w:type"), "fixed")
    width = table.tblPr.find(qn("w:tblW"))
    if width is None:
        width = OxmlElement("w:tblW")
        table.tblPr.insert_element_before(width, "w:jc", "w:tblInd", "w:tblBorders", "w:tblLayout")
    width.set(qn("w:w"), str(sum(widths)))
    width.set(qn("w:type"), "dxa")


def _set_total_width(table, width_twips):
    if isinstance(width_twips, bool) or not isinstance(width_twips, int) or width_twips <= 0:
        raise ValueError("页底签名表宽须为正整数twips。")
    original = [int(column.get(qn("w:w"))) for column in table.tblGrid]
    if not original or any(width <= 0 for width in original):
        raise ValueError("页底签名表列宽无效，不能按比例调整。")
    total = sum(original)
    widths, consumed, boundary = [], 0, 0
    for width in original:
        consumed += width
        end = width_twips * consumed // total
        widths.append(end - boundary)
        boundary = end
    if any(width <= 0 for width in widths):
        raise ValueError("页底签名表宽不足以保留全部列。")
    _set_columns(table, widths)


def _fit_columns(doc, table, *, layout_settings=None):
    import documents as d

    original = [int(column.get(qn("w:w"))) for column in table.tblGrid]
    cells = []
    for row in d._rows(table):
        start = 0
        for cell in d._cells(row):
            end = start + cell.grid_span
            if end > len(original):
                raise ValueError("页底签名模板合并单元格与列网格不一致")
            cells.append((start, end, cell))
            start = end
        if start != len(original):
            raise ValueError("页底签名模板行宽与列网格不一致")
    settings = layout_settings if layout_settings is not None else doc.settings.element
    compression = d._xp(settings, "./w:characterSpacingControl/@w:val")
    attempts = (True, False) if compression in (
        ["compressPunctuation"], ["compressPunctuationAndJapaneseKana"]) else (False,)
    for compressed in attempts:
        widths = original[:]
        constraints = [(start, end, d._sample_header_min_width(
            doc, table, cell, compress_punctuation=compressed)) for start, end, cell in cells]
        for start, end, required in constraints:
            remaining = max(0, required - sum(widths[start:end]))
            for recipient in range(start, end):
                for donor in range(len(widths)):
                    if start <= donor < end:
                        continue
                    slack = min([sum(widths[left:right]) - minimum
                                 for left, right, minimum in constraints
                                 if left <= donor < right and not left <= recipient < right] +
                                [widths[donor] - 1])
                    taken = min(remaining, max(0, slack))
                    widths[donor] -= taken
                    widths[recipient] += taken
                    remaining -= taken
            if remaining:
                break
        else:
            if widths != original:
                _set_columns(table, widths)
            return
    raise ValueError("项目页底内容超出模板总表宽，无法保持字号、0.8cm行高及完整显示；"
                     "请调整签名模板或精简填写内容后重新生成。")


def signature_template_path(root, enabled):
    import documents as d

    path = d._root(root) / "report model/type evaluation results" / SIGNATURE_FILENAMES[enabled]
    if not path.is_file():
        raise ValueError("页底签名模板不存在：" + str(path))
    return path


def _load(root, enabled):
    import documents as d

    path = signature_template_path(root, enabled)
    source = d._isolated_styles(d._load(path), "reportSignature" + str(int(enabled)) + "_",
                                preserve_table_defaults=True)
    tables = _tables(source)
    if len(tables) != 1:
        raise ValueError("页底签名模板须含唯一环境及人员表：" + path.name)
    table = tables[0]
    rows = d._rows(table)
    width = sum(int(value) for value in d._xp(table, "./w:tblGrid/w:gridCol/@w:w"))
    if width <= 0 or d._xp(table, ".//w:sectPr | .//w:br[@w:type='page']"):
        raise ValueError("页底签名模板列宽无效或含内部强制分页：" + path.name)
    if len(rows) == 2:
        equipment = table.getprevious()
        if equipment is None or equipment.tag != qn("w:p") or "所用试验设备信息见" not in d._text(equipment):
            raise ValueError("页底签名模板缺少设备引用段落：" + path.name)
        row, cell = OxmlElement("w:tr"), OxmlElement("w:tc")
        properties = cell.get_or_add_tcPr()
        properties.get_or_add_tcW().set(qn("w:w"), str(width))
        properties.tcW.set(qn("w:type"), "dxa")
        properties.get_or_add_gridSpan().set(qn("w:val"), str(len(table.tblGrid)))
        # The standalone paragraph starts at the body margin, not the table's cell inset.
        indent = d._xp(table, "./w:tblPr/w:tblInd/@w:w")
        margins = OxmlElement("w:tcMar")
        left = OxmlElement("w:left")
        left.set(qn("w:w"), str(max(0, -int(indent[0])) if indent else 0))
        left.set(qn("w:type"), "dxa")
        margins.append(left)
        properties.append(margins)
        cell.append(equipment)
        row.append(cell)
        table.insert(table.index(rows[0]), row)
    elif len(rows) != 3 or "所用试验设备信息见" not in d._text(rows[0]):
        raise ValueError("页底签名模板须为设备、环境、人员三行：" + path.name)
    slots = _slots(table)
    expected = set(LABELS) if enabled else set(LABELS) - {"大气压力"}
    if set(slots) != expected:
        raise ValueError("页底签名模板填写标签与压力开关不一致：" + path.name)
    for index, row in enumerate(d._rows(table)):
        properties = row.get_or_add_trPr()
        height = properties.get_or_add_trHeight()
        if height.get(qn("w:val")) not in (None, str(d.PROJECT_METADATA_ROW_HEIGHT_TWIPS)):
            raise ValueError("页底签名模板行高须为0.8cm：" + path.name)
        height.set(qn("w:val"), str(d.PROJECT_METADATA_ROW_HEIGHT_TWIPS))
        height.set(qn("w:hRule"), "exact")
        no_split = properties.find(qn("w:cantSplit"))
        if no_split is None:
            no_split = OxmlElement("w:cantSplit")
            properties.insert_element_before(no_split, "w:trHeight", "w:tblHeader")
        no_split.set(qn("w:val"), "true")
        for paragraph in d._xp(row, ".//w:p"):
            d._paragraph_setting(paragraph, "pageBreakBefore", False)
            d._paragraph_setting(paragraph, "keepNext", index < 2)
    for name, value in (("tblCaption", "report-project-metadata"),
                        ("tblDescription", DESCRIPTION + str(int(enabled)))):
        node = table.tblPr.find(qn("w:" + name))
        if node is None:
            node = OxmlElement("w:" + name)
            table.tblPr.append(node)
        node.set(qn("w:val"), value)
    return d._source_fragment(source, [table])


def apply_signature_templates(doc, root, enabled, *, project=None, equipment_numbers=None, blank=False,
                              layout_settings=None, width_twips=None):
    import documents as d

    tables = _tables(doc)
    if not tables:
        raise ValueError("项目缺少页底信息组，无法替换签名模板")
    prototype = _load(root, enabled)
    if width_twips is not None:
        _set_total_width(_tables(prototype)[0], width_twips)
    for reference in tables:
        source = deepcopy(prototype)
        table = _tables(source)[0]
        if blank:
            from record_documents import _blank_report_metadata

            _blank_report_metadata(table)
            if (is_signature_table(reference) and
                    reference.tblPr.find(qn("w:tblDescription")).get(qn("w:val")) ==
                    table.tblPr.find(qn("w:tblDescription")).get(qn("w:val"))):
                widths = [int(column.get(qn("w:w"))) for column in reference.tblGrid]
                original = [int(column.get(qn("w:w"))) for column in table.tblGrid]
                if len(widths) != len(original) or sum(widths) != sum(original):
                    raise ValueError("报告页底信息组与本次签名模板总宽或列数不一致")
                if widths != original:
                    _set_columns(table, widths)
        else:
            old = _slots(reference)
            if not (set(LABELS) - {"大气压力"}).issubset(old):
                raise ValueError("项目页底信息组缺少设备、温湿度或人员填写标签")
            values = {label: slot[3].strip() for label, slot in old.items()}
            current = project or {}
            for label, key in (("环境温度", "temperature"), ("相对湿度", "humidity")):
                if d._has_text(current.get(key)):
                    values[label] = current[key]
            for label, key in (("评价人员", "evaluators"), ("复核人员", "reviewers")):
                names = "、".join(name for name in current.get(key, []) if d._has_text(name))
                if names:
                    values[label] = names
            ids = current.get("equipment_ids", [])
            if ids:
                numbers = equipment_numbers or {}
                if any(identity not in numbers for identity in ids):
                    raise ValueError("项目使用的设备未列入本次设备清单")
                values["序号"] = d._number_ranges([numbers[identity] for identity in ids])
            if enabled:
                values["大气压力"] = pressure_output_value(current.get("atmospheric_pressure", ""))
            temperature = d._result_environment_parts(_slots(table)["环境温度"][3], "环境温度")
            if temperature and temperature["separator"]:
                values["序号"] = re.sub(r"[~～]", lambda _: temperature["separator"], values.get("序号", ""))
                if enabled:
                    values["大气压力"] = values["大气压力"].replace("～", temperature["separator"])
            for label, (paragraph, start, end, original, separate) in reversed(list(_slots(table).items())):
                value = values.get(label, "")
                unit = re.search(r"(?:℃|°C|%RH|%|kPa|Pa)\s*$", original)
                keep_unit = unit is not None and (
                    label in ("环境温度", "相对湿度") or
                    (label == "大气压力" and value not in ("/", "待补录")))
                limit = unit.start() if keep_unit else len(original)
                left = len(original) - len(original.lstrip())
                right = max(left, len(original[:limit].rstrip()))
                end = start + right
                start += left
                original = original[left:right]
                if label in ("环境温度", "相对湿度") and d._has_text(value):
                    rendered = d._format_result_environment(original, str(value), label, separate or keep_unit)
                else:
                    rendered = str(value).removesuffix("kPa") if keep_unit and label == "大气压力" else str(value)
                _replace_slot(paragraph, start, end, rendered or " ")
            _fit_columns(source, table, layout_settings=(
                layout_settings if layout_settings is not None else doc.settings.element))
        floating = table.tblPr.find(qn("w:tblpPr"))
        if floating is None:
            old = reference.tblPr.find(qn("w:tblpPr"))
            if old is None:
                raise ValueError("项目页底信息组缺少页底定位")
            floating = deepcopy(old)
            # Nesting ignores the inline table indent; carry it into the outer anchor.
            indent = d._xp(table, "./w:tblPr/w:tblInd/@w:w")
            floating.set(qn("w:horzAnchor"), "margin")
            floating.set(qn("w:tblpX"), indent[0] if indent else "0")
            floating.attrib.pop(qn("w:tblpXSpec"), None)
            table.tblPr.insert(0, floating)
        floating.set(qn("w:vertAnchor"), "margin")
        floating.set(qn("w:tblpYSpec"), "bottom")
        floating.attrib.pop(qn("w:tblpY"), None)
        body = reference.getparent()
        if body is not doc.element.body:
            raise ValueError("项目页底信息组须在分页包装前替换")
        index = body.index(reference)
        body.remove(reference)
        d._insert(Composer(doc), index, source)
