"""Editable master-report options, located by labels rather than row numbers."""

import re
from copy import deepcopy

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph
from docx.text.run import Run

from contracts import DISPLAY_TYPES, MEASUREMENT_FIELDS

MEASUREMENT_CAPTION = "report-measurement-parameters"


def _text(element):
    return "".join(element.xpath(".//w:t/text()"))


def _compact(text):
    return re.sub(r"\s+", "", text)


def _measurement_table(doc):
    labels = ["序号", *(label for _, label in MEASUREMENT_FIELDS)]
    matches = []
    for table in doc.element.body.xpath("./w:tbl"):
        rows = table.xpath("./w:tr")
        headers = [_compact(_text(cell)) for cell in rows[0].xpath("./w:tc")] if rows else []
        extended = (len(headers) == 7 and headers[:6] == labels
                    and table.xpath("./w:tblPr/w:tblCaption/@w:val") == [MEASUREMENT_CAPTION])
        if headers == labels or extended:
            matches.append(table)
    if len(matches) != 1:
        raise ValueError("报告模板缺少或存在重复的计量器具测量参数表")
    table = matches[0]
    rows = table.xpath("./w:tr")[1:]
    count = len(table.xpath("./w:tr")[0].xpath("./w:tc"))
    if not rows or any(len(row.xpath("./w:tc")) != count for row in rows):
        raise ValueError("计量器具测量参数表列数不一致或缺少数据行")
    if table.xpath("./w:tr/w:tc/w:tcPr/w:vMerge | ./w:tr/w:tc/w:tcPr/w:gridSpan"):
        raise ValueError("计量器具测量参数表包含不支持的合并单元格")
    return table, rows


def _option_marks(doc, labels):
    matches = [p for p in doc.element.body.xpath(".//w:p")
               if all(label in _text(p) for label in labels)]
    if len(matches) != 1:
        raise ValueError("报告模板缺少或存在重复的选项段落：" + "、".join(labels))
    paragraph = matches[0]
    marks = []
    preceding = ""
    for run in paragraph.xpath("./w:r"):
        symbols = run.xpath("./w:sym")
        text = _text(run)
        checked = (len(symbols) == 1
                   and symbols[0].get(qn("w:font")) == "Wingdings 2"
                   and symbols[0].get(qn("w:char")) == "F053")
        if checked or text == "□":
            if len(marks) >= len(labels) or not _compact(preceding).endswith(labels[len(marks)]):
                raise ValueError("报告模板选项标志与标签对应不明确")
            if run.xpath("./*[not(self::w:rPr or self::w:sym or self::w:t)]"):
                raise ValueError("报告模板选项标志包含不支持的结构")
            marks.append((run, checked))
        preceding += text
    if len(marks) != len(labels):
        raise ValueError("报告模板选项缺少可识别的方框标志")
    return marks


def template_options(doc):
    table, rows = _measurement_table(doc)
    headers = table.xpath("./w:tr")[0].xpath("./w:tc")
    included = len(headers) == 7
    application = _option_marks(doc, ("新型", "改进型"))
    display = _option_marks(doc, DISPLAY_TYPES)
    if sum(checked for _, checked in application) != 1:
        raise ValueError("报告模板新型/改进型须恰有一项选中")
    return {
        "is_new_type": application[0][1],
        "measurement_parameters": [
            {**{key: "\n".join(Paragraph(p, None).text for p in cell.xpath("./w:p"))
                for (key, _), cell in zip(MEASUREMENT_FIELDS, row.xpath("./w:tc")[1:6])},
             **({"volume_display": "\n".join(Paragraph(p, None).text
                                            for p in row.xpath("./w:tc")[-1].xpath("./w:p"))}
                if included else {})}
            for row in rows
        ],
        "include_volume_display": included,
        "volume_display_label": _text(headers[-1]) if included else "体积量显示",
        "display_types": [label for label, (_, checked) in zip(DISPLAY_TYPES, display) if checked],
    }


def _fill_marks(marks, selected):
    if all(checked == wanted for (_, checked), wanted in zip(marks, selected)):
        return
    prototypes = {checked: run for run, checked in marks}
    if set(prototypes) != {True, False}:
        raise ValueError("报告模板须提供选中及未选中的原始方框格式")
    prototypes = {key: deepcopy(run) for key, run in prototypes.items()}
    for (run, checked), wanted in zip(marks, selected):
        if checked != wanted:
            run.getparent().replace(run, deepcopy(prototypes[wanted]))


def _superscript_cubic_metres(cell):
    for paragraph in cell.xpath("./w:p"):
        for run in list(paragraph.xpath("./w:r")):
            wrapper = Run(run, Paragraph(paragraph, None))
            value = wrapper.text
            if "m3" not in value:
                continue
            # _put produces one formatted run; split only the inserted unit token.
            for piece in re.split(r"(m3)", value):
                if not piece:
                    continue
                parts = (("m", False), ("3", True)) if piece == "m3" else ((piece, False),)
                for text, superscript in parts:
                    fragment = deepcopy(run)
                    part = Run(fragment, wrapper._parent)
                    part.text = text
                    part.font.superscript = superscript
                    run.addprevious(fragment)
            run.getparent().remove(run)


def fill_basic_options(doc, task):
    from documents import _put

    is_new = task.get("is_new_type")
    if is_new is not None:
        _fill_marks(_option_marks(doc, ("新型", "改进型")), (is_new, not is_new))
    display_types = task.get("display_types")
    if display_types is not None:
        _fill_marks(_option_marks(doc, DISPLAY_TYPES), [label in display_types for label in DISPLAY_TYPES])
    parameters = task.get("measurement_parameters")
    included = task.get("include_volume_display", False)
    existing_column = doc.element.body.xpath(
        f"./w:tbl[w:tblPr/w:tblCaption[@w:val='{MEASUREMENT_CAPTION}']]")
    if parameters is None and not included and not existing_column:
        return
    table, original = _measurement_table(doc)
    if not included and len(table.xpath("./w:tblGrid/w:gridCol")) == 7:
        total = sum(int(column.get(qn("w:w"))) for column in table.xpath("./w:tblGrid/w:gridCol"))
        table.tblGrid.remove(table.tblGrid[-1])
        for row in table.xpath("./w:tr"):
            row.remove(row.xpath("./w:tc")[-1])
        _resize_measurement_columns(table, total)
    if parameters is not None:
        if not parameters:
            raise ValueError("请保留至少一行计量器具测量参数")
        for row in original:
            table.remove(row)
        for index, values in enumerate(parameters):
            row = deepcopy(original[index] if index < len(original) else original[-1])
            cells = row.xpath("./w:tc")
            _put(cells[0], str(index + 1))
            for (key, _), cell in zip(MEASUREMENT_FIELDS, cells[1:6]):
                value = values.get(key, "")
                if value.strip() or index >= len(original):
                    _put(cell, value)
                    _superscript_cubic_metres(cell)
            table.append(row)
    if included:
        label = task.get("volume_display_label", "体积量显示")
        if not label.strip():
            raise ValueError("请填写体积量显示的列名")
        columns = table.xpath("./w:tblGrid/w:gridCol")
        total = sum(int(column.get(qn("w:w"))) for column in columns)
        if len(columns) == 6:
            table.tblGrid.append(deepcopy(columns[-1]))
            for row in table.xpath("./w:tr"):
                row.append(deepcopy(row.xpath("./w:tc")[-1]))
        caption = table.tblPr.find(qn("w:tblCaption"))
        if caption is None:
            caption = OxmlElement("w:tblCaption")
            table.tblPr.append(caption)
        caption.set(qn("w:val"), MEASUREMENT_CAPTION)
        for index, row in enumerate(table.xpath("./w:tr")):
            cell = row.xpath("./w:tc")[-1]
            value = label if index == 0 else (parameters[index - 1].get("volume_display", "")
                                             if parameters is not None else "")
            _put(cell, value)
            _superscript_cubic_metres(cell)
        _resize_measurement_columns(table, total)


def _resize_measurement_columns(table, total):
    columns = table.xpath("./w:tblGrid/w:gridCol")
    old_widths = [int(column.get(qn("w:w"))) for column in columns]
    if len(columns) not in (6, 7) or total <= 0 or any(width <= 0 for width in old_widths):
        raise ValueError("计量器具测量参数表缺少有效列宽，无法自动分配。")
    if len(columns) == 7:
        number_width = old_widths[0]
        if total <= number_width:
            raise ValueError("计量器具测量参数表宽度不足，无法保留序号列")
        widths = [number_width, *[round((total - number_width) * width / sum(old_widths[1:]))
                                 for width in old_widths[1:]]]
    else:
        widths = [round(total * width / sum(old_widths)) for width in old_widths]
    widths[-1] += total - sum(widths)
    for column, width in zip(columns, widths):
        column.set(qn("w:w"), str(width))
    for row in table.xpath("./w:tr"):
        cells = row.xpath("./w:tc")
        if len(cells) != len(widths):
            raise ValueError("计量器具测量参数表列数与网格不一致")
        for cell, width in zip(cells, widths):
            cell.get_or_add_tcPr().get_or_add_tcW().set(qn("w:w"), str(width))
            cell.tcPr.tcW.set(qn("w:type"), "dxa")
    table.tblPr.get_or_add_tblLayout().set(qn("w:type"), "fixed")
    width_element = table.tblPr.find(qn("w:tblW"))
    if width_element is None:
        width_element = OxmlElement("w:tblW")
        table.tblPr.insert_element_before(width_element, "w:jc", "w:tblInd", "w:tblBorders", "w:tblLayout")
    width_element.set(qn("w:w"), str(total))
    width_element.set(qn("w:type"), "dxa")
    return widths


def align_measurement_table(doc, reference):
    """Match Word's visible edges, including row-level overrides of table geometry."""
    table, _ = _measurement_table(doc)
    total = sum(int(value) for value in reference.xpath("./w:tblGrid/w:gridCol/@w:w"))
    widths = _resize_measurement_columns(table, total)
    source_row = reference.xpath("./w:tr")[0]
    alignment = (source_row.xpath("./w:trPr/w:jc/@w:val")
                 or reference.xpath("./w:tblPr/w:jc/@w:val") or ["left"])[0]
    settings = {
        "jc": ({"val": alignment}, ("w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:tblLayout")),
        "tblCellSpacing": ({"w": "0", "type": "dxa"}, ("w:tblInd", "w:tblBorders", "w:tblLayout")),
        "tblInd": ({"w": "0", "type": "dxa"}, ("w:tblBorders", "w:tblLayout")),
        "tblCellMar": ({}, ("w:tblLook",)),
    }
    for name, (defaults, successors) in settings.items():
        for old in table.tblPr.findall(qn("w:" + name)):
            table.tblPr.remove(old)
        sources = source_row.xpath("./w:tblPrEx/w:" + name) or reference.xpath("./w:tblPr/w:" + name)
        setting = deepcopy(sources[0]) if sources else OxmlElement("w:" + name)
        if not sources:
            for key, value in defaults.items():
                setting.set(qn("w:" + key), value)
        if name == "jc":
            setting.set(qn("w:val"), alignment)
        table.tblPr.insert_element_before(setting, *successors)
    for row in table.xpath("./w:tr"):
        for override in row.xpath("./w:trPr/w:jc | ./w:trPr/w:tblCellSpacing | "
                                  "./w:tblPrEx/w:jc | ./w:tblPrEx/w:tblInd | "
                                  "./w:tblPrEx/w:tblW | ./w:tblPrEx/w:tblCellMar | "
                                  "./w:tblPrEx/w:tblCellSpacing | ./w:tblPrEx/w:tblLayout"):
            override.getparent().remove(override)
        for cell, width in zip(row.xpath("./w:tc"), widths):
            cell.get_or_add_tcPr().get_or_add_tcW().set(qn("w:w"), str(width))
            cell.tcPr.tcW.set(qn("w:type"), "dxa")
        # In compatibility layouts the first cell's margin participates in the
        # left border position; equal grid widths/indents alone are insufficient.
        for index, sides in ((0, ("left", "start")), (-1, ("right", "end"))):
            cell = row.xpath("./w:tc")[index]
            source = source_row.xpath("./w:tc")[index]
            margins = cell.tcPr.find(qn("w:tcMar"))
            for side in sides:
                if margins is not None:
                    for old in margins.findall(qn("w:" + side)):
                        margins.remove(old)
                values = source.xpath("./w:tcPr/w:tcMar/w:" + side)
                if values:
                    if margins is None:
                        margins = OxmlElement("w:tcMar")
                        cell.tcPr.insert_element_before(margins, "w:textDirection", "w:tcFitText", "w:vAlign")
                    margins.append(deepcopy(values[0]))
