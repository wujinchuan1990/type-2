"""Explicit atmospheric-pressure evidence and optional project footer output."""

from copy import deepcopy
from decimal import Decimal
import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph


LABEL = r"大气压(?:力)?"
NUMBER = r"\d+(?:\.\d+)?"


def extract_pressure(text, source):
    matches = list(re.finditer(
        LABEL + r"(?:\s*[：:]\s*|\s+)(.*?)(?=[\u4e00-\u9fff]+\s*[：:]|$)", text))
    return [{"value": match[1].strip(), "source": source} for match in matches]


def pressure_bounds(text):
    text = re.sub(r"[\s（）()]", "", text)
    match = re.fullmatch(r"(" + NUMBER + r")(?:[~～至—–-](" + NUMBER + r"))?(kPa|hPa|Pa)", text, re.I)
    if not match:
        raise ValueError("须为带Pa、hPa或kPa单位的明确数值或范围")
    scale = {"pa": Decimal(".001"), "hpa": Decimal(".1"), "kpa": Decimal(1)}[match[3].lower()]
    low, high = Decimal(match[1]) * scale, Decimal(match[2] or match[1]) * scale
    if low <= 0 or high < low:
        raise ValueError("大气压力数值或范围无效")
    return low, high


def pressure_input_value(text):
    text = text.strip()
    if not text:
        return ""
    if text == "/":
        return "/"
    if not re.search(r"[A-Za-z]", text):
        text += " kPa"
    try:
        low, high = pressure_bounds(text)
    except ValueError as exc:
        raise ValueError("大气压力：" + str(exc)) from exc
    value = format(low.normalize(), "f")
    if high != low:
        value += "～" + format(high.normalize(), "f")
    return value + " kPa"


def pressure_output_value(text):
    value = pressure_input_value(text)
    if not value:
        return "待补录"
    if value == "/":
        return "/"
    number = value.removesuffix(" kPa")
    return ("（" + number + "）" if "～" in number else number) + "kPa"


def summarize_pressure(project, parsed, report_type):
    import documents as d

    if report_type == "燃气表":
        from gas_records import project_kind
        name = project_kind(project)
        normalize = d._project_name
    elif report_type in ("流量计", "流量计（气体）", "流量计（液体）"):
        from flowmeter_records import _project_name
        name, normalize = _project_name(project["name"]), _project_name
    else:
        name, normalize = d._project_name(project["name"]), d._project_name
    records = [record for data in parsed for record in data.get("records", [])
               if normalize(record.get("project", "")) == name]
    bounds, sources, issues = [], [], []
    for record in records:
        candidates = record.get("atmospheric_pressure_candidates", [])
        if not candidates:
            issues.append(record["source"] + "：未识别到大气压力。")
            continue
        current = []
        for candidate in candidates:
            sources.append(candidate["source"] + "：" + candidate["value"])
            try:
                current.append(pressure_bounds(candidate["value"]))
            except ValueError as exc:
                issues.append(candidate["source"] + "：大气压力无法汇总（" + str(exc) + "）。")
        if len(set(current)) > 1:
            issues.append(record["source"] + "：同一记录存在多个大气压力值，未猜测采用。")
        elif current and len(current) == len(candidates):
            bounds.append(current[0])
    if not records:
        issues.append("未找到本项目原始记录，无法识别大气压力。")
    value = ""
    if not issues and bounds:
        low, high = min(pair[0] for pair in bounds), max(pair[1] for pair in bounds)
        value = format(low.normalize(), "f")
        if high != low:
            value += "～" + format(high.normalize(), "f")
        value += " kPa"
    return {"atmospheric_pressure": value, "pressure_sources": list(dict.fromkeys(sources)),
            "pressure_issues": list(dict.fromkeys(issues))}


def set_project_pressure(doc, enabled, value="", *, blank=False):
    import documents as d

    tables = [table for table in d._tables(doc)
              if all(label in d._text(table) for label in ("环境温度", "评价人员", "复核人员"))]
    if not tables:
        raise ValueError("项目缺少页底环境及人员信息表，无法设置大气压力栏。")
    for table in tables:
        rows = [row for row in d._rows(table) if "环境温度" in d._text(row)]
        if len(rows) != 1:
            raise ValueError("项目页底环境行不唯一，无法安全设置大气压力栏。")
        row = rows[0]
        cells = d._cells(row)
        if not enabled:
            active = False
            for cell in cells:
                text = d._text(cell)
                if re.fullmatch(LABEL + r"\s*[：:]?", text):
                    active = True
                    d._replace_result_field_text(cell, "")
                elif active:
                    if any(label in text for label in ("环境温度", "相对湿度", "评价人员", "复核人员")):
                        active = False
                    else:
                        d._replace_result_field_text(cell, "")
            for paragraph in d._xp(row, ".//w:p"):
                text = d._text(paragraph)
                if "大气压" in text:
                    text = re.sub(r"[ \t]*" + LABEL + r"\s*[：:].*?(?=评价人员\s*[：:]|复核人员\s*[：:]|$)", "", text)
                    d._replace_result_field_text(paragraph, text)
            continue
        text = "".join(d._text(cell) for cell in cells)
        fields = []
        for label, next_label, unit in (("环境温度", "相对湿度", "℃"), ("相对湿度", LABEL, "%")):
            match = re.search(label + r"\s*[：:](.*?)(?=" + next_label + r"\s*[：:]|$)", text)
            if not match:
                raise ValueError("项目页底缺少唯一环境标签：" + label)
            fields.append(label + "：" + ("        " + unit if blank else match[1].strip()))
        fields.append("大气压力：" + ("        kPa" if blank else pressure_output_value(value)))
        prototype = cells[0]
        paragraphs = d._xp(prototype, "./w:p")
        paragraph_style = deepcopy(paragraphs[0].pPr) if paragraphs else None
        runs = d._xp(prototype, ".//w:r")
        if not runs:
            raise ValueError("项目环境行缺少可保留的文字格式。")
        style = deepcopy(runs[0].rPr)
        width = sum(int(column.get(qn("w:w"))) for column in table.tblGrid)
        if width <= 0:
            raise ValueError("项目页底表宽无效。")
        for cell in cells[1:]:
            row.remove(cell)
        for child in list(prototype):
            if child.tag != qn("w:tcPr"):
                prototype.remove(child)
        pr = prototype.get_or_add_tcPr()
        pr.get_or_add_gridSpan().set(qn("w:val"), str(len(table.tblGrid)))
        pr.get_or_add_tcW().set(qn("w:w"), str(width))
        paragraph = OxmlElement("w:p")
        prototype.append(paragraph)
        if paragraph_style is not None:
            paragraph.append(paragraph_style)
        d._left_metadata_cell(prototype)
        ppr = paragraph.get_or_add_pPr()
        for tabs in d._xp(ppr, "./w:tabs"):
            ppr.remove(tabs)
        for percent in (31, 62):
            tab = OxmlElement("w:tab")
            tab.set(qn("w:val"), "left")
            tab.set(qn("w:pos"), str(round(width * percent / 100)))
            ppr.get_or_add_tabs().append(tab)
        run = Paragraph(paragraph, None).add_run("\t".join(fields))
        if style is not None:
            run._r.insert(0, style)
