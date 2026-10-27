"""Experimental fragment widths; never shrink fonts or flatten native fields."""

from copy import deepcopy
from math import ceil
import re

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from docx.table import Table
from docx.text.run import Run

import documents as d
import result_pagination as rp
from observation_documents import _basic_table, _fit_basic_labels
from separation_layout import template_properties
from word_fonts import text_width


RESULT_WIDTHS = {"序号": 737, "项目名称": 1417, "技术要求": 3515, "样机编号": 2098, "试验结果": 1417}
RESULT_WIDTH_DESCRIPTION = (
    "结果表统一列宽：序号1.3、项目名称2.5、技术要求6.2、样机编号3.7、试验结果2.5cm，总宽16.2cm；"
    "项目名称内部子列等分，技术要求内部子列按源宽比例分配。")

HEADER_DESCRIPTION = (
    "页眉统一两行：首行空白，第二行保留报告编号所在文字；两行固定21磅，第二行段后半行。"
    "附件1标题页面居中、编号右对齐；长编号放不下时仅左移标题，保留原字号三个空格宽度。")
RESULT_TITLE_DESCRIPTION = "results的“型式评价结果-xxx”标题统一1.5倍行距、段前段后0行，保留字号、加粗及居中。"
RESULT_INTRO_DESCRIPTION = "results的试验开始时间、结束时间及数据记录引导段统一1.5倍行距、段前段后0行，保留原文字及其他格式。"
RESULT_CLOSING_DESCRIPTION = "results的合格判定要求、项目结论及异常情况记录段统一1.5倍行距、段前段后0行，保留原文字及其他格式。"


def _normalize_result_paragraph(doc, paragraph):
    fmt = d.Paragraph(paragraph, doc).paragraph_format
    fmt.line_spacing = 1.5
    fmt.space_before = Pt(0)
    fmt.space_after = Pt(0)
    spacing = paragraph.get_or_add_pPr().find(qn("w:spacing"))
    for name in ("beforeLines", "afterLines", "beforeAutospacing", "afterAutospacing"):
        spacing.set(qn("w:" + name), "0")


def _verify_zero_paragraph_spacing(doc, paragraph):
    roots = [paragraph, *_styles(d.Paragraph(paragraph, doc).style)]
    for side in ("before", "after"):
        for suffix in ("", "Lines", "Autospacing"):
            name = side + suffix
            values = [value for root in roots
                      for value in d._xp(root, "./w:pPr/w:spacing/@w:" + name)]
            values += d._xp(doc.styles.element,
                            "./w:docDefaults/w:pPrDefault/w:pPr/w:spacing/@w:" + name)
            value = values[0] if values else "0"
            if value not in ("0", "false", "off"):
                raise ValueError("试验拆分结果段落的段前段后未保持0行，停止交付。")


def _result_titles(doc):
    return [p for p in d._xp(doc.element.body, "./w:p")
            if re.match(r"^型式评价结果\s*[-—－–]", d._text(p).strip())]


def normalize_result_titles(doc):
    titles = _result_titles(doc)
    for p in titles:
        _normalize_result_paragraph(doc, p)
    return {"count": len(titles), "description": RESULT_TITLE_DESCRIPTION}


def verify_result_titles(doc):
    titles = _result_titles(doc)
    for p in titles:
        spacing = p.find("./" + qn("w:pPr") + "/" + qn("w:spacing"))
        if (spacing is None or spacing.get(qn("w:line")) != "360"
                or spacing.get(qn("w:lineRule")) != "auto"):
            raise ValueError("试验拆分结果标题未保持1.5倍行距，停止交付。")
        _verify_zero_paragraph_spacing(doc, p)
    return {"count": len(titles), "description": RESULT_TITLE_DESCRIPTION}


def _result_intro(doc):
    titles = set(_result_titles(doc))
    active = False
    for block in doc.element.body:
        if block in titles:
            active = True
        elif block.tag != qn("w:p"):
            active = False
        elif active and re.match(r"^试验(?:的)?(?:开始时间|结束时间|数据记录)",
                                 re.sub(r"\s+", "", d._text(block))):
            yield block


def normalize_result_intro(doc):
    paragraphs = list(_result_intro(doc))
    for paragraph in paragraphs:
        _normalize_result_paragraph(doc, paragraph)
    return {"count": len(paragraphs), "description": RESULT_INTRO_DESCRIPTION}


def verify_result_intro(doc):
    paragraphs = list(_result_intro(doc))
    for paragraph in paragraphs:
        spacing = paragraph.find("./" + qn("w:pPr") + "/" + qn("w:spacing"))
        if (spacing is None or spacing.get(qn("w:line")) != "360"
                or spacing.get(qn("w:lineRule")) != "auto"):
            raise ValueError("试验拆分结果开始时间、结束时间或数据记录段未保持1.5倍行距，停止交付。")
        _verify_zero_paragraph_spacing(doc, paragraph)
    return {"count": len(paragraphs), "description": RESULT_INTRO_DESCRIPTION}


def _result_closing(doc):
    for paragraph in d._xp(doc.element.body, "./w:p"):
        text = re.sub(r"\s+", "", d._text(paragraph))
        if ((text.startswith("本试验项目") and ("判定要求" in text or "结论" in text))
                or (text.startswith("试验") and "异常情况记录" in text)):
            yield paragraph


def normalize_result_closing(doc):
    paragraphs = list(_result_closing(doc))
    for paragraph in paragraphs:
        _normalize_result_paragraph(doc, paragraph)
    return {"count": len(paragraphs), "description": RESULT_CLOSING_DESCRIPTION}


def verify_result_closing(doc):
    paragraphs = list(_result_closing(doc))
    for paragraph in paragraphs:
        spacing = paragraph.find("./" + qn("w:pPr") + "/" + qn("w:spacing"))
        if (spacing is None or spacing.get(qn("w:line")) != "360"
                or spacing.get(qn("w:lineRule")) != "auto"):
            raise ValueError("试验拆分结果合格判定、结论或异常记录段未保持1.5倍行距，停止交付。")
        _verify_zero_paragraph_spacing(doc, paragraph)
    return {"count": len(paragraphs), "description": RESULT_CLOSING_DESCRIPTION}


def _report_headers(doc):
    seen = set()
    for section in doc.sections:
        for header in (section.header, section.first_page_header, section.even_page_header):
            if header.is_linked_to_previous or str(header.part.partname) in seen:
                continue
            seen.add(str(header.part.partname))
            element = header._element
            if d._xp(element, ".//w:tbl | .//w:drawing | .//w:pict | .//w:object | .//w:sdt"):
                raise ValueError("试验拆分页眉包含非段落内容，不能安全统一为两行。")
            paragraphs = d._xp(element, "./w:p")
            visible = [p for p in paragraphs if "".join(_node_text(n) for n in _visible_nodes(p)).strip()]
            if not visible:
                if d._xp(element, ".//w:fldChar | .//w:fldSimple | .//w:instrText"):
                    raise ValueError("试验拆分页眉只有域而没有可识别的报告编号，不能改写。")
                continue
            if len(visible) != 1 or not re.search(r"(?:报告|证书)编号\s*[：:]", d._text(visible[0])):
                raise ValueError("试验拆分页眉必须只有一个包含报告编号的文字段落，不能删除其他内容。")
            if d._xp(visible[0], ".//w:br | .//w:cr"):
                raise ValueError("试验拆分页眉编号段落含手动换行，不能安全统一为两行。")
            for p in paragraphs:
                if p is not visible[0] and d._xp(
                        p, ".//w:bookmarkStart | .//w:bookmarkEnd | .//w:fldChar | "
                           ".//w:fldSimple | .//w:instrText | .//w:commentRangeStart | "
                           ".//w:commentRangeEnd | .//w:commentReference | .//w:sectPr"):
                    raise ValueError("试验拆分页眉空段包含书签或域等受保护内容，不能删除。")
            yield element, paragraphs, visible[0]


def normalize_report_headers(doc):
    headers = list(_report_headers(doc))
    for element, paragraphs, content in headers:
        for p in paragraphs:
            if p is not content:
                element.remove(p)
        blank = OxmlElement("w:p")
        element.insert(element.index(content), blank)
        for p, after in ((blank, "0"), (content, "120")):
            fmt = d.Paragraph(p, doc).paragraph_format
            fmt.line_spacing = Pt(21)
            fmt.line_spacing_rule = WD_LINE_SPACING.EXACTLY
            fmt.space_before = Pt(0)
            fmt.space_after = Pt(int(after) / 20)
            fmt.page_break_before = False
            fmt.keep_with_next = p is blank
            fmt.keep_together = True
            spacing = p.get_or_add_pPr().find(qn("w:spacing"))
            for name, value in (("beforeLines", "0"), ("afterLines", "0" if p is blank else "50"),
                                ("beforeAutospacing", "0"), ("afterAutospacing", "0")):
                spacing.set(qn("w:" + name), value)
    return {"parts": len(headers), "description": HEADER_DESCRIPTION}


def verify_report_headers(doc):
    headers = list(_report_headers(doc))
    for _element, paragraphs, content in headers:
        if len(paragraphs) != 2 or paragraphs[1] is not content:
            raise ValueError("试验拆分输出页眉未保持两行，停止交付。")
        for index, p in enumerate(paragraphs):
            spacing = p.find("./" + qn("w:pPr") + "/" + qn("w:spacing"))
            if (spacing is None or spacing.get(qn("w:line")) != "420"
                    or spacing.get(qn("w:lineRule")) != "exact"
                    or spacing.get(qn("w:afterLines")) != ("50" if index else "0")):
                raise ValueError("试验拆分输出页眉的21磅行距或段后间距改变，停止交付。")
    return {"parts": len(headers), "description": HEADER_DESCRIPTION}


def _result_width_changes(doc):
    changes, count = [], 0
    for table in d._tables(doc):
        rows = d._rows(table)
        if not rows or not d._result_header_row(rows[0]):
            continue
        headers = list(d._cells(rows[0]))
        if [d._norm(d._text(cell)) for cell in headers] != list(RESULT_WIDTHS):
            raise ValueError("统一结果列宽需要明确的五列表头，停止输出。")
        grid = d._xp(table, "./w:tblGrid/w:gridCol")
        old = [int(node.get(qn("w:w"), "0")) for node in grid]
        if not old or min(old) <= 0:
            raise ValueError("结果表缺少有效列网格，不能统一列宽。")
        adjusted = []
        for column, span, cell in rp._grid_cells(rows[0]):
            weights = old[column:column + span]
            if len(weights) != span:
                raise ValueError("结果表头合并与列网格不一致，不能统一列宽。")
            label = d._norm(d._text(cell))
            total = RESULT_WIDTHS[label]
            if label == "项目名称":
                weights = [1] * span
            denominator = sum(weights)
            widths = [total * weight // denominator for weight in weights]
            order = sorted(range(span), key=lambda i: (-(total * weights[i] % denominator), i))
            for index in order[:total - sum(widths)]:
                widths[index] += 1
            if min(widths) <= 0:
                raise ValueError("结果表内部子列过窄，不能统一列宽。")
            adjusted.extend(widths)
        if len(adjusted) != len(grid):
            raise ValueError("结果表头未覆盖完整列网格，不能统一列宽。")
        changes.extend((node, "w:w", value) for node, value in zip(grid, adjusted))
        width = table.tblPr.find(qn("w:tblW"))
        if width is None:
            raise ValueError("结果表缺少总宽设置，不能统一列宽。")
        changes.extend(((width, "w:w", sum(RESULT_WIDTHS.values())), (width, "w:type", "dxa")))
        layout = table.tblPr.find(qn("w:tblLayout"))
        changes.append((table.tblPr, "layout", layout))
        for row in rows:
            cells = list(rp._grid_cells(row))
            if (not cells or cells[0][0] != 0
                    or cells[-1][0] + cells[-1][1] != len(grid)):
                raise ValueError("结果数据行与列网格不一致，不能统一列宽。")
            for column, span, cell in cells:
                width = cell.find("./" + qn("w:tcPr") + "/" + qn("w:tcW"))
                if width is None or column + span > len(grid):
                    raise ValueError("结果单元格宽度或合并结构不完整，不能统一列宽。")
                changes.extend(((width, "w:w", sum(adjusted[column:column + span])),
                                (width, "w:type", "dxa")))
        count += 1
    if not count:
        raise ValueError("没有可统一列宽的试验结果表，停止输出。")
    return changes, count


def result_widths(doc, *, enabled, project):
    if not enabled:
        return {"applied": False, "description": "结果表沿用源文件宽度。"}
    if d._norm(d._project_name(project)) == "指示装置":
        return {"applied": False, "description": "指示装置不调整，沿用源文件宽度。"}
    changes, count = _result_width_changes(doc)
    for node, attribute, value in changes:
        if attribute == "layout":
            node.get_or_add_tblLayout().set(qn("w:type"), "fixed")
        else:
            node.set(qn(attribute), str(value))
    return {"applied": True, "tables": count, "column_widths_twips": dict(RESULT_WIDTHS),
            "description": RESULT_WIDTH_DESCRIPTION}


def verify_result_widths(doc):
    changes, _ = _result_width_changes(doc)
    for node, attribute, value in changes:
        valid = (value is not None and value.get(qn("w:type")) == "fixed"
                 if attribute == "layout" else node.get(qn(attribute)) == str(value))
        if not valid:
            raise ValueError("Word分页改变了已指定的结果表列宽，停止交付。")


def _styles(style):
    seen = set()
    while style is not None and style.style_id not in seen:
        seen.add(style.style_id)
        yield style.element
        style = style.base_style


def _run_value(doc, paragraph, run, name, default=None):
    roots = [run._r, *_styles(run.style), *_styles(paragraph.style)]
    values = [value for root in roots for value in d._xp(root, "./w:rPr/" + name)]
    values += d._xp(doc.styles.element, "./w:docDefaults/w:rPrDefault/w:rPr/" + name)
    return values[0] if values else default


def _width(doc, paragraph, element, text):
    run = Run(element, paragraph)
    size = _run_value(doc, paragraph, run, "w:sz/@w:val")
    if size is None:
        raise ValueError("观察单行排版无法确定原字号，停止输出。")
    for script, allowed in (("ascii", {"Arial"}), ("eastAsia", {"宋体", "SimSun"})):
        value = _run_value(doc, paragraph, run, f"w:rFonts/@w:{script}")
        if value is not None and value not in allowed:
            raise ValueError("观察单行排版不支持该原字体的宽度测量：" + value)
    scale = float(_run_value(doc, paragraph, run, "w:w/@w:val", "100")) / 100
    spacing = float(_run_value(doc, paragraph, run, "w:spacing/@w:val", "0")) / 20
    return text_width(text, float(size) / 2) * scale + len(text) * spacing


def _indent(doc, paragraph):
    roots = [paragraph._p, *_styles(paragraph.style)]
    result = {}
    for side in ("left", "right", "firstLine", "hanging"):
        values = [v for root in roots for v in d._xp(root, "./w:pPr/w:ind/@w:" + side)]
        result[side] = int(values[0]) if values else 0
    return max(0, result["left"] + result["right"] + result["firstLine"] - result["hanging"])


def _margins(doc, table, cell):
    roots = [cell.tcPr, table.tblPr]
    roots += [node.find(qn("w:tblPr")) for node in _styles(Table(table, doc).style)]
    total = 0
    for side in ("left", "right"):
        values = [v for root in roots if root is not None
                  for v in d._xp(root, f"./w:tcMar/w:{side}/@w:w | ./w:tblCellMar/w:{side}/@w:w")]
        total += int(values[0]) if values else 108
    return total


def _set_widths(table, widths):
    for column, width in zip(d._xp(table, "./w:tblGrid/w:gridCol"), widths):
        column.set(qn("w:w"), str(width))
    for row in d._rows(table):
        for cell, width in zip(d._cells(row), widths):
            cell.get_or_add_tcPr().get_or_add_tcW().set(qn("w:w"), str(width))
            cell.tcPr.tcW.set(qn("w:type"), "dxa")
    table.tblPr.get_or_add_tblLayout().set(qn("w:type"), "fixed")


def basic_labels(doc):
    table = _basic_table(doc)
    rows = d._rows(table)
    if (len(d._xp(table, "./w:tblGrid/w:gridCol")) != 4
            or any(len(d._cells(row)) != 4 for row in rows)
            or d._xp(table, "./w:tr/w:tc/w:tcPr/w:gridSpan | ./w:tr/w:tc/w:tcPr/w:vMerge")):
        raise ValueError("观察基本信息不是同步的四列网格，停止输出。")
    for row in rows:
        for index, labels in ((0, {"申请单位", "规格型号"}), (2, {"计量器具名称", "样机编号"})):
            cell = d._cells(row)[index]
            if d._norm(d._text(cell)) not in labels:
                raise ValueError("观察基本信息标签无法识别：" + d._text(cell))
            if d._xp(cell, ".//w:fldChar | .//w:fldSimple | .//w:sdt | .//w:tbl | .//w:tab"):
                raise ValueError("观察标签含不能安全合并的域或嵌套结构。")
            paragraphs = d._xp(cell, "./w:p")
            if not paragraphs or d._xp(cell, ".//w:br[@w:type and @w:type!='textWrapping']"):
                raise ValueError("观察标签包含分页或分栏，不能当作文字换行移除。")
            for node in d._xp(cell, ".//w:br | .//w:cr"):
                node.getparent().remove(node)
            for paragraph in paragraphs[1:]:
                for node in list(paragraph):
                    if node.tag != qn("w:pPr"):
                        paragraphs[0].append(node)
                cell.remove(paragraph)
    _fit_basic_labels(doc)
    widths = [int(v) for v in d._xp(table, "./w:tblGrid/w:gridCol/@w:w")]
    for index in (0, 2):
        needed = widths[index]
        for row in rows:
            cell = d._cells(row)[index]
            paragraph = d.Paragraph(d._xp(cell, "./w:p")[0], doc)
            points = sum(_width(doc, paragraph, r._r, r.text) for r in paragraph.runs)
            needed = max(needed, ceil(points * 20) + _margins(doc, table, cell) + _indent(doc, paragraph) + 20)
        growth = needed - widths[index]
        if widths[index + 1] <= growth + 216:
            raise ValueError("观察标签在原字号/总宽内没有足够填写空间，停止输出。")
        widths[index] = needed
        widths[index + 1] -= growth
    _set_widths(table, widths)
    for row in rows:
        for index in (1, 3):
            cell = d._cells(row)[index]
            # Override inherited noWrap as well; long values are allowed to wrap.
            nowrap = cell.tcPr.find(qn("w:noWrap"))
            if nowrap is None:
                nowrap = OxmlElement("w:noWrap")
                cell.tcPr.append(nowrap)
            nowrap.set(qn("w:val"), "0")
        if d._norm(d._text(d._cells(row)[0])) == "规格型号":
            cell = d._cells(row)[3]
            paragraph = d.Paragraph(d._xp(cell, "./w:p")[0], doc)
            needed = max(sum(_width(doc, paragraph, r._r, part)
                             for r in paragraph.runs for part in [r.text.split("\n")[0]]), 0)
            if ceil(needed * 20) + _margins(doc, table, cell) + _indent(doc, paragraph) > widths[3]:
                raise ValueError("观察样机编号不能在原字号下独占一行，停止输出。")
    return widths


def _visible_nodes(element):
    return d._xp(element, ".//w:r/w:t | .//w:r/w:tab | .//w:r/w:br | .//w:r/w:cr")


def _node_text(node):
    return node.text or "" if node.tag == qn("w:t") else "\t" if node.tag == qn("w:tab") else "\n"


def _header_layout(doc, element, width, center):
    paragraphs = [p for p in d._xp(element, "./w:p") if "".join(_node_text(n) for n in _visible_nodes(p)).strip()]
    if not paragraphs:
        return None
    if len(paragraphs) != 1 or d._xp(element, ".//w:tbl | .//w:drawing | .//w:pict | .//w:object"):
        raise ValueError("观察页眉不是受支持的原生单段文字，停止输出。")
    p = paragraphs[0]
    paragraph = d.Paragraph(p, doc)
    nodes = _visible_nodes(p)
    text = "".join(_node_text(n) for n in nodes)
    match = re.fullmatch(r"\s*(附件\s*1)\s+(型式评价记录)\s+((?:报告|证书)编号[：:].*?\S)\s*", text)
    if not match:
        raise ValueError("观察页眉的附件/记录/编号结构无法识别，未改写原字段。")
    spans = [match.span(i) for i in (1, 2, 3)]
    widths = [0.0, 0.0, 0.0]
    space_width = None
    offset = 0
    for node in nodes:
        value = _node_text(node)
        run = next((a for a in node.iterancestors() if a.tag == qn("w:r")), None)
        if run is None:
            raise ValueError("观察页眉文字缺少原生字体运行，停止输出。")
        for index, (start, end) in enumerate(spans):
            piece = value[max(0, start - offset):max(0, min(len(value), end - offset))]
            widths[index] += _width(doc, paragraph, run, piece)
            if index == 1 and piece:
                space_width = _width(doc, paragraph, run, "   ")
        offset += len(value)
    if space_width is None or space_width <= 0:
        raise ValueError("观察页眉无法确定三个空格的原字号宽度，停止输出。")
    right_gap = width - 1 - widths[2] - center - widths[1] / 2
    title_center = center if right_gap >= 6 else (
        width - 1 - widths[2] - space_width - widths[1] / 2
    )
    gaps = [title_center - widths[1] / 2 - widths[0],
            width - 1 - widths[2] - title_center - widths[1] / 2]
    if gaps[0] < 6:
        raise ValueError("观察页眉在原字号下超过页面单行宽度，标题左移并保留三个空格后仍无法容纳；未缩字或截断编号。")
    return paragraph, nodes, match, spans, {
        "text": list(match.groups()), "width_pt": width, "text_widths_pt": widths,
        "center_pt": title_center, "page_center_pt": center,
        "number_gap_pt": gaps[1], "three_spaces_pt": space_width, "gap_pt": min(gaps),
    }


def _header(doc, element, width, center):
    layout = _header_layout(doc, element, width, center)
    if layout is None:
        return None
    paragraph, nodes, match, spans, result = layout
    offset = 0
    boundaries = {spans[0][1], spans[1][1]}
    for node in nodes:
        value = _node_text(node)
        replacement = "".join(ch if any(a <= offset + i < b for a, b in spans)
                              else "\t" if offset + i in boundaries else ""
                              for i, ch in enumerate(value))
        offset += len(value)
        parent, position = node.getparent(), node.getparent().index(node)
        parent.remove(node)
        for index, part in enumerate(replacement.split("\t")):
            if index:
                parent.insert(position, OxmlElement("w:tab"))
                position += 1
            if part:
                new = OxmlElement("w:t")
                new.set(qn("xml:space"), "preserve")
                new.text = part
                parent.insert(position, new)
                position += 1
    inherited = set()
    style = paragraph.style
    while style is not None:
        inherited.update(stop.position for stop in style.paragraph_format.tab_stops)
        style = style.base_style
    tabs = paragraph.paragraph_format.tab_stops
    tabs.clear_all()
    for position in sorted(inherited):
        tabs.add_tab_stop(position, WD_TAB_ALIGNMENT.CLEAR)
    tabs.add_tab_stop(Pt(result["center_pt"]), WD_TAB_ALIGNMENT.CENTER)
    tabs.add_tab_stop(Pt(width - 1), WD_TAB_ALIGNMENT.RIGHT)
    paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    for name in ("left_indent", "right_indent", "first_line_indent"):
        setattr(paragraph.paragraph_format, name, Pt(0))
    return result


def observe_format(doc):
    # The plan and export both budget against the final, shared report geometry.
    template = template_properties("observe")
    for section in doc.sections:
        for name in ("pgSz", "pgMar", "docGrid"):
            old, new = section._sectPr.find(qn("w:" + name)), template.find(qn("w:" + name))
            if new is not None:
                if old is not None:
                    section._sectPr.remove(old)
                section._sectPr.append(deepcopy(new))
    return {"basic_widths_twips": basic_labels(doc),
            "headers": normalize_attachment_headers(doc)}


def normalize_attachment_headers(doc):
    headers = {}
    for section in doc.sections:
        width = (section.page_width - section.left_margin - section.right_margin) / 12700
        center = (section.page_width / 2 - section.left_margin) / 12700
        for header in (section.header, section.first_page_header, section.even_page_header):
            if header.is_linked_to_previous:
                continue
            if not re.match(r"\s*附件\s*1", d._text(header._element)):
                continue
            headers[str(header.part.partname)] = (header._element, width, center)
    return [result for element, width, center in headers.values()
            if (result := _header(doc, element, width, center)) is not None]


def verify_observe_lines(doc, pdf):
    """Check physical PDF baselines, not just no-wrap XML settings."""
    from collections import Counter

    compact = lambda value: re.sub(r"\s+", "", value)
    labels = Counter(
        compact(d._text(d._cells(row)[index]))
        for row in d._rows(_basic_table(doc)) for index in (0, 2)
    )
    verify_header_lines(doc, pdf, labels=labels)


def verify_header_lines(doc, pdf, *, labels=None):
    from collections import Counter

    compact = lambda value: re.sub(r"\s+", "", value)
    labels = labels or Counter()
    expected_headers = {}
    for section in doc.sections:
        width = (section.page_width - section.left_margin - section.right_margin) / 12700
        center = (section.page_width / 2 - section.left_margin) / 12700
        for header in (section.header, section.first_page_header, section.even_page_header):
            if not re.match(r"\s*附件\s*1", d._text(header._element)):
                continue
            layout = _header_layout(doc, header._element, width, center)
            if layout is not None:
                result = layout[-1]
                expected_headers[compact("".join(result["text"]))] = (
                    section.left_margin.pt + result["center_pt"],
                    section.left_margin.pt + width - 1, result,
                )
    observed = Counter()
    for page in pdf:
        baselines = []
        for block in page.get_text("dict")["blocks"]:
            for line in block.get("lines", []):
                for span in line["spans"]:
                    x, y = span["origin"]
                    group = next((g for g in baselines if abs(g[0] - y) < 1), None)
                    if group is None:
                        group = [y, []]
                        baselines.append(group)
                    group[1].append((x, span["text"]))
        lines = [(y, compact("".join(text for _, text in sorted(parts))))
                 for y, parts in baselines]
        matched = [geometry for header, geometry in expected_headers.items()
                   if any(header in text for y, text in lines if y < page.rect.height / 4)]
        if expected_headers and not matched:
            raise ValueError(f"试验拆分输出第{page.number + 1}页页眉未完整单行显示，停止交付。")
        if expected_headers:
            titles = [rectangle for rectangle in page.search_for("型式评价记录")
                      if rectangle.y1 < page.rect.height / 4]
            if len(titles) != 1 or not any(
                    abs((titles[0].x0 + titles[0].x1) / 2 - center) <= 1
                    for center, _, _ in matched):
                raise ValueError(f"试验拆分输出第{page.number + 1}页页眉标题未按水平中心/长编号左移规则定位，停止交付。")
            if not any(
                    abs(rectangle.x1 - right) <= 1
                    and abs(rectangle.x0 - titles[0].x1 - result["number_gap_pt"]) <= 1
                    for _, right, result in matched
                    for rectangle in page.search_for(result["text"][2])
                    if rectangle.y1 < page.rect.height / 4):
                raise ValueError(f"试验拆分输出第{page.number + 1}页页眉编号未右对齐或标题间距不符，停止交付。")
        for _, text in lines:
            for label in labels:
                observed[label] += text.count(label)
    if labels - observed:
        raise ValueError("观察输出基本信息标签未完整单行显示，停止交付。")
