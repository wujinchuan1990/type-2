"""Shared result-only formatting for newly exported split copies."""

from copy import deepcopy
from pathlib import Path
import re
import unicodedata

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree

import documents as d
import result_pagination as rp
import signature_templates as signature


def _content(cell):
    content = []
    for node in cell:
        if node.tag == qn("w:tcPr"):
            continue
        copy = deepcopy(node)
        for child in copy.iter():
            for name in list(child.attrib):
                if etree.QName(name).localname.startswith("rsid") or etree.QName(name).localname in {
                        "paraId", "textId"}:
                    del child.attrib[name]
        content.append(etree.tostring(copy, method="c14n"))
    return tuple(content)


def _effective_content(table):
    return [tuple(_content(cell) for cell in effective)
            for _physical, effective in d._effective_rows(table)]


def _merge_signature(table, allowed):
    return [tuple((column, width, d._text(owner) if
                   set(range(column, column + width)).issubset(allowed) else _content(owner))
                  for (column, width, _cell), owner in zip(rp._grid_cells(row), owners))
            for row, (_physical, owners) in zip(d._rows(table), d._effective_rows(table))]


def merge_result_cells(doc):
    merged = 0
    for table in d._tables(doc):
        rows = d._rows(table)
        if not rows or not d._result_header_row(rows[0]):
            continue
        allowed = rp._allowed_columns(table)
        effective = [[(d._text(cell), cell) for cell in owners]
                     for _physical, owners in d._effective_rows(table)]
        before = _merge_signature(table, allowed)
        previous = {}
        merged_owners = {}
        for row, owners in zip(rows[1:], effective[1:]):
            current = {}
            for (column, width, cell), (text, owner) in zip(rp._grid_cells(row), owners):
                if not set(range(column, column + width)).issubset(allowed):
                    continue
                key = (width, text)
                prior = previous.get(column)
                if not text.strip():
                    continue
                if prior is not None and prior[0] == key:
                    first = prior[1]
                    if not rp._continued(cell):
                        protected = (
                            ".//w:drawing | .//w:pict | .//w:object | .//w:fldChar | .//w:fldSimple | "
                            ".//w:sdt | .//w:ins | .//w:del | .//w:tbl | .//w:hyperlink | "
                            ".//w:bookmarkStart | .//w:bookmarkEnd | .//w:commentReference | "
                            ".//w:footnoteReference | .//w:endnoteReference | "
                            ".//*[local-name()='oMath' or local-name()='oMathPara']"
                        )
                        if d._xp(first, protected) or d._xp(owner, protected):
                            raise ValueError("同文字结果格包含图片、域、书签或其他非普通文字，不能仅按文字安全合并。")
                        paragraph = deepcopy(d._xp(cell, "./w:p")[0])
                        for child in list(paragraph):
                            if child.tag != qn("w:pPr"):
                                paragraph.remove(child)
                        for child in list(cell):
                            if child.tag != qn("w:tcPr"):
                                cell.remove(child)
                        cell.append(paragraph)
                        merged += 1
                    first.get_or_add_tcPr().get_or_add_vMerge().set(qn("w:val"), "restart")
                    cell.get_or_add_tcPr().get_or_add_vMerge().set(qn("w:val"), "continue")
                    if owner is not first:
                        merged_owners[owner] = first
                    current[column] = (key, first)
                else:
                    while owner in merged_owners:
                        owner = merged_owners[owner]
                    current[column] = (key, owner)
            previous = current
        if _merge_signature(table, allowed) != before:
            raise ValueError("结果单元格合并改变了原文字、列结构或样机/实测格内容，停止导出。")
    return merged


def _value(value):
    text = re.sub(r"\s+", "", unicodedata.normalize("NFKC", value)).replace("～", "~")
    text = re.sub(r"(?:℃|°C|%RH|%|kPa|Pa)$", "", text)
    return text.strip("()（）")


def normalize_result_row_heights(doc):
    for table in d._tables(doc):
        rows = d._rows(table)
        if not rows or not d._result_header_row(rows[0]):
            continue
        for row in rows:
            properties = row.get_or_add_trPr()
            height = properties.get_or_add_trHeight()
            for duplicate in list(properties.findall(qn("w:trHeight")))[1:]:
                properties.remove(duplicate)
            height.set(qn("w:val"), str(d.PROJECT_METADATA_ROW_HEIGHT_TWIPS))
            height.set(qn("w:hRule"), "atLeast")


def format_metadata(doc):
    tables = [table for table in d._tables(doc)
              if all(label in d._text(table) for label in ("环境温度", "评价人员", "复核人员"))]
    for reference in tables:
        if d._xp(reference, ".//w:drawing | .//w:pict | .//w:object | .//w:fldChar | .//w:fldSimple | "
                            ".//w:bookmarkStart | .//w:bookmarkEnd | .//w:ins | .//w:del | .//w:tbl"):
            raise ValueError("拆分结果页底含图片、域、书签、修订或嵌套表，不能安全替换V1模板。")
        slots = signature._slots(reference)
        required = set(signature.LABELS) - {"大气压力"}
        if not required.issubset(slots):
            raise ValueError("拆分结果页底缺少设备、温湿度或人员标签，不能猜测套用V1模板。")
        expected = {label: _value(slot[3]) for label, slot in slots.items()}
        enabled = "大气压力" in slots
        project = {"atmospheric_pressure": slots["大气压力"][3].strip()} if enabled else None
        # Work on one metadata block, so combined projects keep their own values.
        fragment = d._source_fragment(doc, [reference])
        signature.apply_signature_templates(
            fragment, Path(__file__).resolve().parents[1], enabled, project=project,
            layout_settings=doc.settings.element)
        replacement = next(table for table in d._tables(fragment)
                           if signature.is_signature_table(table))
        actual = {label: _value(slot[3]) for label, slot in signature._slots(replacement).items()}
        if actual != expected:
            raise ValueError("V1页底模板替换改变了拆分结果的原设备、环境或人员值，停止导出。")
        from docxcompose.composer import Composer

        position = doc.element.body.index(reference)
        doc.element.body.remove(reference)
        d._insert(Composer(doc), position, fragment)


def format_results(doc):
    from separation_layout import zero_heading_spacing

    normalize_result_row_heights(doc)
    merge_result_cells(doc)
    format_metadata(doc)
    for paragraph in d._xp(doc.element.body, "./w:p"):
        if re.match(r"^型式评价结果\s*[-—－–]", d._text(paragraph).strip()):
            zero_heading_spacing(paragraph)
    body = doc.element.body
    if (len(body) >= 2 and body[-2].tag == qn("w:tbl")
            and rp._metadata_caption(body[-2]) == rp.METADATA_CAPTION):
        paragraph = OxmlElement("w:p")
        spacing = paragraph.get_or_add_pPr().get_or_add_spacing()
        for key, value in (("before", "0"), ("after", "0"), ("line", "1"), ("lineRule", "exact")):
            spacing.set(qn("w:" + key), value)
        body[-1].addprevious(paragraph)
    if (len(body) < 3 or body[-3].tag != qn("w:tbl")
            or rp._metadata_caption(body[-3]) != rp.METADATA_CAPTION):
        return
    paragraph = body[-2]
    if (paragraph.tag != qn("w:p") or any(child.tag != qn("w:pPr") for child in paragraph)
            or d._xp(paragraph, "./w:pPr/w:sectPr")
            or not d._xp(paragraph, "./w:pPr/w:spacing[@w:line='1'][@w:lineRule='exact']")):
        return
    table = body[-3]
    floating = table.tblPr.find(qn("w:tblpPr"))
    heights = d._xp(table, "./w:tr/w:trPr/w:trHeight")
    if (floating is None or floating.get(qn("w:vertAnchor")) != "margin"
            or len(heights) != 3
            or any(height.get(qn("w:hRule")) != "exact"
                   or height.get(qn("w:val")) != str(d.PROJECT_METADATA_ROW_HEIGHT_TWIPS)
                   for height in heights)):
        raise ValueError("拆分结果页底高度或锚点不明确，不能安全预留表后段落空间。")
    section = doc.sections[-1]
    if any(value is None for value in (section.page_height, section.top_margin, section.bottom_margin)):
        raise ValueError("拆分结果正文区尺寸不明确，不能猜测页底位置。")
    available = int((section.page_height - section.top_margin - section.bottom_margin) / 635)
    # Include the nested table's 1-twip terminator and the visible final paragraph's 1-pt font.
    position = available - sum(int(height.get(qn("w:val"))) for height in heights) - 1 - 20
    if position < 0:
        raise ValueError("拆分结果正文区不足以容纳完整页底和表后段落，停止导出。")
    floating.attrib.pop(qn("w:tblpYSpec"), None)
    floating.set(qn("w:tblpY"), str(position))
    for setting in ("keepNext", "keepLines", "pageBreakBefore", "widowControl"):
        d._paragraph_setting(paragraph, setting, False)
    properties = paragraph.get_or_add_pPr()
    snap = properties.find(qn("w:snapToGrid"))
    if snap is None:
        snap = OxmlElement("w:snapToGrid")
        properties.append(snap)
    snap.set(qn("w:val"), "false")
    run_properties = properties.find(qn("w:rPr"))
    if run_properties is None:
        run_properties = OxmlElement("w:rPr")
        properties.append(run_properties)
    for name in ("sz", "szCs"):
        size = run_properties.find(qn("w:" + name))
        if size is None:
            size = OxmlElement("w:" + name)
            run_properties.append(size)
        size.set(qn("w:val"), "2")
    hidden = run_properties.find(qn("w:vanish"))
    if hidden is not None:
        run_properties.remove(hidden)
