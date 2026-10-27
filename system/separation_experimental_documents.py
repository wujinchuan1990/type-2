"""Explicit, fail-closed sample plans for independent experimental fragments."""

from copy import deepcopy
import re

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pydantic import BaseModel, ConfigDict, Field, StrictInt, field_validator
from typing import List

import documents as d
import result_pagination as rp
import separation_documents as sd


class ModelGroup(BaseModel):
    model_config = ConfigDict(extra="forbid")
    model: str
    count: StrictInt = Field(ge=1, le=99)

    @field_validator("model")
    @classmethod
    def valid_model(cls, value):
        if not value.strip() or value != value.strip() or len(value) > 80 or re.search(r"[\x00-\x1f]", value):
            raise ValueError("型号须为1～80个非控制字符，不能有首尾空格。")
        return value


def roster(groups):
    return [{"model": group.model, "labels": [f"{chr(65 + i)}{n:03d}" for n in range(1, group.count + 1)]}
            for i, group in enumerate(groups)]


def selection(groups, selected, kind):
    if not 1 <= len(groups) <= 26 or sum(g.count for g in groups) > 120:
        raise ValueError("每方案支持1～26个型号、合计最多120台样机。")
    if len({g.model for g in groups}) != len(groups):
        raise ValueError("同一方案型号名称不能重复。")
    labels = [label for group in roster(groups) for label in group["labels"]]
    if len(selected) != len(set(selected)) or set(selected) - set(labels):
        raise ValueError("参试编号重复或不属于本方案型号/台数，不能采信前端编号。")
    if kind == "result" and not selected:
        raise ValueError("每个试验结果方案至少选择一台实际参试样机。")
    return [label for label in labels if label in selected] if kind == "result" else labels


def suggested_filename(part_id, kind, groups, selected, unit=None):
    if kind == "items":
        return "type evaluation items.docx"
    count = len(groups)
    counts = "".join(str(group.count) for group in groups)
    return f"{count} {'model' if count == 1 else 'models'}-{counts}.docx"


def result_subprojects(doc, unit):
    """Read name subcolumns only, never adjacent technical-condition columns."""
    subprojects = []
    for table in list(doc.element.body)[unit["start_block"]:unit["end_block"]]:
        if table.tag != qn("w:tbl"):
            continue
        rows = d._rows(table)
        if not rows or not d._result_header_row(rows[0]):
            continue
        headers = [(column, span) for column, span, cell in rp._grid_cells(rows[0])
                   if d._norm(d._text(cell)) == "项目名称"]
        if len(headers) != 1:
            raise ValueError("结果项目名称列边界不明确，不能确定子项目目录。")
        start, span = headers[0]
        end = start + span
        for row, (_physical, owners) in zip(rows[1:], list(d._effective_rows(table))[1:]):
            if d._result_header_row(row):
                continue
            names = []
            for (column, width, _cell), owner in zip(rp._grid_cells(row), owners):
                if column < end and column + width > start:
                    if column < start or column + width > end:
                        raise ValueError("结果数据格跨越项目名称列边界，不能猜测子项目。")
                    names.append(re.sub(r"\s+", "", d._text(owner)))
            if len(names) <= 1:
                continue
            if names[0] != re.sub(r"\s+", "", unit["project"]) or any(not name for name in names):
                raise ValueError("结果主项目或子项目名称不明确，不能生成可靠目录。")
            child = "—".join(names[1:])
            if child not in subprojects:
                subprojects.append(child)
    return subprojects


def result_project_title(unit):
    children = unit.get("subprojects", [])
    return unit["project"] + ("—" + "、".join(children) if children else "")


def result_project_folder(unit):
    if (unit.get("kind") != "result" or not re.fullmatch(r"[0-9]+", str(unit.get("sequence", "")))
            or not isinstance(unit.get("project"), str) or not unit["project"].strip()):
        raise ValueError("结果项目缺少明确的序号或名称，请核对源报告后再保存计划。")
    name = sd.default_filename({**unit, "project": result_project_title(unit)}, 1)[:-len("-1 model.docx")]
    if len(name.encode("utf-8")) > 240:
        raise ValueError("结果项目文件夹名称过长，请核对源报告的项目名称。")
    return name


def _plain_cell(cell):
    if d._xp(cell, ".//w:drawing | .//w:pict | .//w:object | .//w:fldChar | .//w:fldSimple | "
                   ".//w:sdt | .//w:ins | .//w:del | .//w:tbl | .//w:commentReference | "
                   ".//w:bookmarkStart | .//w:bookmarkEnd | .//w:sectPr | "
                   ".//*[local-name()='oMath' or local-name()='oMathPara']"):
        raise ValueError("待重建的样机行含图片、域、书签、修订或嵌套结构，不能安全复制/清空。")


def _remove_ephemeral_bookmarks(doc):
    references = " ".join(d._xp(doc.element, ".//w:instrText/text() | .//w:fldSimple/@w:instr"))
    links = set(d._xp(doc.element, ".//w:hyperlink/@w:anchor"))
    identities = {node.get(qn("w:id")) for node in d._xp(doc.element, ".//w:bookmarkStart")
                  if node.get(qn("w:name"), "").startswith(("_Hlk", "_GoBack", "OLE_LINK"))
                  and node.get(qn("w:name")) not in references and node.get(qn("w:name")) not in links}
    for node in d._xp(doc.element, ".//w:bookmarkStart | .//w:bookmarkEnd"):
        if node.get(qn("w:id")) in identities:
            node.getparent().remove(node)


def observe(doc, groups):
    _remove_ephemeral_bookmarks(doc)
    candidates = []
    for table in d._tables(doc):
        for row in d._rows(table):
            cells = d._cells(row)
            if cells and d._norm(d._text(cells[0])) == "规格型号":
                candidates.append((table, row, cells))
    if not candidates:
        return _flow_observe(doc, groups)
    if len({id(t) for t, _, _ in candidates}) != 1:
        raise ValueError("观察顶部规格型号行不能唯一定位。")
    table = candidates[0][0]
    rows = d._rows(table)
    last = max(rows.index(row) for _, row, _ in candidates)
    for row in rows[last + 1:]:
        cells = d._cells(row)
        if (len(cells) == 4 and not d._text(cells[0]).strip() and not d._text(cells[2]).strip()
                and d._text(cells[1]).strip() and d._text(cells[3]).strip()):
            candidates.append((table, row, cells))
        else:
            break
    positions = [rows.index(row) for _, row, _ in candidates]
    if positions != list(range(positions[0], positions[-1] + 1)):
        raise ValueError("规格型号行不连续，不能移除夹在其中的其他信息。")
    for _, row, cells in candidates:
        if (len(cells) != 4 or d._norm(d._text(cells[2])) not in {"样机编号", ""}
                or d._xp(row, "./w:tc/w:tcPr/w:vMerge")):
            raise ValueError("观察基本信息须为四列规格型号/型号/样机编号/编号列表，且不能纵向合并。")
        for cell in cells:
            _plain_cell(cell)
    original = candidates[0][1]
    insertion = table.index(original)
    for offset, group in enumerate(roster(groups)):
        row = deepcopy(original)
        cells = d._cells(row)
        d._put(cells[1], group["model"])
        d._put(cells[3], "\n".join(group["labels"]))
        for cell in cells:
            cell.get_or_add_tcPr().get_or_add_vAlign().set(qn("w:val"), "center")
        for height in d._xp(row, "./w:trPr/w:trHeight"):
            height.set(qn("w:hRule"), "atLeast")
        table.insert(insertion + offset, row)
    for _, row, _ in candidates:
        table.remove(row)
    return {"models": roster(groups)}


def _flow_observe(doc, groups):
    matches = [(table, index) for table in d._tables(doc)
               for index, row in enumerate(d._rows(table))
               if [d._norm(d._text(c)) for c in d._cells(row)] ==
               ["序号", "型号", "口径", "准确度等级", "样机编号"]]
    if len(matches) != 1:
        raise ValueError("观察顶部规格型号行不能唯一定位。")
    table, index = matches[0]
    rows = d._rows(table)
    if index != 1 or len(d._cells(rows[0])) != 4:
        raise ValueError("流量计观察顶部不是已识别的申请单位/五列样机表结构。")
    grid = d._xp(table, "./w:tblGrid/w:gridCol")
    widths = [int(column.get(qn("w:w"))) for column in grid]
    spans = [int(next(iter(d._xp(cell, "./w:tcPr/w:gridSpan/@w:val")), "1"))
             for cell in d._cells(rows[0])]
    if len(widths) != 5 or spans != [1, 2, 1, 1]:
        raise ValueError("流量计申请单位行与五列网格不一致，不能安全转换。")
    new_widths = [widths[0], widths[1] + widths[2], widths[3], widths[4]]
    for column in grid:
        table.tblGrid.remove(column)
    for width in new_widths:
        column = OxmlElement("w:gridCol")
        column.set(qn("w:w"), str(width))
        table.tblGrid.append(column)
    for cell, width in zip(d._cells(rows[0]), new_widths):
        for span in d._xp(cell, "./w:tcPr/w:gridSpan"):
            span.getparent().remove(span)
        cell.get_or_add_tcPr().get_or_add_tcW().set(qn("w:w"), str(width))
        cell.tcPr.tcW.set(qn("w:type"), "dxa")
    for row in rows[1:]:
        if len(d._cells(row)) != 5:
            raise ValueError("流量计样机表混有其他信息，不能安全替换。")
        for cell in d._cells(row):
            _plain_cell(cell)
    # Diameter and accuracy were not authorized for deletion. Keep their native
    # paragraphs separately; do not imply a mapping to the placeholder models.
    extras = OxmlElement("w:p")
    d._put(extras, "原口径、准确度等级（保留源顺序，未映射至占位型号）：")
    table.addnext(extras)
    for source in rows[2:]:
        for cell in d._cells(source)[2:4]:
            for paragraph in d._xp(cell, "./w:p"):
                new = deepcopy(paragraph)
                extras.addnext(new)
                extras = new
    for row in rows[1:]:
        table.remove(row)
    for group in roster(groups):
        row = deepcopy(rows[0])
        for cell, text in zip(d._cells(row), ("规格型号：", group["model"], "样机编号", "\n".join(group["labels"]))):
            d._put(cell, text)
            cell.get_or_add_tcPr().get_or_add_vAlign().set(qn("w:val"), "center")
        table.append(row)
    return {"models": roster(groups), "preserved_source_specs": True}


def observe_record_samples(doc, labels):
    import observation_pagination as op

    replacements = []
    for table in d._tables(doc):
        rows = d._rows(table)
        if not rows or "型式评价大纲章节号及项目名称" not in d._text(rows[0]):
            continue
        op._layout(table)
        headers = list(op._columns(rows[0]))
        if len(headers) == 5:
            continue
        if d._norm(d._text(headers[2][2])) != "样机编号":
            raise ValueError("观察项目记录的样机编号列不明确，不能替换原编号。")
        sample_span = headers[2][:2]
        source_lists = {}
        for _table, row, key in op._project_rows([table]):
            cells = [cell for start, stop, cell in op._columns(row) if (start, stop) == sample_span]
            if len(cells) != 1:
                raise ValueError("观察项目记录的样机编号格与表头网格不一致，不能替换。")
            cell = cells[0]
            value = d._text(cell).strip()
            if not value or value in {"/", "／", "-", "－", "—", "不适用", "不涉及"}:
                continue
            if d._xp(cell, "./w:tcPr/w:vMerge"):
                raise ValueError("观察项目记录的样机编号格有纵向合并，不能猜测各项目编号范围。")
            project = key[1][1]
            normalized = d._norm(value)
            if project in source_lists and source_lists[project] != normalized:
                raise ValueError("同一观察子项目的编号分散在不同清单中，不能逐格猜测重建编号范围。")
            source_lists[project] = normalized
            _plain_cell(cell)
            replacements.append(cell)
    for cell in replacements:
        d._put(cell, "\n".join(labels))
    return {"replaced_cells": len(replacements), "labels": labels}


def result_blocks(doc):
    """Track every grid column, explicit condition restart, and table continuation."""
    blocks, current = [], None
    previous_table = None
    for table in d._tables(doc):
        rows = d._rows(table)
        if not rows or not d._result_header_row(rows[0]):
            continue
        header = [d._norm(d._text(c)) for c in d._cells(rows[0])]
        if header != ["序号", "项目名称", "技术要求", "样机编号", "试验结果"]:
            raise ValueError("结果表头不是受支持的序号/项目名称/技术要求/样机编号/试验结果结构。")
        expected_width = sum(int(x) for x in d._xp(table, "./w:tblGrid/w:gridCol/@w:w"))
        if expected_width <= 0:
            raise ValueError("结果表缺少原生列宽。")
        for ri, (physical, effective) in enumerate(d._effective_rows(table)):
            if ri == 0:
                continue
            if d._result_header_row(rows[ri]):
                raise ValueError("表内嵌入重复表头，请先按原生续表分隔后重试。")
            if len(physical) < 5 or len(physical) != len(effective):
                raise ValueError("结果表数据行列结构不明确。")
            for cell in physical:
                _plain_cell(cell)
            if any(d._xp(c, "./w:tcPr/w:vMerge") for c in physical[-2:]):
                raise ValueError("样机/实测格纵向合并，不能安全重建。")
            label = d._text(physical[-2]).strip()
            if not label or len(label) > 100:
                raise ValueError("结果表行缺少可靠的源样机编号。")
            spans = tuple(int(next(iter(d._xp(c, "./w:tcPr/w:gridSpan/@w:val")), "1")) for c in physical)
            trailing = int(next(iter(d._xp(rows[ri], "./w:trPr/w:gridAfter/@w:val")), "0"))
            if sum(spans) + trailing != len(d._xp(table, "./w:tblGrid/w:gridCol")):
                raise ValueError("结果行网格与表格列宽不一致。")
            values = tuple(d._text(c) for c in effective[:-2])
            if not values[0].strip() or not values[1].strip() or not any(v.strip() for v in values[2:]):
                raise ValueError("结果行序号、子项目或技术条件为空，不能猜测源条件。")
            key = (spans, values)
            # Deepest requirement column is the leaf condition, not the table or
            # a flattened requirement string. A fresh merge is a distinct test.
            fresh_leaf = d._xp(physical[-3], "./w:tcPr/w:vMerge/@w:val") == ["restart"]
            same = current and current["key"] == key and label not in current["source_labels"]
            continuation = table is not previous_table
            if continuation and previous_table is not None:
                sibling = previous_table.getnext()
                while sibling is not None and sibling is not table:
                    if d._text(sibling).strip() or sibling.tag != qn("w:p"):
                        same = False
                        break
                    sibling = sibling.getnext()
            if not same or (fresh_leaf and not continuation):
                current = {"key": key, "rows": [], "owners": effective[:-2],
                           "source_labels": [], "table": table}
                blocks.append(current)
            current["rows"].append(rows[ri])
            current["source_labels"].append(label)
            previous_table = table
    if not blocks:
        raise ValueError("未找到可安全识别的原生试验结果表；不支持的结构不会伪造成功。")
    return blocks


def results(doc, labels):
    _remove_ephemeral_bookmarks(doc)
    generated_separators = set(d._split_repeated_project_headers(doc))
    blocks = result_blocks(doc)
    for block in blocks:
        source = block["rows"][0]
        table = source.getparent()
        insertion = table.index(source)
        for index, label in enumerate(labels):
            row = deepcopy(source)
            cells = d._cells(row)
            for cell, owner in zip(cells[:-2], block["owners"]):
                for child in list(cell):
                    if child.tag != qn("w:tcPr"):
                        cell.remove(child)
                if index == 0:
                    for child in owner:
                        if child.tag != qn("w:tcPr"):
                            cell.append(deepcopy(child))
                else:
                    cell.append(OxmlElement("w:p"))
                merge = cell.get_or_add_tcPr().get_or_add_vMerge()
                merge.set(qn("w:val"), "restart" if index == 0 else "continue")
            d._put(cells[-2], label)
            d._put(cells[-1], "")
            for height in d._xp(row, "./w:trPr/w:trHeight"):
                height.set(qn("w:hRule"), "atLeast")
            table.insert(insertion + index, row)
        for row in block["rows"]:
            row.getparent().remove(row)
    for table in list(d._tables(doc)):
        if len(d._rows(table)) == 1 and d._result_header_row(d._rows(table)[0]):
            separator = table.getprevious()
            if separator in generated_separators:
                separator.getparent().remove(separator)
            table.getparent().remove(table)
    rebuilt = result_blocks(doc)
    if len(rebuilt) != len(blocks) or any(b["source_labels"] != labels for b in rebuilt):
        raise ValueError("重建后子条件数量或参试编号不完整，停止输出。")
    if any(d._text(row["_value_cell"]) for row in d._result_schema(doc)):
        raise ValueError("存在未清空的实测结果格，停止输出。")
    # Discard only generated pagination, never explicit source page boundaries.
    previous = None
    for table in d._tables(doc):
        if not d._rows(table) or not d._result_header_row(d._rows(table)[0]):
            previous = None
            continue
        between = []
        node = previous.getnext() if previous is not None else None
        while node is not None and node is not table and node.tag == qn("w:p") and not d._text(node).strip():
            between.append(node)
            node = node.getnext()
        if node is table:
            protected = ".//w:drawing | .//w:pict | .//w:object | .//w:sectPr | " \
                        ".//w:fldChar | .//w:fldSimple | .//w:instrText | .//w:bookmarkStart | " \
                        ".//w:bookmarkEnd | .//w:sdt | .//w:commentRangeStart | .//w:commentRangeEnd | " \
                        ".//w:commentReference | .//w:footnoteReference | .//w:endnoteReference"
            page_break = ".//w:br[@w:type='page'] | ./w:pPr/w:pageBreakBefore[" \
                         "not(@w:val) or @w:val='true' or @w:val='1' or @w:val='on']"
            if any(d._xp(p, protected) or (p not in generated_separators and d._xp(p, page_break))
                   for p in between):
                previous = table
                continue
            for paragraph in between:
                paragraph.getparent().remove(paragraph)
            # Equivalent grids can be one flowing table, avoiding artificial
            # table boundaries that otherwise prevent Word from filling a page.
            if d._xp(previous, "./w:tblGrid/w:gridCol/@w:w") == d._xp(table, "./w:tblGrid/w:gridCol/@w:w"):
                for row in d._rows(table)[1:]:
                    previous.append(row)
                table.getparent().remove(table)
                table = previous
        previous = table
    from separation_result_format import merge_result_cells, normalize_result_row_heights

    normalize_result_row_heights(doc)
    merge_result_cells(doc)
    return {"conditions": len(blocks), "rows": len(blocks) * len(labels), "selected": labels,
            "requirements": [list(b["key"][1]) for b in blocks]}


def planning_fragment(doc, start, end):
    from separation_documents import effective_sections

    sections = effective_sections(doc.element.body)
    last = next((s for s in sections if s[0] <= end - 1 <= s[1]), None)
    if last is None:
        raise ValueError("试验拆分方案无法定位片段分节及页眉。")
    fragment = d._source_fragment(doc, list(doc.element.body)[start:end])
    body = fragment.element.body
    body.replace(body[-1], deepcopy(last[2]))
    for _begin, stop, properties in sections:
        if start <= stop < end - 1:
            paragraph = body[stop - start]
            if paragraph.tag != qn("w:p"):
                raise ValueError("试验拆分方案内部的分节位置不是段落。")
            paragraph.get_or_add_pPr().append(deepcopy(properties))
    return fragment


def transform(doc, kind, groups, selected):
    from separation_experimental_format import (
        normalize_attachment_headers, normalize_report_headers, normalize_result_closing,
        normalize_result_intro, normalize_result_titles,
    )

    headers = normalize_report_headers(doc)
    if kind != "observe":
        normalize_attachment_headers(doc)
    if kind == "items":
        return {"unchanged": True, "header_format": headers}
    labels = selection(groups, selected, kind)
    if kind == "observe":
        from separation_experimental_format import observe_format
        from gas_observation_flow import split as split_gas_flow

        structure = observe(doc, groups)
        flow_ranges = split_gas_flow(doc, len(groups))
        if flow_ranges is not None:
            structure["gas_flow_ranges"] = flow_ranges
        structure["record_samples"] = observe_record_samples(doc, labels)
        structure["format"] = observe_format(doc)
        structure["header_format"] = headers
        return structure
    return {**results(doc, labels), "header_format": headers,
            "title_format": normalize_result_titles(doc),
            "intro_format": normalize_result_intro(doc),
            "closing_format": normalize_result_closing(doc)}
