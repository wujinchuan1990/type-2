"""Continue result tables using measured Word pages, never estimated row counts."""

from copy import deepcopy
from itertools import count
from pathlib import Path
import shutil
import subprocess
import sys
from tempfile import TemporaryDirectory
import threading
import time
import uuid

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import documents as d
from native_word import render


TABLE_PREFIX = "rp_table_"
ROW_PREFIX = "rp_row_"
OVERFLOW_PREFIX = "rp_group_overflow_"
METADATA_CAPTION = "report-project-metadata"
METADATA_GROUP_CAPTION = "report-project-metadata-group"
HEIGHT_PREFIX = "rh_equal_"

EQUAL_HEIGHT_SCRIPT = r'''
on run argv
    set expectedName to POSIX file (item 1 of argv) as text
    tell application "Microsoft Word"
        set d to document (item 2 of argv)
        if full name of d is not expectedName then error "Document path mismatch."
        repeat with i from 3 to count argv by 3
            set firstName to item i of argv
            set lastName to item (i + 1) of argv
            set expectedRows to item (i + 2) of argv as integer
            set a to start of bookmark of bookmark firstName of d
            set b to start of bookmark of bookmark lastName of d
            set firstRange to create range d start a end a
            set lastRange to create range d start b end b
            set firstCell to cell 1 of firstRange
            set lastCell to cell 1 of lastRange
            if (row index of lastCell) - (row index of firstCell) + 1 is not expectedRows then error "Shared requirement row mapping changed."
            set e to end of content of text object of lastCell
            set groupRange to create range d start a end e
            set selectedRows to row options of groupRange
            distribute row height selectedRows
            set commonHeight to height of firstCell
            if commonHeight < 0 or commonHeight > 10000 then error "Word returned invalid shared row height."
            if commonHeight < 22.7 then set commonHeight to 22.7
            set height rule of selectedRows to row height at least
            set height of selectedRows to commonHeight
            set heightDifference to (height of selectedRows) - commonHeight
            if heightDifference < -0.01 or heightDifference > 0.01 then error "Shared requirement rows did not receive equal minimum heights."
        end repeat
        repaginate d
        save d
        return ((count argv) - 2) div 3
    end tell
end run
'''


def _metadata_caption(table):
    values = d._xp(table, "./w:tblPr/w:tblCaption/@w:val")
    return values[0] if values else ""


def _metadata_groups(doc):
    return [table for table in d._tables(doc) if _metadata_caption(table) == METADATA_GROUP_CAPTION]


def prepare_metadata_groups(path):
    """A single non-splittable outer row keeps all three editable inner rows together."""
    doc = d._load(path)
    for table in d._tables(doc):
        if _metadata_caption(table) != METADATA_CAPTION:
            continue
        rows = d._rows(table)
        floating = table.tblPr.find(qn("w:tblpPr"))
        if len(rows) != 3 or floating is None:
            raise ValueError("项目页底信息组不是三行页底表格，无法保证整体同页。")
        width = sum(int(value) for value in d._xp(table, "./w:tblGrid/w:gridCol/@w:w"))
        if width <= 0:
            raise ValueError("项目页底信息组缺少有效宽度。")
        outer = OxmlElement("w:tbl")
        properties = OxmlElement("w:tblPr")
        properties.append(deepcopy(floating))
        for name, attributes in (
            ("tblOverlap", {"val": "never"}),
            ("tblW", {"w": str(width), "type": "dxa"}),
            ("jc", {"val": "left"}),
            ("tblInd", {"w": "0", "type": "dxa"}),
            ("tblLayout", {"type": "fixed"}),
            ("tblCaption", {"val": METADATA_GROUP_CAPTION}),
        ):
            node = OxmlElement("w:" + name)
            for key, value in attributes.items():
                node.set(qn("w:" + key), value)
            properties.append(node)
        borders, margins = OxmlElement("w:tblBorders"), OxmlElement("w:tblCellMar")
        for side in ("top", "left", "bottom", "right"):
            border, margin = OxmlElement("w:" + side), OxmlElement("w:" + side)
            border.set(qn("w:val"), "nil")
            margin.set(qn("w:w"), "0")
            margin.set(qn("w:type"), "dxa")
            borders.append(border)
            margins.append(margin)
        properties.insert(5, borders)
        properties.insert(len(properties) - 1, margins)
        grid, column = OxmlElement("w:tblGrid"), OxmlElement("w:gridCol")
        column.set(qn("w:w"), str(width))
        grid.append(column)
        row, row_properties = OxmlElement("w:tr"), OxmlElement("w:trPr")
        row_properties.append(OxmlElement("w:cantSplit"))
        heights = d._xp(table, "./w:tr/w:trPr/w:trHeight")
        if (len(heights) != 3 or any(height.get(qn("w:hRule")) != "exact"
                                    or height.get(qn("w:val")) != str(d.PROJECT_METADATA_ROW_HEIGHT_TWIPS)
                                    for height in heights)):
            raise ValueError("项目页底三行必须各为固定0.8cm，无法安全预留整体高度。")
        # Floating-table positioning otherwise measures only the portion that fits,
        # even with cantSplit, allowing a nested personnel row to spill.
        height = OxmlElement("w:trHeight")
        height.set(qn("w:val"), str(sum(int(item.get(qn("w:val"))) for item in heights) + 1))
        height.set(qn("w:hRule"), "atLeast")
        row_properties.append(height)
        row.append(row_properties)
        cell, cell_properties = OxmlElement("w:tc"), OxmlElement("w:tcPr")
        cell_width = OxmlElement("w:tcW")
        cell_width.set(qn("w:w"), str(width))
        cell_width.set(qn("w:type"), "dxa")
        cell_properties.append(cell_width)
        cell.append(cell_properties)
        table.addprevious(outer)
        table.tblPr.remove(floating)
        cell.append(table)
        # Word requires a paragraph after a nested table; it must not add a visible line.
        terminator = OxmlElement("w:p")
        spacing = terminator.get_or_add_pPr().get_or_add_spacing()
        for key, value in (("before", "0"), ("after", "0"), ("line", "1"), ("lineRule", "exact")):
            spacing.set(qn("w:" + key), value)
        for setting in ("keepNext", "keepLines", "pageBreakBefore"):
            d._paragraph_setting(terminator, setting, False)
        snap = OxmlElement("w:snapToGrid")
        snap.set(qn("w:val"), "false")
        terminator.get_or_add_pPr().append(snap)
        cell.append(terminator)
        row.append(cell)
        outer.extend([properties, grid, row])
        following = outer.getnext()
        if following is not None and following.tag == qn("w:sectPr"):
            outer.addnext(deepcopy(terminator))
    doc.save(path)


def _measure_metadata(doc):
    tables = []
    for table in _metadata_groups(doc):
        inner = d._xp(table, "./w:tr/w:tc/w:tbl")
        if len(inner) != 1 or len(d._rows(inner[0])) != 3:
            raise ValueError("项目页底整体结构已变化，无法核对三行同页。")
        tables.append(inner[0])
    return _measure_metadata_tables(doc, tables)


def _measure_metadata_tables(doc, tables):
    groups = []
    existing_ids = [int(value) for value in d._xp(doc.element, ".//w:bookmarkStart/@w:id") if value.isdigit()]
    identities = count(max(existing_ids, default=-1) + 1)
    for index, table in enumerate(tables):
        if len(d._rows(table)) != 3:
            raise ValueError("项目页底整体结构已变化，无法核对三行同页。")
        names = []
        for row_index, row in enumerate(d._rows(table)):
            for cell_index, cell in enumerate(d._cells(row)):
                paragraphs = d._xp(cell, ".//w:p")
                if not paragraphs:
                    raise ValueError("项目页底单元格缺少段落，无法检查人员及环境内容。")
                base = f"{ROW_PREFIX}metadata_{index}_{row_index}_{cell_index}"
                for suffix, paragraph in (("_s", paragraphs[0]), ("_e", paragraphs[-1])):
                    name = base + suffix
                    _bookmark(doc, paragraph, name, end=suffix == "_e", identity=next(identities))
                    names.append(name)
        groups.append(names)
    return groups


def _validate_metadata_pages(groups, bookmarks):
    for index, names in enumerate(groups, 1):
        if any(name not in bookmarks or bookmarks[name].get("page", 0) < 1 for name in names):
            raise ValueError(f"第{index}个项目页底信息组页码不完整，不能确认三行同页。")
        if len({bookmarks[name]["page"] for name in names}) != 1:
            raise ValueError(f"第{index}个项目的设备、环境和人员信息仍跨页，未发布新版本；请核对该组内容及页面高度。")


def _measure_excel_signatures(doc):
    from excel_print_documents import PRINT_BODY_CAPTION, PRINT_SIGNATURE_CAPTION

    tables = d._tables(doc)
    bodies = [table for table in tables if _metadata_caption(table).startswith(PRINT_BODY_CAPTION)]
    signatures = [table for table in tables if _metadata_caption(table).startswith(PRINT_SIGNATURE_CAPTION)]
    identities = count(max((int(value) for value in d._xp(doc.element, ".//w:bookmarkStart/@w:id")
                            if value.isdigit()), default=-1) + 1)
    groups = []
    matched = set()
    for index, signature in enumerate(signatures):
        identity = _metadata_caption(signature)[len(PRINT_SIGNATURE_CAPTION):]
        candidates = [table for table in bodies if _metadata_caption(table) == PRINT_BODY_CAPTION + identity]
        if len(candidates) != 1 or identity in matched:
            raise ValueError("Excel试验员页底行与数据页的对应关系不唯一，无法核对分页。")
        matched.add(identity)
        names = []
        for label, table in (("body", candidates[0]), ("signature", signature)):
            paragraphs = d._xp(table, ".//w:p")
            if not paragraphs:
                raise ValueError("Excel记录或试验员行缺少可定位段落，未发布不完整记录。")
            for suffix, paragraph in (("s", paragraphs[0]), ("e", paragraphs[-1])):
                name = f"{ROW_PREFIX}excel_{index}_{label}_{suffix}"
                _bookmark(doc, paragraph, name, end=suffix == "e", identity=next(identities))
                names.append(name)
        groups.append(names)
    if len(matched) != len(bodies):
        raise ValueError("Excel打印记录缺少对应的页底签名行，未发布不完整记录。")
    return groups


def _validate_excel_signature_pages(groups, bookmarks):
    for index, names in enumerate(groups, 1):
        if any(name not in bookmarks or bookmarks[name].get("page", 0) < 1 for name in names):
            raise ValueError(f"第{index}组Excel记录的签名行页码不完整，未发布新版本。")
        if len({bookmarks[name]["page"] for name in names}) != 1:
            raise ValueError(f"第{index}组Excel记录的数据与试验员/横线未保持同页，请核对该打印页高度；未发布新版本。")


def _bookmark(doc, paragraph, name, *, end=False, identity=None):
    if identity is None:
        ids = [int(value) for value in d._xp(doc.element, ".//w:bookmarkStart/@w:id") if value.isdigit()]
        identity = max(ids, default=-1) + 1
    identity = str(identity)
    start, stop = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
    start.set(qn("w:id"), identity)
    start.set(qn("w:name"), name)
    stop.set(qn("w:id"), identity)
    if end:
        paragraph.extend([start, stop])
    else:
        position = 1 if paragraph.find(qn("w:pPr")) is not None else 0
        paragraph.insert(position, start)
        paragraph.insert(position + 1, stop)


def _remove_marks(doc, prefix):
    starts = [node for node in d._xp(doc.element, ".//w:bookmarkStart")
              if node.get(qn("w:name"), "").startswith(prefix)]
    identities = {node.get(qn("w:id")) for node in starts}
    for node in starts + [node for node in d._xp(doc.element, ".//w:bookmarkEnd")
                          if node.get(qn("w:id")) in identities]:
        node.getparent().remove(node)


def _table_name(table):
    names = d._xp(table, ".//w:bookmarkStart/@w:name")
    found = [name for name in names if name.startswith(TABLE_PREFIX)]
    if len(found) > 1:
        raise ValueError("结果续表标记重复，无法安全分页。")
    return found[0] if found else ""


def _row_settings(table, *, keep_sample_groups=True):
    linked_rows = {index for start, stop in _subproject_groups(table)
                   if not any(_overflow_row(row) for row in d._rows(table)[start:stop])
                   for index in range(start, stop - 1)} if keep_sample_groups else set()
    for index, row in enumerate(d._rows(table)):
        properties = row.get_or_add_trPr()
        for node in list(properties.findall(qn("w:tblHeader"))):
            properties.remove(node)
        if index == 0:
            header = OxmlElement("w:tblHeader")
            properties.insert_element_before(header, "w:tblCellSpacing", "w:jc", "w:hidden",
                                             "w:ins", "w:del", "w:trPrChange")
        split = properties.find(qn("w:cantSplit"))
        if split is None:
            split = OxmlElement("w:cantSplit")
            properties.insert_element_before(split, "w:trHeight", "w:tblHeader", "w:tblCellSpacing",
                                             "w:jc", "w:hidden", "w:ins", "w:del", "w:trPrChange")
        split.set(qn("w:val"), "true")
        for paragraph in d._xp(row, ".//w:p"):
            d._paragraph_setting(paragraph, "keepNext", index == 0 or index in linked_rows)


def mark_result_tables(doc, *, keep_sample_groups=True):
    group = uuid.uuid4().hex[:16]
    for table in d._tables(doc):
        rows = d._rows(table)
        if not rows or not d._result_header_row(rows[0]):
            continue
        if not _table_name(table):
            name = TABLE_PREFIX + group + "_" + uuid.uuid4().hex[:8]
            _bookmark(doc, d._xp(rows[0], ".//w:p")[0], name)
        _row_settings(table, keep_sample_groups=keep_sample_groups)


def _grid_cells(row):
    before = d._xp(row, "./w:trPr/w:gridBefore/@w:val")
    column = int(before[0]) if before else 0
    for cell in d._cells(row):
        span = d._xp(cell, "./w:tcPr/w:gridSpan/@w:val")
        width = int(span[0]) if span else 1
        yield column, width, cell
        column += width


def _continued(cell):
    merge = cell.find("./" + qn("w:tcPr") + "/" + qn("w:vMerge"))
    return merge is not None and merge.get(qn("w:val")) != "restart"


def _allowed_columns(table):
    columns = set()
    for column, width, cell in _grid_cells(d._rows(table)[0]):
        if d._norm(d._text(cell)) in {"序号", "项目名称", "技术要求", "要求"}:
            columns.update(range(column, column + width))
    return columns


def _subproject_groups(table):
    rows = d._rows(table)
    headers = list(_grid_cells(rows[0]))
    if not any(d._norm(d._text(cell)) == "样机编号" for _, _, cell in headers):
        return []
    columns = _allowed_columns(table)
    for column, width, cell in headers:
        if d._norm(d._text(cell)) in {"条件", "试验条件", "流量点"}:
            columns.update(range(column, column + width))
    owners, groups, previous, start = {}, [], None, 1
    for index, row in enumerate(rows[1:], 1):
        values = {}
        for column, width, cell in _grid_cells(row):
            covered = set(range(column, column + width))
            if not covered.issubset(columns):
                continue
            if not _continued(cell):
                text = d._norm(d._text(cell))
                merge = cell.find("./" + qn("w:tcPr") + "/" + qn("w:vMerge"))
                # Explicit merge starts distinguish separate tests with identical text.
                value = (column, width, text, index if merge is not None or not text else None)
                for position in covered:
                    owners[position] = value
            for position in covered:
                values[position] = owners.get(position, ("unresolved", index, position))
        key = tuple((column, values.get(column, ("missing", index))) for column in sorted(columns))
        if previous is not None and key != previous:
            groups.append((start, index))
            start = index
        previous = key
    if len(rows) > 1:
        groups.append((start, len(rows)))
    return groups


def shared_requirement_groups(table):
    rows = d._rows(table)
    # Composite headers end with the flow-condition leaf; the preceding grid
    # column holds the shared requirement, not a wider project classification.
    columns = {column + max(0, width - 2) for column, width, cell in _grid_cells(rows[0])
               if d._norm(d._text(cell)) in {"技术要求", "要求"}}
    regions = []
    for index, row in enumerate(rows[1:], 1):
        for column, width, cell in _grid_cells(row):
            if (not set(range(column, column + width)).intersection(columns)
                    or not d._text(cell).strip()
                    or d._xp(cell, "./w:tcPr/w:vMerge/@w:val") != ["restart"]):
                continue
            stop = index + 1
            while stop < len(rows):
                following = [(span, target) for start, span, target in _grid_cells(rows[stop])
                             if start == column]
                if not following or following[0][0] != width or not _continued(following[0][1]):
                    break
                stop += 1
            if stop - index > 1:
                regions.append((index, stop))
    groups = []
    for start, stop in sorted(regions):
        if groups and start < groups[-1][1]:
            groups[-1] = (groups[-1][0], max(groups[-1][1], stop))
        else:
            groups.append((start, stop))
    return groups


def _mask_outside_height_merges(table, groups):
    rows = d._rows(table)
    for index, row in enumerate(rows[1:], 1):
        for column, width, cell in list(_grid_cells(row)):
            if d._xp(cell, "./w:tcPr/w:vMerge/@w:val") != ["restart"]:
                continue
            region = [(index, cell)]
            stop = index + 1
            while stop < len(rows):
                following = [(span, target) for start, span, target in _grid_cells(rows[stop])
                             if start == column]
                if not following or following[0][0] != width or not _continued(following[0][1]):
                    break
                region.append((stop, following[0][1]))
                stop += 1
            if not any(index < end and stop > start and (index < start or stop > end)
                       for start, end in groups):
                continue
            # Only the measurement copy drops merges belonging to a wider row range.
            for _, target in region:
                for node in d._xp(target, "./w:tcPr/w:vMerge"):
                    target.tcPr.remove(node)
                for node in list(target):
                    if node.tag != qn("w:tcPr"):
                        target.remove(node)
                paragraph = OxmlElement("w:p")
                spacing = paragraph.get_or_add_pPr().get_or_add_spacing()
                for key, value in (("before", "0"), ("after", "0"), ("line", "1"), ("lineRule", "exact")):
                    spacing.set(qn("w:" + key), value)
                target.append(paragraph)


def _cell_width_geometry(table):
    grid = [int(value) for value in d._xp(table, "./w:tblGrid/w:gridCol/@w:w")]
    if not grid or min(grid) <= 0:
        raise ValueError("结果表列网格无效，无法核对Word行高测量宽度。")
    edges = [0]
    for width in grid:
        edges.append(edges[-1] + width)
    rows = []
    for row in d._rows(table):
        cells = []
        for column, span, cell in _grid_cells(row):
            if column < 0 or span < 1 or column + span > len(grid):
                raise ValueError("结果单元格超出列网格，无法核对Word行高测量宽度。")
            widths = d._xp(cell, "./w:tcPr/w:tcW/@w:w")
            types = d._xp(cell, "./w:tcPr/w:tcW/@w:type")
            if len(widths) != 1:
                raise ValueError("结果单元格宽度不明确，无法核对Word行高测量宽度。")
            cells.append((edges[column], edges[column + span], types[0] if types else "dxa", int(widths[0])))
        rows.append(tuple(cells))
    return edges[-1], tuple(rows)


def equalize_shared_row_heights(path, *, timeout_seconds, progress=None, renderer=None):
    measure = render if renderer is None else renderer
    if not callable(measure):
        raise ValueError("共享行高测量器不可用，未修改源文档。")
    document = d._load(path)
    tables = d._tables(document)
    references = [(index, start, stop) for index, table in enumerate(tables)
                  if d._rows(table) and d._result_header_row(d._rows(table)[0])
                  for start, stop in shared_requirement_groups(table)]
    if not references:
        return 0
    probe = deepcopy(document)
    probe_tables = d._tables(probe)
    for index in {index for index, _, _ in references}:
        _mask_outside_height_merges(probe_tables[index], shared_requirement_groups(tables[index]))
    groups = prepare_shared_row_heights(probe)
    before = [[d._norm(d._text(cell)) for row in d._rows(table) for cell in d._cells(row)]
              for table in probe_tables]
    with TemporaryDirectory(prefix="shared-row-heights-", dir=Path(path).parent) as temporary:
        measured_path = Path(temporary) / "measurement.docx"
        probe.save(measured_path)
        measure(measured_path, measured_path.with_suffix(".pdf"),
                timeout_seconds=timeout_seconds, progress=progress, _shared_row_height_groups=groups)
        measured = d._tables(d._load(measured_path))
    after = [[d._norm(d._text(cell)) for row in d._rows(table) for cell in d._cells(row)]
             for table in measured]
    if before != after:
        raise ValueError("Word均分行高测量改变了表格内容，停止输出。")
    for index, start, stop in references:
        # Word may coalesce unused grid lines without moving any physical cell edge.
        if _cell_width_geometry(tables[index]) != _cell_width_geometry(measured[index]):
            raise ValueError("Word行高测量改变了原表列宽，不能按不同宽度回写行高。")
        heights = [d._xp(row, "./w:trPr/w:trHeight/@w:val")
                   for row in d._rows(measured[index])[start:stop]]
        if any(len(value) != 1 or not value[0].isdigit() for value in heights):
            raise ValueError("Word未返回完整共享技术要求行高，停止输出。")
        values = [int(value[0]) for value in heights]
        if len(set(values)) != 1 or min(values) < d.PROJECT_METADATA_ROW_HEIGHT_TWIPS:
            raise ValueError("Word共享技术要求行高未均分或低于0.8cm，停止输出。")
        for row in d._rows(tables[index])[start:stop]:
            properties = row.get_or_add_trPr()
            height = properties.get_or_add_trHeight()
            for duplicate in list(properties.findall(qn("w:trHeight")))[1:]:
                properties.remove(duplicate)
            height.set(qn("w:val"), str(values[0]))
            height.set(qn("w:hRule"), "atLeast")
    document.save(path)
    return len(references)


def prepare_shared_row_heights(doc):
    _remove_marks(doc, HEIGHT_PREFIX)
    groups = []
    for table in d._tables(doc):
        rows = d._rows(table)
        if not rows or not d._result_header_row(rows[0]):
            continue
        for start, stop in shared_requirement_groups(table):
            identity = HEIGHT_PREFIX + uuid.uuid4().hex[:16]
            names = (identity + "_s", identity + "_e")
            for index, name in ((start, names[0]), (stop - 1, names[1])):
                cell = d._cells(rows[index])[-1]
                if d._xp(cell, "./w:tcPr/w:vMerge"):
                    raise ValueError("共享技术要求组的实测格有纵向合并，无法安全均分行高。")
                paragraphs = d._xp(cell, "./w:p")
                if not paragraphs:
                    raise ValueError("共享技术要求组缺少可定位的实测格段落，无法均分行高。")
                _bookmark(doc, paragraphs[0], name)
            groups.append((*names, stop - start))
    return groups


def _overflow_row(row):
    return any(name.startswith(OVERFLOW_PREFIX) for name in d._xp(row, ".//w:bookmarkStart/@w:name"))


def _release_oversized_groups(doc, table, pages):
    rows = d._rows(table)
    changed = False
    identities = count(max((int(value) for value in d._xp(doc.element, ".//w:bookmarkStart/@w:id")
                            if value.isdigit()), default=-1) + 1)
    for start, stop in _subproject_groups(table):
        if pages[start] == pages[stop - 1] or any(_overflow_row(row) for row in rows[start:stop]):
            continue
        # Word already tried the intact keep-next chain. If it still spans pages,
        # allow natural row breaks and retain that decision through later splits.
        for row in rows[start:stop]:
            paragraph = d._xp(d._cells(row)[-1], "./w:p")[0]
            _bookmark(doc, paragraph, OVERFLOW_PREFIX + uuid.uuid4().hex[:16], identity=next(identities))
        changed = True
    if changed:
        _row_settings(table)
    return changed


def _copy_cell_text(target, source):
    if d._xp(source, ".//w:bookmarkStart | .//w:commentRangeStart | .//w:commentReference"):
        raise ValueError("待复制的续表说明含书签或批注，无法安全重复，请人工处理。")
    for child in list(target):
        if child.tag != qn("w:tcPr"):
            target.remove(child)
    for child in source:
        if child.tag != qn("w:tcPr"):
            target.append(deepcopy(child))
    if _continued(target):
        target.find("./" + qn("w:tcPr") + "/" + qn("w:vMerge")).set(qn("w:val"), "restart")


def repair_continuation_cells(doc, *, keep_sample_groups=True):
    previous = {}
    warnings = []
    for table in d._tables(doc):
        name = _table_name(table)
        if not name:
            continue
        group = name[len(TABLE_PREFIX):].split("_")[0]
        prior = previous.setdefault(group, {})
        allowed = _allowed_columns(table)
        for index, row in enumerate(d._rows(table)[1:], 1):
            for column, width, cell in _grid_cells(row):
                inherited = prior.get(column)
                if index == 1 and column in allowed and not d._text(cell).strip() and inherited is not None:
                    old_width, owner = inherited
                    if old_width != width or not set(range(column, column + width)).issubset(allowed):
                        raise ValueError("续表列宽合并关系变化，不能跨列复制序号、项目名称或技术要求。")
                    if d._text(owner).strip():
                        _copy_cell_text(cell, owner)
                    elif not _continued(cell):
                        warnings.append("续页首行有空白说明，但前页该列也为空，未补造内容，请核对。")
                if not _continued(cell):
                    prior[column] = (width, cell)
                elif inherited is None:
                    raise ValueError("结果表存在无起点的纵向合并，不能安全定位续页。")
        _row_settings(table, keep_sample_groups=keep_sample_groups)
    return list(dict.fromkeys(warnings))


def _measure_marks(doc):
    _remove_marks(doc, ROW_PREFIX)
    existing_ids = [int(value) for value in d._xp(doc.element, ".//w:bookmarkStart/@w:id") if value.isdigit()]
    identities = count(max(existing_ids, default=-1) + 1)
    marked = []
    for table in d._tables(doc):
        name = _table_name(table)
        if not name:
            continue
        rows = d._rows(table)
        markers = []
        contents = []
        for index, row in enumerate(rows):
            independent = [cell for cell in d._cells(row) if not _continued(cell)]
            if not independent:
                raise ValueError("结果行没有可独立定位的单元格，无法确定实际续页边界。")
            paragraphs = d._xp(independent[-1], "./w:p")
            if not paragraphs:
                raise ValueError("结果行缺少原生段落，无法定位实际页码。")
            base = ROW_PREFIX + name.rsplit("_", 1)[1] + "_" + str(index)
            _bookmark(doc, paragraphs[0], base + "_s", identity=next(identities))
            _bookmark(doc, paragraphs[-1], base + "_e", end=True, identity=next(identities))
            markers.append((base + "_s", base + "_e"))
            for column, cell in enumerate(independent[:-1]):
                paragraphs = d._xp(cell, ".//w:p")
                if not paragraphs:
                    raise ValueError("结果单元格缺少原生段落，无法检查文字是否完整分页。")
                cell_base = base + "_c" + str(column)
                _bookmark(doc, paragraphs[0], cell_base + "_s", identity=next(identities))
                _bookmark(doc, paragraphs[-1], cell_base + "_e", end=True, identity=next(identities))
                contents.append((cell_base + "_s", cell_base + "_e"))
        marked.append((name, markers, contents))
    return marked


def _safe_boundaries(table, starts):
    allowed = _allowed_columns(table)
    rows = d._rows(table)
    owners, state = [], {}
    for index, row in enumerate(rows):
        current = {}
        for column, _width, cell in _grid_cells(row):
            if not _continued(cell):
                state[column] = index
            current[column] = state.get(column, index)
        owners.append(current)
    boundaries = []
    for start in starts:
        adjusted = start
        while True:
            earlier = [owner for column, owner in owners[adjusted].items()
                       if column not in allowed and owner < adjusted]
            if not earlier:
                break
            adjusted = min(earlier)
        if adjusted <= (boundaries[-1] if boundaries else 1):
            raise ValueError("样机编号或实测结果的合并组本身跨页，无法在不拆散数据的情况下补表头，请调整该组。")
        boundaries.append(adjusted)
    return boundaries


def split_measured_tables(doc, measured, bookmarks, *, keep_sample_groups=True):
    _remove_marks(doc, ROW_PREFIX)
    by_name = {_table_name(table): table for table in d._tables(doc) if _table_name(table)}
    changed = False
    for name, markers, contents in measured:
        table = by_name[name]
        pages = []
        for start, end in markers:
            if start not in bookmarks or end not in bookmarks:
                raise ValueError("Word未返回完整结果行页码，不能猜测分页。")
            first, last = bookmarks[start]["page"], bookmarks[end]["page"]
            if first != last:
                raise ValueError("单个结果行仍跨越多页，无法安全插入续表表头，请调整该行内容。")
            pages.append(first)
        if len(pages) < 2:
            continue
        if pages != sorted(pages):
            raise ValueError("Word返回的结果行页码顺序异常，已停止自动分页。")
        if keep_sample_groups and _release_oversized_groups(doc, table, pages):
            changed = True
            continue
        if pages[0] != pages[1]:
            raise ValueError("结果表表头与首行数据未能保持同页，请检查首行高度。")
        starts = [i for i in range(2, len(pages)) if pages[i] != pages[i - 1]]
        if not starts:
            for start, end in contents:
                if start not in bookmarks or end not in bookmarks:
                    raise ValueError("Word未返回完整说明单元格页码，无法确认文字完整性。")
                if bookmarks[start]["page"] != pages[0] or bookmarks[end]["page"] != pages[0]:
                    raise ValueError("技术要求或其他单元格文字仍跨页，请调整该单元格；未截断内容或发布报告。")
            continue
        if not keep_sample_groups:
            # Repeated context can change row heights. Remeasure before fixing
            # the remaining boundaries instead of stranding rows on extra pages.
            starts = [starts[(len(starts) - 1) // 2]]
        starts = _safe_boundaries(table, starts)
        rows = d._rows(table)
        for index, (start, stop) in enumerate(zip([1, *starts], [*starts, len(rows)])):
            chunk = deepcopy(table)
            for row in d._rows(chunk):
                chunk.remove(row)
            chunk.append(deepcopy(rows[0]))
            for row in rows[start:stop]:
                chunk.append(deepcopy(row))
            table.addprevious(chunk)
            header = next(node for node in d._xp(chunk, ".//w:bookmarkStart")
                          if node.get(qn("w:name")) == name)
            header.set(qn("w:name"), name.rsplit("_", 1)[0] + "_" + uuid.uuid4().hex[:8])
            # Composer will not run again after native pagination.
            old_id = header.get(qn("w:id"))
            ids = [int(value) for value in d._xp(doc.element, ".//w:bookmarkStart/@w:id") if value.isdigit()]
            new_id = str(max(ids, default=-1) + 1)
            header.set(qn("w:id"), new_id)
            for end in d._xp(chunk, ".//w:bookmarkEnd"):
                if end.get(qn("w:id")) == old_id:
                    end.set(qn("w:id"), new_id)
            if index:
                d._start_continuation_table(chunk)
                separator = chunk.getprevious()
                if d._xp(separator, "./w:pPr/w:pageBreakBefore[not(@w:val) or "
                                    "@w:val='true' or @w:val='1' or @w:val='on']"):
                    d._explicit_page_start(separator)
            _row_settings(chunk, keep_sample_groups=keep_sample_groups)
        table.getparent().remove(table)
        changed = True
    return changed


def _dismiss_prompt(stop, errors):
    if sys.platform == "win32":
        return  # Word以COM方式打开且DisplayAlerts=0，不会弹出域更新提示
    script = '''
tell application "System Events" to tell process "Microsoft Word"
    repeat with w in windows
        if exists static text "This document contains fields that may refer to other files. Do you want to update the fields in this document?" of w then
            click button "No" of w
            return "handled"
        end if
    end repeat
end tell
return "waiting"
'''
    while not stop.wait(0.5):
        try:
            result = subprocess.run(["osascript", "-e", script], capture_output=True, text=True, timeout=5)
        except (OSError, subprocess.TimeoutExpired) as exc:
            errors.append("Word域更新提示处理失败：" + str(exc))
            return
        if result.returncode:
            errors.append("Word域更新提示处理失败：" + result.stderr.strip())
            return


def finalize_result_pages(path, progress=None):
    from native_word import _LOCK

    if progress:
        progress("等待本机Word队列")
    with _LOCK:
        return _finalize_result_pages(path, progress)


def _finalize_result_pages(path, progress=None):
    deadline = time.monotonic() + 300
    warnings = []
    with TemporaryDirectory(prefix="result-pagination-", dir=Path(path).parent) as temporary:
        working = Path(temporary) / "working.docx"
        pdf = working.with_suffix(".pdf")
        shutil.copyfile(path, working)
        for iteration in range(20):
            if progress:
                progress(f"第{iteration + 1}轮：定位续表、页底信息及签名")
            if time.monotonic() >= deadline:
                raise ValueError("结果续表自动分页超过5分钟，未发布未完成的报告，请核对Word状态后重试。")
            doc = d._load(working)
            _remove_marks(doc, ROW_PREFIX)
            warnings.extend(repair_continuation_cells(doc))
            measured = _measure_marks(doc)
            metadata = _measure_metadata(doc)
            signatures = _measure_excel_signatures(doc)
            names = [marker for _name, pairs, contents in measured
                     for pair in [*pairs, *contents] for marker in pair]
            names.extend(name for group in metadata for name in group)
            names.extend(name for group in signatures for name in group)
            doc.save(working)
            d._normalize_report_fonts(working)
            stop, errors = threading.Event(), []
            watcher = threading.Thread(target=_dismiss_prompt, args=(stop, errors), daemon=True)
            watcher.start()
            try:
                result = render(working, pdf, timeout_seconds=min(180, deadline - time.monotonic()),
                                bookmark_names=names)
            finally:
                stop.set()
                watcher.join(timeout=6)
            if errors:
                raise ValueError("；".join(errors))
            doc = d._load(working)
            if not split_measured_tables(doc, measured, result.get("bookmarks", {})):
                _validate_metadata_pages(metadata, result.get("bookmarks", {}))
                _validate_excel_signature_pages(signatures, result.get("bookmarks", {}))
                _remove_marks(doc, ROW_PREFIX)
                _remove_marks(doc, TABLE_PREFIX)
                _remove_marks(doc, OVERFLOW_PREFIX)
                doc.save(working)
                d._normalize_report_fonts(working)
                shutil.copyfile(working, path)
                return {"pages": result["pages"], "passes": iteration + 1,
                        "metadata_groups_checked": len(metadata),
                        "excel_signature_groups_checked": len(signatures),
                        "warnings": list(dict.fromkeys(warnings))}
            doc.save(working)
    raise ValueError("结果续表分页未收敛，已停止输出，不按估算行数生成报告。")
