"""Editable observation continuations based exclusively on physical Word pages."""

from copy import deepcopy
from itertools import count
import re
import uuid

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import documents as d
import result_pagination as rp


PREFIX = "op_measure_"
FLOW_PREFIX = "op_flow_"


def _columns(row):
    position = 0
    for cell in d._cells(row):
        span = int(next(iter(d._xp(cell, "./w:tcPr/w:gridSpan/@w:val")), "1"))
        yield position, position + span, cell
        position += span


def _layout(table):
    headers = list(_columns(d._rows(table)[0]))
    texts = [d._norm(d._text(cell)) for _, _, cell in headers]
    if (len(headers) not in (5, 6) or texts[0] != "型式评价大纲章节号及项目名称"
            or texts[1] != "要求" or texts[-1] != "备注"):
        raise ValueError("观察项目表的名称、要求或备注列无法唯一定位，未猜测续页。")
    return headers[0][1], headers[-1][0]


def _mark_columns(table):
    headers = list(_columns(d._rows(table)[0]))
    positive, negative = headers[-3:-1]
    if (d._text(positive[2]).strip() not in {"+", "＋"}
            or d._text(negative[2]).strip() not in {"-", "－", "−"}):
        raise ValueError("观察表的＋／－标记列无法唯一定位，不能猜填跨页标记。")
    return {(start, stop) for start, stop, _cell in (positive, negative)}


def _project_rows(tables):
    names = {}
    scope = None
    previous = None
    group = 0
    for table in tables:
        current_scope = tuple(d._xp(table, "./w:tblPr/w:tblCaption/@w:val")) or (id(table),)
        if current_scope != scope:
            names = {}
            scope = current_scope
        name_end, _ = _layout(table)
        headers = list(_columns(d._rows(table)[0]))
        sample_column = headers[2][:2] if len(headers) == 6 else None
        for row in d._rows(table)[1:]:
            for start, stop, cell in _columns(row):
                if start >= name_end:
                    continue
                text = d._norm(d._text(cell))
                if text:
                    if names.get((start, stop)) != text:
                        names = {key: value for key, value in names.items() if key[1] <= start}
                    names[(start, stop)] = text
            key = tuple(sorted((start, stop, text) for (start, stop), text in names.items()
                               if stop <= name_end))
            if not key:
                raise ValueError("观察项目没有明确名称，不能对应跨页标记。")
            sample = tuple(d._text(cell) for start, stop, cell in _columns(row)
                           if (start, stop) == sample_column)
            identity = (scope, key, sample)
            if identity != previous:
                group += 1
                previous = identity
            yield table, row, (group, identity)


def _is_mark(cell):
    text = "".join(d._text(cell).split())
    return bool(text) and all(character in {"×", "x", "X", "✕", "✖", "☓"}
                              for character in text)


def _repeat_mark(cell):
    return not d._text(cell).strip() or _is_mark(cell)


def _center_mark(cell):
    cell.get_or_add_tcPr().get_or_add_vAlign().set(qn("w:val"), "center")
    for paragraph in d._xp(cell, "./w:p"):
        properties = paragraph.get_or_add_pPr()
        properties.get_or_add_jc().set(qn("w:val"), "center")
        spacing = properties.get_or_add_spacing()
        for name in ("before", "after", "beforeLines", "afterLines",
                     "beforeAutospacing", "afterAutospacing"):
            spacing.set(qn("w:" + name), "0")
        indent = properties.get_or_add_ind()
        for name in ("left", "right", "firstLine", "hanging", "leftChars", "rightChars",
                     "firstLineChars", "hangingChars"):
            indent.set(qn("w:" + name), "0")


def _prepare_marks(tables):
    rows = list(_project_rows(tables))
    sources = {}
    marked_columns = {}
    for table, row, key in rows:
        columns = _mark_columns(table)
        present = {(start, stop) for start, stop, cell in _columns(row)
                   if (start, stop) in columns and _is_mark(cell)}
        if present:
            if key in marked_columns and marked_columns[key] != present:
                raise ValueError("同一观察项目的＋／－叉号列不一致，不能猜选续页标记。")
            marked_columns[key] = present
        for start, stop, cell in _columns(row):
            if (start, stop) not in columns:
                continue
            text = d._text(cell).strip()
            if not text:
                continue
            if not _is_mark(cell):
                continue
            identity = (key, start, stop)
            previous = sources.get(identity)
            if previous is not None and d._text(previous).strip().lower() != text.lower():
                raise ValueError("同一观察项目的叉号标记不一致，不能猜选续页标记。")
            sources[identity] = cell
    for table, row, key in rows:
        for start, stop, cell in _columns(row):
            if (start, stop) not in _mark_columns(table):
                continue
            source = sources.get((key, start, stop))
            if source is not None and not d._text(cell).strip():
                _copy_name(cell, source)
            if _is_mark(cell):
                _center_mark(cell)


def report_tables(doc):
    """Select only the composed observe section, never appended uploaded records."""
    blocks = list(doc.element.body)
    starts = [i for i, block in enumerate(blocks)
              if block.tag == qn("w:p") and d._norm(d._text(block)) == d._norm("二、观察项目记录")]
    if not starts:
        return []
    start = starts[0]
    stops = [i for i in range(start + 1, len(blocks))
             if blocks[i].tag == qn("w:p")
             and d._norm(d._text(blocks[i])) == d._norm("三、试验项目记录")]
    if not stops:
        raise ValueError("观察项目记录缺少后续章节边界，不能安全处理续页。")
    tables = [block for block in blocks[start + 1:stops[0]]
              if block.tag == qn("w:tbl") and d._rows(block)
              and "型式评价大纲章节号及项目名称" in d._text(d._rows(block)[0])]
    if not tables:
        raise ValueError("观察项目记录中没有可定位的项目表，不能确认续页。")
    return tables


def prepare(doc):
    tables = report_tables(doc)
    if not tables:
        return []
    for table in tables:
        _layout(table)
        if d._xp(table, ".//w:tbl | .//w:fldChar | .//w:fldSimple | .//w:instrText | "
                       ".//w:sdt | .//w:del | .//w:ins | .//w:altChunk | "
                       ".//*[local-name()='oMath' or local-name()='oMathPara']"):
            raise ValueError("观察表含嵌套表、域、内容控件、修订或公式，无法安全映射字符分页；请先核对素材结构。")
        if table.tblPr.find(qn("w:tblCaption")) is None:
            caption = OxmlElement("w:tblCaption")
            caption.set(qn("w:val"), FLOW_PREFIX + uuid.uuid4().hex[:12])
            table.tblPr.append(caption)
    _prepare_marks(tables)
    d._modern_word_compatibility(doc)
    return tables


def _units(text):
    return len(text.encode("utf-16-le")) // 2


def _leaf_text(node):
    if node.tag == qn("w:t"):
        return node.text or ""
    values = {"tab": "\t", "br": "\v", "cr": "\v", "noBreakHyphen": "\u2011",
              "softHyphen": "\u00ad", "drawing": "\x01", "pict": "\x01", "object": "\x01"}
    for name, value in values.items():
        if node.tag == qn("w:" + name):
            if name == "br" and node.get(qn("w:type"), "textWrapping") != "textWrapping":
                raise ValueError("观察单元格含显式分页或分栏符，无法按自然续页安全拆分。")
            return value
    return None


def _text(node):
    leaf = _leaf_text(node)
    if leaf is not None:
        return leaf
    if node.tag in {qn("w:pPr"), qn("w:rPr"), qn("w:tcPr")}:
        return ""
    if node.tag == qn("w:sym"):
        raise ValueError("观察单元格含特殊字体符号，无法确认Word字符映射，请核对素材。")
    return "".join(_text(child) for child in node)


def measure(doc, tables):
    rp._remove_marks(doc, PREFIX)
    identities = count(max((int(value) for value in d._xp(doc.element, ".//w:bookmarkStart/@w:id")
                            if value.isdigit()), default=-1) + 1)
    measured = []
    for table_index, table in enumerate(tables):
        name_end, _remarks = _layout(table)
        rows = []
        for row_index, row in enumerate(d._rows(table)[1:], 1):
            cells = []
            for cell_index, (start, stop, cell) in enumerate(_columns(row)):
                if start < name_end < stop:
                    raise ValueError("观察名称列与要求列交叉合并，不能安全补全续页名称。")
                paragraphs = d._xp(cell, "./w:p")
                if not paragraphs:
                    raise ValueError("观察单元格缺少可编辑段落，不能确认内容完整。")
                if start < name_end and not d._text(cell).strip():
                    cells.append([])
                    continue
                marks = []
                for paragraph_index, paragraph in enumerate(paragraphs):
                    text = _text(paragraph)
                    base = f"{PREFIX}{table_index}_{row_index}_{cell_index}_{paragraph_index}"
                    rp._bookmark(doc, paragraph, base + "_s", identity=next(identities))
                    rp._bookmark(doc, paragraph, base + "_e", end=True, identity=next(identities))
                    marks.append((base, text))
                cells.append(marks)
            rows.append(cells)
        measured.append(rows)
    return measured


def queries(measured):
    return [mark for rows in measured for cells in rows for marks in cells for mark in marks]


def marked_queries(doc):
    """Read an already marked source without regenerating its measurement identities."""
    result, seen = [], set()
    for paragraph in d._xp(doc.element, ".//w:p"):
        marks = [node for node in d._xp(paragraph, ".//w:bookmarkStart")
                 if node.get(qn("w:name"), "").startswith(PREFIX)]
        if not marks:
            continue
        names = [node.get(qn("w:name")) for node in marks]
        if (len(marks) != 2 or not names[0].endswith("_s")
                or names[1] != names[0][:-2] + "_e"
                or any(node.getparent() is not paragraph for node in marks)):
            raise ValueError("观察测量书签不是同一段落内的唯一首尾对，不能提交原生任务。")
        base = names[0][:-2]
        children = list(paragraph)
        start, end = (children.index(node) for node in marks)
        if (base in seen or _text(paragraph) !=
                "".join(_text(node) for node in children[start + 1:end])):
            raise ValueError("观察测量书签重复或未覆盖完整源段落，不能提交原生任务。")
        seen.add(base)
        result.append((base, _text(paragraph)))
    return result


def _slice(node, first, last, *, final=False):
    """Partition rich XML at UTF-16 boundaries; zero-width anchors occur only once."""
    position = 0

    def visit(source):
        nonlocal position
        leaf = _leaf_text(source)
        if leaf is not None:
            size = _units(leaf)
            start, stop = position, position + size
            position = stop
            if max(start, first) >= min(stop, last):
                return None
            result = deepcopy(source)
            if source.tag == qn("w:t"):
                raw = leaf.encode("utf-16-le")
                try:
                    result.text = raw[max(0, first - start) * 2:min(size, last - start) * 2].decode("utf-16-le")
                except UnicodeDecodeError as exc:
                    raise ValueError("Word观察分页边界落在字符内部，未截断内容。") from exc
                result.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            elif first > start or last < stop:
                raise ValueError("Word观察分页边界落在图片或特殊字符内部，未截断内容。")
            return result
        if source.tag in {qn("w:pPr"), qn("w:rPr"), qn("w:tcPr")}:
            return deepcopy(source)
        if source.tag == qn("w:lastRenderedPageBreak"):
            return None
        if not len(source):
            return deepcopy(source) if first <= position < last or final and position == last else None
        result = deepcopy(source)
        for child in list(result):
            result.remove(child)
        for child in source:
            copied = visit(child)
            if copied is not None:
                result.append(copied)
        meaningful = [child for child in result if child.tag not in {qn("w:pPr"), qn("w:rPr")}]
        return result if meaningful or source.tag == qn("w:p") else None

    result = visit(node)
    if result is None:
        result = OxmlElement("w:p")
    if first and result.tag == qn("w:p"):
        for numbering in d._xp(result, "./w:pPr/w:numPr"):
            numbering.getparent().remove(numbering)
    return result


def _blank(paragraph):
    return (not _text(paragraph).strip()
            and not d._xp(paragraph, ".//w:drawing | .//w:pict | .//w:object | "
                          ".//w:bookmarkStart | .//w:bookmarkEnd | .//w:commentRangeStart | "
                          ".//w:commentRangeEnd | ./w:pPr/w:numPr"))


def _trim_remarks(cell):
    paragraphs = d._xp(cell, "./w:p")
    while len(paragraphs) > 1 and _blank(paragraphs[0]):
        cell.remove(paragraphs.pop(0))
    if paragraphs:
        paragraph = paragraphs[0]
        if d._xp(paragraph, ".//w:bookmarkStart | .//w:bookmarkEnd"):
            return
        match = re.match(r"^(?:[ \t]*[\v\r\n])+", _text(paragraph))
        if match:
            text = _text(paragraph)
            cell.replace(paragraph, _slice(paragraph, _units(match.group()), _units(text), final=True))


def _set_content(cell, paragraphs):
    for child in list(cell):
        if child.tag != qn("w:tcPr"):
            cell.remove(child)
    cell.extend(paragraphs or [OxmlElement("w:p")])


def _copy_name(target, source):
    paragraphs = [deepcopy(paragraph) for paragraph in d._xp(source, "./w:p")]
    for paragraph in paragraphs:
        for mark in d._xp(paragraph, ".//w:bookmarkStart | .//w:bookmarkEnd"):
            mark.getparent().remove(mark)
    _set_content(target, paragraphs)


def _protect_name(cell):
    paragraphs = d._xp(cell, "./w:p")
    paragraph = paragraphs[0]
    for following in paragraphs[1:]:
        line = OxmlElement("w:r")
        line.append(OxmlElement("w:br"))
        paragraph.append(line)
        for child in list(following):
            if child.tag != qn("w:pPr"):
                paragraph.append(child)
        cell.remove(following)
    d._paragraph_setting(paragraph, "keepLines", True)
    d._paragraph_setting(paragraph, "keepNext", False)


def finish_name_layout(tables):
    """Merge shared parents only inside already measured single-page continuations."""
    before = [table.xml for table in tables]
    for table in tables:
        name_end, _ = _layout(table)
        chains, names = {}, {}
        for row in d._rows(table)[1:]:
            present = set()
            for start, stop, cell in _columns(row):
                if start >= name_end:
                    continue
                span = (start, stop)
                text = d._norm(d._text(cell))
                if text:
                    if names.get(span) != text:
                        names = {key: value for key, value in names.items() if key[1] <= start}
                    names[span] = text
                elif not d._xp(cell, "./w:tcPr/w:vMerge"):
                    names.pop(span, None)
                if not names.get(span):
                    continue
                identity = tuple(sorted((a, b, value) for (a, b), value in names.items()
                                        if b <= stop))
                present.add(span)
                previous = chains.get(span)
                if stop < name_end and previous is not None and previous[0] == identity:
                    first = previous[1]
                    merge = first.get_or_add_tcPr().find(qn("w:vMerge"))
                    if merge is None:
                        merge = OxmlElement("w:vMerge")
                        first.tcPr.append(merge)
                    merge.set(qn("w:val"), "restart")
                    merge = cell.get_or_add_tcPr().find(qn("w:vMerge"))
                    if merge is None:
                        merge = OxmlElement("w:vMerge")
                        cell.tcPr.append(merge)
                    merge.set(qn("w:val"), "continue")
                    _set_content(cell, [])
                else:
                    chains[span] = (identity, cell)
                if d._text(cell).strip():
                    _protect_name(cell)
                _center_mark(cell)
            chains = {key: value for key, value in chains.items() if key in present}
    return any(table.xml != original for table, original in zip(tables, before))


def validate_name_layout(tables, measured, locations):
    for table, rows in zip(tables, measured):
        pages = set()
        for row in rows:
            for paragraphs in row:
                for base, _ in paragraphs:
                    if base not in locations:
                        raise ValueError("观察名称合并后的Word页码缺失，未发布未完成版本。")
                    pages.update(page for page, _ in locations[base])
        if len(pages) != 1:
            raise ValueError("观察名称居中合并后表格仍跨页，请核对名称长度或页面设置；未发布未完成版本。")


def _remove_measurements(element):
    starts = [node for node in d._xp(element, ".//w:bookmarkStart")
              if node.get(qn("w:name"), "").startswith(PREFIX)]
    identities = {node.get(qn("w:id")) for node in starts}
    for node in starts + [node for node in d._xp(element, ".//w:bookmarkEnd")
                          if node.get(qn("w:id")) in identities]:
        node.getparent().remove(node)


def cleanup(doc):
    rp._remove_marks(doc, PREFIX)
    for caption in d._xp(doc.element, ".//w:tblPr/w:tblCaption"):
        if caption.get(qn("w:val"), "").startswith(FLOW_PREFIX):
            caption.getparent().remove(caption)


def _join_flow(tables):
    """Release only our generated page boundaries after a name changes row flow."""
    for table in tables:
        if table.getparent() is None:
            continue
        identity = rp._metadata_caption(table)
        if not identity.startswith(FLOW_PREFIX):
            continue
        while True:
            separator = table.getnext()
            following = separator.getnext() if separator is not None else None
            if (separator is None or separator.tag != qn("w:p")
                    or following is None or following.tag != qn("w:tbl")
                    or rp._metadata_caption(following) != identity):
                break
            if (d._text(separator).strip()
                    or d._xp(separator, ".//w:sectPr | .//w:bookmarkStart | "
                                       ".//w:drawing | .//w:fldChar | .//w:fldSimple")):
                raise ValueError("观察自动续表间出现有效内容，不能安全连续重排。")
            for row in d._rows(following)[1:]:
                table.append(row)
            following.getparent().remove(following)
            separator.getparent().remove(separator)


def _move_project_start(tables, row):
    for table in tables:
        _remove_measurements(table)
    _join_flow(tables)
    table = row.getparent()
    rows = d._rows(table)
    index = rows.index(row)
    origins = {}
    name_end, _ = _layout(table)
    for preceding in rows[1:index + 1]:
        for start, stop, cell in _columns(preceding):
            if start < name_end and d._text(cell).strip():
                origins[(start, stop)] = cell
    for start, stop, cell in _columns(row):
        if start < name_end:
            if not d._text(cell).strip() and (start, stop) in origins:
                _copy_name(cell, origins[(start, stop)])
            for merge in d._xp(cell, "./w:tcPr/w:vMerge"):
                merge.set(qn("w:val"), "restart")
    if index == 1:
        previous = table.getprevious()
        if previous is not None and d._xp(previous, ".//w:br[@w:type='page']"):
            raise ValueError("观察项目在另起页后仍无法完整容纳名称，请核对名称长度或页面设置；未删字或缩字。")
        continuation = table
    else:
        continuation = deepcopy(table)
        for old in d._rows(continuation)[1:]:
            continuation.remove(old)
        for following in rows[index:]:
            continuation.append(following)
        table.addnext(continuation)
    d._start_continuation_table(continuation)
    d._explicit_page_start(continuation.getprevious())


def _repair_name_layout(tables, measured, locations):
    changed = False
    for table, row_marks in zip(tables, measured):
        name_end, _ = _layout(table)
        mark_columns = _mark_columns(table)
        rows = d._rows(table)[1:]
        if len(rows) != len(row_marks):
            raise ValueError("Word改变了观察表行结构，不能安全确认名称完整。")
        for row_index, (row, cells_marks) in enumerate(zip(rows, row_marks)):
            columns = list(_columns(row))
            data_pages = {page for (start, stop, cell), marks in zip(columns, cells_marks)
                          if start >= name_end and not ((start, stop) in mark_columns
                                                       and _repeat_mark(cell)) for base, _text in marks
                          for page, _position in locations.get(base, [])}
            for (start, stop, cell), marks in zip(columns, cells_marks):
                if start >= name_end or not marks:
                    continue
                name_pages = set()
                for base, _text_value in marks:
                    if base not in locations:
                        raise ValueError("Word未返回完整观察名称页码，不能确认名称未被截断。")
                    name_pages.update(page for page, _offset in locations[base])
                if not name_pages or not data_pages:
                    raise ValueError("Word观察名称或项目内容页码缺失，不能猜测另起页。")
                if len(name_pages) == 1 and min(name_pages) == min(data_pages):
                    continue
                already_bound = all(d._xp(paragraph, "./w:pPr/w:keepLines[@w:val='true']")
                                    for paragraph in d._xp(cell, "./w:p"))
                if already_bound and not d._xp(cell, "./w:tcPr/w:vMerge"):
                    if len(name_pages) == 1 and min(name_pages) > min(data_pages):
                        _move_project_start(tables, row)
                        return True
                    raise ValueError("观察项目名称在名称同页约束下仍不能完整显示，请核对名称长度或页面设置；未删字或缩字。")
                # Release the merged name without binding the long requirements row.
                for next_row in rows[row_index:]:
                    candidates = [c for a, b, c in _columns(next_row) if (a, b) == (start, stop)]
                    if not candidates:
                        break
                    merge = candidates[0].find("./" + qn("w:tcPr") + "/" + qn("w:vMerge"))
                    if next_row is not row and (merge is None or merge.get(qn("w:val")) == "restart"):
                        break
                    if merge is not None:
                        merge.getparent().remove(merge)
                    target = candidates[0]
                    if next_row is not row and not d._text(target).strip():
                        _copy_name(target, cell)
                    target.get_or_add_tcPr().get_or_add_vAlign().set(qn("w:val"), "top")
                    _protect_name(target)
                changed = True
    if changed:
        for table in tables:
            _remove_measurements(table)
        _join_flow(tables)
    return changed


def _checked_segments(base, text, locations):
    size = _units(text)
    segments = locations.get(base)
    if (not segments or segments[0][1] != 0
            or any(page < 1 or offset < 0 or offset > size for page, offset in segments)
            or any(a[0] >= b[0] or a[1] >= b[1]
                   for a, b in zip(segments, segments[1:]))):
        raise ValueError("Word观察字符分页定位不完整或顺序异常。")
    return segments


def _carry_nonfinal_remarks(doc, tables, measured, locations):
    row_marks = {row: marks for table, marks_by_row in zip(tables, measured)
                 for row, marks in zip(d._rows(table)[1:], marks_by_row)}
    projects = list(_project_rows(tables))

    def paragraphs(row, column):
        columns = list(_columns(row))
        marks_by_cell = row_marks[row]
        if len(columns) != len(marks_by_cell):
            raise ValueError("Word改变了观察表单元格结构，不能安全移动备注。")
        matches = [(cell, marks) for (start, stop, cell), marks in zip(columns, marks_by_cell)
                   if (start, stop) == column]
        if len(matches) != 1:
            raise ValueError("观察要求或备注单元格无法唯一对应，不能安全移动备注。")
        cell, marks = matches[0]
        nodes = d._xp(cell, "./w:p")
        if len(nodes) != len(marks):
            raise ValueError("Word改变了观察段落结构，不能安全移动备注。")
        entries = []
        for paragraph, (base, text) in zip(nodes, marks):
            if _text(paragraph) != text:
                raise ValueError("Word观察文字与测量前不一致，未按旧位置移动备注。")
            entries.append((paragraph, text, _checked_segments(base, text, locations)))
        return cell, entries

    for (table, row, key), (following_table, following_row, following_key) in zip(projects, projects[1:]):
        if key != following_key:
            continue
        headers = list(_columns(d._rows(table)[0]))
        next_headers = list(_columns(d._rows(following_table)[0]))
        requirement, remark = headers[1][:2], headers[-1][:2]
        if requirement != next_headers[1][:2] or remark != next_headers[-1][:2]:
            continue
        _, requirements = paragraphs(row, requirement)
        _, next_requirements = paragraphs(following_row, requirement)
        requirement_pages = [page for _, text, segments in requirements if text.strip()
                             for page, _ in segments]
        next_pages = [page for _, text, segments in next_requirements if text.strip()
                      for page, _ in segments]
        if not requirement_pages or not next_pages:
            continue
        last_page, next_page = max(requirement_pages), min(next_pages)
        if next_page <= last_page:
            continue
        cell, notes = paragraphs(row, remark)
        for index, (_, text, segments) in enumerate(notes):
            overflow = [page for page, _ in segments if page > last_page]
            if not text.strip() or not overflow or min(overflow) != next_page:
                continue
            target, _ = paragraphs(following_row, remark)
            first = d._xp(target, "./w:p")[0]
            # Move the entire suffix, retaining note order, runs and paragraph formatting.
            for paragraph, _, _ in notes[index:]:
                first.addprevious(paragraph)
                d._paragraph_setting(paragraph, "keepLines", True)
            if not d._xp(cell, "./w:p"):
                cell.append(OxmlElement("w:p"))
            rp._remove_marks(doc, PREFIX)
            return True
    return False


def split(doc, tables, measured, locations, *, carry_nonfinal_remarks=False):
    if len(tables) != len(measured):
        raise ValueError("Word改变了观察表数量，不能安全连续重排。")
    if carry_nonfinal_remarks:
        for base, text in queries(measured):
            _checked_segments(base, text, locations)
    if _repair_name_layout(tables, measured, locations):
        return True
    if carry_nonfinal_remarks and _carry_nonfinal_remarks(doc, tables, measured, locations):
        return True
    changed = False
    for table, row_marks in zip(tables, measured):
        rows = d._rows(table)
        name_end, remarks_start = _layout(table)
        mark_columns = _mark_columns(table)
        names = {}
        pieces = []
        for row, cells_marks in zip(rows[1:], row_marks):
            columns = list(_columns(row))
            if len(columns) != len(cells_marks):
                raise ValueError("Word改变了观察表单元格结构，不能安全续页。")
            for start, stop, cell in columns:
                if start < name_end and d._text(cell).strip():
                    names[(start, stop)] = cell
            by_cell = []
            row_pages = set()
            for (start, stop, cell), marks in zip(columns, cells_marks):
                contents = {}
                if marks and start >= name_end and not ((start, stop) in mark_columns and _repeat_mark(cell)):
                    paragraphs = d._xp(cell, "./w:p")
                    if len(paragraphs) != len(marks):
                        raise ValueError("Word改变了观察段落结构，不能安全续页。")
                    for paragraph, (base, expected) in zip(paragraphs, marks):
                        if _text(paragraph) != expected:
                            raise ValueError("Word观察文字与测量前不一致，未按旧位置截断。")
                        size = _units(expected)
                        segments = _checked_segments(base, expected, locations)
                        for index, (page, first) in enumerate(segments):
                            last = segments[index + 1][1] if index + 1 < len(segments) else size
                            contents.setdefault(page, []).append(_slice(paragraph, first, last,
                                                                        final=last == size))
                            row_pages.add(page)
                by_cell.append(contents)
            if not row_pages:
                raise ValueError("观察行没有可测量的要求或备注，不能猜测行页码。")
            pages = sorted(row_pages)
            for index, page in enumerate(pages):
                clone = deepcopy(row)
                for (start, stop, source), target, content in zip(columns, d._cells(clone), by_cell):
                    if (start, stop) in mark_columns and _repeat_mark(source):
                        _copy_name(target, source)
                        if d._text(target).strip():
                            _center_mark(target)
                    elif start >= name_end:
                        _set_content(target, content.get(page, []))
                    elif index:
                        origin = names.get((start, stop))
                        if origin is not None:
                            _copy_name(target, origin)
                if len(pages) > 1:
                    for height in d._xp(clone, "./w:trPr/w:trHeight"):
                        height.getparent().remove(height)
                    for constraint in d._xp(clone, "./w:trPr/w:cantSplit"):
                        constraint.getparent().remove(constraint)
                pieces.append((page, clone, dict(names), index > 0))
        page_order = [page for page, _row, _names, _continued in pieces]
        if page_order != sorted(page_order):
            raise ValueError("观察行的实际页码相互交叉，无法安全重建续表；未删除或重复要求/备注。")
        page_changes = any(a != b for a, b in zip(page_order, page_order[1:]))
        if not page_changes:
            continue
        chunks = []
        active_page = None
        for page, row, origins, continued in pieces:
            new_page = page != active_page
            if new_page:
                chunk = deepcopy(table)
                for old in d._rows(chunk):
                    chunk.remove(old)
                header = deepcopy(rows[0])
                for mark in d._xp(header, ".//w:bookmarkStart | .//w:bookmarkEnd"):
                    mark.getparent().remove(mark)
                chunk.append(header)
                chunks.append(chunk)
                active_page = page
            _remove_measurements(row)
            for start, stop, cell in _columns(row):
                if new_page and start < name_end and not d._text(cell).strip():
                    origin = origins.get((start, stop))
                    if origin is None:
                        raise ValueError("观察续页名称为空且没有明确的前项名称，不能猜填。")
                    _copy_name(cell, origin)
                if new_page:
                    for merge in d._xp(cell, "./w:tcPr/w:vMerge"):
                        if merge.get(qn("w:val"), "continue") == "continue":
                            merge.set(qn("w:val"), "restart")
                if start == remarks_start and (continued or new_page and len(chunks) > 1):
                    _trim_remarks(cell)
            chunk.append(row)
        rp._remove_marks(doc, PREFIX)
        # Repeated headers can push a tail onto a new page; release stale later breaks.
        identity = rp._metadata_caption(table)
        if identity.startswith(FLOW_PREFIX):
            while True:
                separator = table.getnext()
                following = separator.getnext() if separator is not None else None
                if (separator is None or separator.tag != qn("w:p")
                        or following is None or following.tag != qn("w:tbl")
                        or rp._metadata_caption(following) != identity):
                    break
                if (d._text(separator).strip()
                        or d._xp(separator, ".//w:sectPr | .//w:bookmarkStart | "
                                           ".//w:drawing | .//w:fldChar | .//w:fldSimple")):
                    raise ValueError("观察自动续表间出现有效内容，不能安全连续重排。")
                for following_row in d._rows(following)[1:]:
                    chunks[-1].append(following_row)
                following.getparent().remove(following)
                separator.getparent().remove(separator)
        for index, chunk in enumerate(chunks):
            _remove_measurements(chunk)
            table.addprevious(chunk)
            if index:
                d._start_continuation_table(chunk)
                d._explicit_page_start(chunk.getprevious())
        table.getparent().remove(table)
        return True
    return changed
