"""Model-scoped gas flow requirements shared by splitting and composition."""

from copy import deepcopy
from decimal import Decimal
import re
import unicodedata

from docx.oxml.ns import qn

import documents as d
import observation_pagination as op


NUMBER = r"(?:[0-9]+(?:\.[0-9]+)?|\.[0-9]+)"
UNIT = r"m\s*[3³]\s*/\s*h"
BLOCK = re.compile(
    r"型号\s*[：:]\s*(?P<model>.+?)\s*"
    r"最大流量\s*[：:]\s*(?P<maximum>" + NUMBER + r")\s*" + UNIT + r"\s*[，,；;]?\s*"
    r"最小流量\s*[：:]\s*(?P<minimum>" + NUMBER + r")\s*" + UNIT + r"\s*[，,；;]?\s*"
    r"分界流量\s*[：:]\s*(?P<transition>" + NUMBER + r")\s*" + UNIT + r"[。.]?",
    re.S,
)


def _cells(doc):
    tables = [t for t in d._tables(doc) if d._rows(t)
              and d._norm(d._text(d._cells(d._rows(t)[0])[0])) == "型式评价大纲章节号及项目名称"]
    found, identities = [], set()
    for table, row, identity in op._project_rows(tables):
        names = identity[1][1]
        if not any(re.fullmatch(r"\d*流量范围", name) for _a, _b, name in names):
            continue
        identities.add(names)
        headers = list(op._columns(d._rows(table)[0]))
        requirement = headers[1][:2]
        matches = [cell for start, stop, cell in op._columns(row) if (start, stop) == requirement]
        if len(matches) != 1:
            raise ValueError("燃气表流量范围的要求列无法唯一定位。")
        found.extend(matches)
    if len(identities) > 1:
        raise ValueError("燃气表存在多个不同的流量范围项目，不能猜测型号对应。")
    return found


def _read(cells):
    text = "".join(op._text(cell) for cell in cells)
    matches = list(BLOCK.finditer(text))
    cursor = 0
    for match in matches:
        if text[cursor:match.start()].strip():
            raise ValueError("燃气表流量范围含无法识别的要求内容，不能删去或猜填。")
        cursor = match.end()
    if not matches or text[cursor:].strip():
        raise ValueError("燃气表流量范围须为从型号到分界流量m3/h的完整要求段。")
    return text, matches


def describe(doc):
    cells = _cells(doc)
    if not cells:
        return None
    _text, matches = _read(cells)
    return {"blocks": [match.groupdict() for match in matches]}


def _raw_offset(node, position):
    logical, raw = 0, 0
    for child in node.iter():
        text = op._leaf_text(child)
        if text is None:
            continue
        if child.tag == qn("w:t"):
            raw += max(0, min(len(text), position - logical))
        logical += len(text)
    return raw


def _block_copy(cells, first, last):
    result = deepcopy(cells[0])
    for child in list(result):
        if child.tag != qn("w:tcPr"):
            result.remove(child)
    offset = 0
    for cell in cells:
        for paragraph in d._xp(cell, "./w:p"):
            text = op._text(paragraph)
            start, end = max(0, first - offset), min(len(text), last - offset)
            if start < end:
                result.append(op._slice(paragraph, op._units(text[:start]), op._units(text[:end])))
            offset += len(text)
    return result


def _rewrite(doc, values):
    from separation_experimental_documents import _plain_cell

    cells = _cells(doc)
    if not cells:
        return None
    for cell in cells:
        _plain_cell(cell)
    _text, matches = _read(cells)
    replacements = []
    for index, fields in enumerate(values):
        source = matches[min(index, len(matches) - 1)]
        block = _block_copy(cells, source.start(), source.end())
        _block_text, parsed = _read([block])
        if len(parsed) != 1:
            raise ValueError("燃气表流量范围要求段复制不完整，停止输出。")
        for key in sorted(fields, key=lambda key: parsed[0].start(key), reverse=True):
            start, end = parsed[0].span(key)
            d._replace_text_span(block, _raw_offset(block, start), _raw_offset(block, end), fields[key])
        replacements.extend(list(d._xp(block, "./w:p")))
    for cell in cells:
        d._put(cell, "")
    first = cells[0]
    for paragraph in d._xp(first, "./w:p"):
        first.remove(paragraph)
    first.extend(replacements)
    actual = describe(doc)
    if len(actual["blocks"]) != len(values) or any(
            any(block[key] != value for key, value in expected.items())
            for block, expected in zip(actual["blocks"], values)):
        raise ValueError("燃气表流量范围的型号数量或填写内容校验失败。")
    return actual


def split(doc, model_count):
    is_gas = any(
        d._norm(d._text(cell)) == "计量器具名称" and "燃气表" in d._text(cells[index + 1])
        for table in d._tables(doc) for row in d._rows(table)
        for cells in [d._cells(row)] for index, cell in enumerate(cells[:-1])
    )
    if not is_gas:
        return None
    if not 1 <= model_count <= 26:
        raise ValueError("燃气表流量范围支持1～26个型号。")
    return _rewrite(doc, [{"model": chr(65 + index) * 3} for index in range(model_count)])


def _number(value):
    text = format(value, "f")
    return text.rstrip("0").rstrip(".") if "." in text else text


def fill(doc, task):
    if task.get("report_type") != "燃气表" or not _cells(doc):
        return None
    schema = task.get("sample_schema", {})
    headers = [re.sub(r"\s+", "", unicodedata.normalize("NFKC", h))
               for h in schema.get("headers", [])]
    models = [i for i, h in enumerate(headers) if h in {"型号", "规格型号", "型号规格"}]
    flows = [i for i, h in enumerate(headers) if re.fullmatch(r"流量范围(?:\(m3/h\))?", h)]
    if len(models) != 1 or len(flows) != 1 or not schema.get("groups"):
        raise ValueError("燃气表流量范围须有基本信息总览中的唯一型号列、流量范围（m3/h）列和型号分组。")
    values = []
    for index, group in enumerate(schema["groups"], 1):
        fields = group.get("values", [])
        if len(fields) != len(headers):
            raise ValueError(f"燃气表第{index}组流量范围与总览表头不对应。")
        model = fields[models[0]].strip()
        if "规格" in headers:
            model = " ".join(v for v in (model, fields[headers.index("规格")].strip()) if v)
        if not fields[models[0]].strip() or re.search(r"[\r\n\v]", model):
            raise ValueError(f"燃气表第{index}组型号缺失或含换行，不能保留素材占位值。")
        flow = re.sub(r"\s+", "", unicodedata.normalize("NFKC", fields[flows[0]]))
        match = re.fullmatch(r"(" + NUMBER + r")[~～至\-–—](" + NUMBER + r")(?:m3/h)?", flow)
        if match is None:
            raise ValueError(f"燃气表第{index}组流量范围须为明确的最小值~最大值（m3/h），不能沿用素材旧值。")
        minimum, maximum = (Decimal(value) for value in match.groups())
        if not 0 <= minimum < maximum:
            raise ValueError(f"燃气表第{index}组流量范围须满足0≤最小流量<最大流量。")
        values.append({"model": model, "minimum": _number(minimum), "maximum": _number(maximum),
                       "transition": _number(maximum * Decimal("0.1"))})
    from separation_experimental_documents import _remove_ephemeral_bookmarks

    _remove_ephemeral_bookmarks(doc)
    return _rewrite(doc, values)
