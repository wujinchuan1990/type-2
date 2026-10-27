"""Read paragraph-style flowmeter records without applying water-meter rules.

Records expose the usual ``fields/records/warnings/outline`` envelope. Each
record has flat, source-located metric ``rows`` plus untouched ``raw_rows``.
Flow matching uses explicit parameters, never table order or computed errors.
"""

from copy import deepcopy
from decimal import Decimal, InvalidOperation, localcontext
from pathlib import Path
import re

from docx.oxml.ns import qn
from docx.text.paragraph import Paragraph

import documents
from curve_results import _manual_result


MAIN = "最大允许误差、测量重复性"
RETEST = "计量性能复测"
NUMBER = r"[+\-−]?(?:\d+(?:\.\d*)?|\.\d+)"
FLOW_UNIT = r"m(?:3|³|\^3)/h"


def _compact(text):
    return re.sub(r"\s+", "", text).replace("：", ":")


def _project_name(text):
    text = re.sub(r"^\d+[.．、]\s*", "", text.strip())
    text = _compact(text).rstrip("。.")
    if text in (MAIN, "最大允许误差和测量重复性"):
        return MAIN
    if text == "耐压强度、贮存环境性能、电磁兼容试验后计量性能复测":
        return RETEST
    return text


def _decimal(text):
    text = _compact(str(text)).replace("−", "-")
    if not re.fullmatch(NUMBER, text):
        return None
    try:
        value = Decimal(text)
        return value if value.is_finite() else None
    except InvalidOperation:
        return None


def _raw(cell):
    return "\n".join(Paragraph(p, None).text for p in cell.findall(qn("w:p")))


def _grid(table):
    """Expand grid spans while retaining each effective cell's physical origin."""
    origins = {}
    rows = []
    for ri, (physical, owners) in enumerate(documents._effective_rows(table), 1):
        row = []
        tr = physical[0].getparent() if physical else None
        before = tr.find("./" + qn("w:trPr") + "/" + qn("w:gridBefore")) if tr is not None else None
        if before is not None:
            row.extend([None] * int(before.get(qn("w:val"), "0")))
        for ci, (cell, owner) in enumerate(zip(physical, owners), 1):
            origins.setdefault(cell, (ri, ci))
            span = cell.find("./" + qn("w:tcPr") + "/" + qn("w:gridSpan"))
            width = int(span.get(qn("w:val"), "1")) if span is not None else 1
            if owner not in origins:
                raise ValueError("无法定位合并单元格的原始行列")
            item = {"text": _raw(owner), "origin": origins[owner], "cell": owner}
            row.extend([item] * width)
        rows.append(row)
    return rows


def _location(source, block, table=None, row=None, column=None):
    value = "%s · 正文块%d" % (source, block)
    if table is not None:
        value += "/表%d" % table
    if row is not None:
        value += "/行%d" % row
    if column is not None:
        value += "/列%d" % column
    return value


def _between(text, start, end=r"$"):
    match = re.search("(?:" + start + r")\s*:?\s*(.*?)\s*(?=" + end + ")", text)
    return match.group(1).strip() if match else None


def _blank_value(text):
    return text.strip().strip("()（）_ ").strip()


def _set(context, key, value, source):
    if value is not None:
        old = context.get(key, "")
        if old and value and old != value:
            raise ValueError("%s：同一记录的%s存在冲突" % (source, key))
        context[key] = value
        context["provenance"][key] = source


def _parameter(context, key, value, source):
    candidates = context.setdefault("parameters", {}).setdefault(key, [])
    candidates.append({"value": value, "unit": "m3/h", "source": source})
    if not context.get(key):
        context[key] = value
        context["provenance"][key] = source


def _sample_paragraph(context, text, source):
    from atmospheric_pressure import extract_pressure

    context.setdefault("atmospheric_pressure_candidates", []).extend(extract_pressure(text, source))
    compact = _compact(text)
    for key, start, end in (
        ("sample_name", "样品名称", r"证书编号|$"),
        ("certificate", "证书编号", r"$"),
        ("model", r"型号/规格|规格型号", r"制造厂|$"),
        ("manufacturer", "制造厂", r"$"),
        ("sample_id", r"^(?:样品)?编号", r"准确度|样品状态|$"),
    ):
        # Labels may have typography spaces, but sample/model values retain spaces.
        spaced_start = start
        spaced_end = end
        for label in ("型号/规格", "制造厂", "编号"):
            replacement = r"\s*".join(re.escape(c) for c in label)
            spaced_start = spaced_start.replace(label, replacement)
            spaced_end = spaced_end.replace(label, replacement)
        value = _between(text.replace("：", ":"), spaced_start, spaced_end)
        if value is not None:
            _set(context, key, _blank_value(value), source)
    if "环境条件" in compact:
        for key, label, unit, next_label in (
            ("temperature", "温度", "℃|°C", "湿度|大气压|其它|其他|$"),
            ("humidity", "湿度", "%RH|%rh|％RH|%", "大气压|其它|其他|$"),
        ):
            value = _between(compact, label, next_label)
            if value is not None:
                match = re.search(r"(" + unit + r")", value)
                _set(context, key, _blank_value(value[:match.start()] if match else value), source)
                context[key + "_unit"] = match.group() if match else ""
    if "流量范围" in compact:
        match = re.search(
            r"流量范围:?[(（]?(" + NUMBER + r")[~～至-](" + NUMBER + r")[)）]?"
            + FLOW_UNIT, compact, re.I,
        )
        if match:
            _parameter(context, "qmin", match[1], source)
            _parameter(context, "qmax", match[2], source)
        else:
            context.setdefault("issues", []).append(source + "：流量范围缺失或格式/单位不明确")
    for key in ("qt", "qmin", "qmax"):
        label = r"(?:分界流量(?:qt)?|qt)" if key == "qt" else key
        for match in re.finditer(
            r"(?<![A-Za-z])" + label + r"(?:[（(]" + FLOW_UNIT + r"[）)])?"
            + r"(?::|=)?(" + NUMBER + r")(?:" + FLOW_UNIT + r")?"
            + r"(?=$|[;,；，]|qmin|qmax|qt|分界流量)",
            compact, re.I,
        ):
            _parameter(context, key, match[1], source)


def _headers(row):
    names = []
    for item in row:
        text = _compact(item["text"]) if item else ""
        name = re.sub(r"[（(].*?[）)]", "", text)
        names.append(name)
    required = ("检测流量点", "序号", "流量计读数", "标准器读数", "示值误差", "平均示值误差", "重复性")
    if all(names.count(name) == 1 for name in required):
        return {name: names.index(name) for name in required}
    return None


def _table_measurements(record, table, filename, block, table_index):
    grid = _grid(table)
    headers = _headers(grid[0]) if grid else None
    if not headers:
        raise ValueError("%s：不支持的流量计测量表头" % _location(filename, block, table_index))
    units = {}
    for name, ci in headers.items():
        header = _compact(grid[0][ci]["text"])
        unit = re.search(r"[（(](.*?)[）)]", header)
        units[name] = unit[1] if unit else ""
    if not re.fullmatch(FLOW_UNIT, units["检测流量点"], re.I):
        raise ValueError("检测流量点须有明确的m3/h单位")
    for name in ("示值误差", "平均示值误差", "重复性"):
        if units[name] not in ("%", "％"):
            raise ValueError("%s须有明确的百分数单位" % name)
    for name in ("流量计读数", "标准器读数"):
        if units[name].lower() not in ("l", "m3", "m³"):
            raise ValueError("%s缺少支持的体积单位" % name)
    seen = set()
    for ri, row in enumerate(grid[1:], 2):
        if _headers(row):
            continue
        if len(row) != len(grid[0]) or any(item is None for item in row):
            raise ValueError("测量数据列数与表头不一致或存在缺失网格")
        active = [name for name in headers if name not in ("序号", "检测流量点")
                  and row[headers[name]]["text"].strip()]
        if not active:
            continue
        flow_item = row[headers["检测流量点"]]
        flow = flow_item["text"].strip()
        if _decimal(flow) is None:
            raise ValueError("%s：实测数据缺少明确的实际流量" % _location(filename, block, table_index, ri))
        # A summary owner may span unused retest slots; it is evidence only once.
        fresh = [name for name in active if (name, row[headers[name]]["origin"]) not in seen]
        if not fresh:
            continue
        record["raw_rows"].append({
            "source": _location(filename, block, table_index, ri),
            "cells": [{"raw": item["text"], "source": _location(
                filename, block, table_index, *item["origin"])} for item in row],
        })
        ordinal = row[headers["序号"]]["text"].strip()
        for name in ("检测流量点", "序号") + tuple(fresh):
            item = row[headers[name]]
            identity = (name, item["origin"])
            if identity in seen:
                continue
            seen.add(identity)
            if not item["text"].strip():
                continue
            record["rows"].append({
                "label": "%s | %s | %s" % (flow, ordinal, name),
                "metric": name, "value": item["text"].strip(), "raw": item["text"],
                "unit": units[name], "flow": flow, "flow_unit": units["检测流量点"],
                "flow_source": _location(filename, block, table_index, *flow_item["origin"]),
                "repeat": ordinal, "conditions": {},
                "source": _location(filename, block, table_index, *item["origin"]),
            })


def parse_upload(path):
    """Parse supported flowmeter sheets; never supply missing source values."""
    doc = documents._load(path)
    filename = Path(path).name
    result = {"fields": {}, "records": [], "warnings": [], "outline": {}}
    context, current, project = None, None, ""
    table_index = 0

    def finish():
        if context is not None and current is None:
            raise ValueError("样品记录缺少项目或支持的测量数据表")
        if current is not None and not current["rows"]:
            raise ValueError(current["source"] + "：没有实测数据")

    for block, element in enumerate(doc.element.body, 1):
        if element.tag == qn("w:p"):
            text = Paragraph(element, doc).text
            compact = _compact(text)
            source = _location(filename, block)
            if "测量记录单" in compact or "样品名称" in compact:
                if "测量记录单" in compact:
                    finish()
                    context, current, project = None, None, ""
                    continue
                finish()
                context = {key: "" for key in (
                    "sample_id", "model", "certificate", "temperature", "humidity",
                    "temperature_unit", "humidity_unit", "start", "end", "qmin", "qmax", "qt",
                )}
                context.update(provenance={}, parameters={}, issues=[])
                current, project = None, ""
            if "测量数据" in compact:
                if context is None or current is not None or project:
                    raise ValueError(source + "：测量项目缺少独立的前置样品身份信息")
                match = re.search(r"测量数据:(.+)", compact)
                project = _project_name(match[1]) if match else ""
                if project not in (MAIN, RETEST):
                    raise ValueError(source + "：缺少或不支持的流量计项目名称：" + project)
            elif "检测日期" in compact:
                if current is None:
                    raise ValueError(source + "：检测日期没有对应的测量记录")
                start, end = documents._dates(text)
                current["start"], current["end"] = start, end
                current["provenance"]["time"] = source
            elif context is not None:
                if current is not None and any(label in compact for label in (
                    "型号/规格", "规格型号", "证书编号", "环境条件", "流量范围",
                )):
                    raise ValueError(source + "：新记录未明确给出样品边界")
                _sample_paragraph(context, text, source)
        elif element.tag == qn("w:tbl"):
            table_index += 1
            source = _location(filename, block, table_index)
            if context is None or not project or not context["sample_id"] or not context["model"]:
                raise ValueError(source + "：测量表缺少明确的样品编号、型号或项目")
            if element.findall(".//" + qn("w:tbl")):
                raise ValueError(source + "：不支持嵌套测量表")
            if current is None:
                current = dict(deepcopy(context), project=project, source=source, rows=[], raw_rows=[])
                result["records"].append(current)
            _table_measurements(current, element, filename, block, table_index)
    finish()
    if not result["records"]:
        raise ValueError("未找到支持的流量计样品测量记录")
    for record in result["records"]:
        for key, field in (("certificate", "report_number"), ("sample_name", "instrument_name"),
                           ("manufacturer", "manufacturer"), ("model", "model")):
            if record.get(key):
                candidate = {"value": record[key], "source": record["provenance"][key]}
                result["fields"].setdefault(field, []).append(candidate)
        _, issues = documents._record_conditions([record])
        result["warnings"].extend(record["issues"] + issues)
        for key in ("qmin", "qmax"):
            if not record.get(key):
                result["warnings"].append(record["source"] + "：未识别明确的" + key)
    certificates = {entry["value"] for entry in result["fields"].get("report_number", [])}
    if not certificates:
        result["warnings"].append("未识别到样品证书编号对应的报告编号，请手动核对。")
    elif len(certificates) > 1:
        result["warnings"].append("原始记录的报告编号不一致，请核对全部候选。")
    result["warnings"] = list(dict.fromkeys(result["warnings"]))
    return result


def _model_key(model, dn=""):
    compact = _compact(model).upper()
    match = re.search(r"DN(\d+(?:\.\d+)?)", compact)
    if match:
        diameter = Decimal(match[1])
        compact = compact[:match.start()] + compact[match.end():]
    else:
        diameter = _decimal(re.sub(r"(?i)DN|mm|[（）()]", "", _compact(dn)))
    return (compact.strip("/-"), diameter) if compact and diameter is not None else None


def _outline_qt(outline, source):
    if outline is None:
        return {}
    doc = documents._load(outline) if isinstance(outline, (str, Path)) else outline
    if not hasattr(doc, "element"):
        raise ValueError("qt识别需要本次大纲的DOCX路径或原生Document")
    _, table = documents.outline_samples(doc)
    block = next(index for index, element in enumerate(doc.element.body, 1)
                 if element is table or table in element.iterdescendants())
    table_index = list(doc.element.body.iter(qn("w:tbl"))).index(table) + 1
    values = {}
    headers = None
    for row in _grid(table):
        texts = [_compact(item["text"]).lower() if item else "" for item in row]
        qt_columns = [i for i, text in enumerate(texts) if re.fullmatch(
            r"(?:分界流量(?:qt)?|qt)(?:[（(]" + FLOW_UNIT + r"[）)])?", text)]
        model_columns = [i for i, text in enumerate(texts)
                         if text in ("型号", "型号/规格", "规格型号", "样机型号规格")]
        dn_columns = [i for i, text in enumerate(texts) if re.fullmatch(
            r"(?:规格|dn|(?:公称)?(?:通径|口径)(?:dn)?)(?:[（(]mm[）)])?", text)]
        if len(qt_columns) == len(model_columns) == 1:
            headers = (qt_columns[0], model_columns[0],
                       dn_columns[0] if len(dn_columns) == 1 else None)
            continue
        if headers is None:
            continue
        qi, mi, di = headers
        if max(qi, mi, di or 0) >= len(row) or row[qi] is None or row[mi] is None:
            continue
        key = _model_key(texts[mi], texts[di] if di is not None else "")
        if key is None:
            continue
        item = row[qi]
        raw = item["text"].strip()
        value = re.sub(FLOW_UNIT + r"$", "", _compact(raw), flags=re.I)
        values.setdefault(key, []).append({
            "value": value, "unit": "m3/h",
            "source": _location(source, block, table_index, *item["origin"]),
        })
    return values


def _explicit(record, key, outline_values):
    candidates = list(record.get("parameters", {}).get(key, []))
    if not candidates and record.get(key):
        candidates.append({"value": record[key], "source": record.get("provenance", {}).get(key, record["source"])})
    if key == "qt":
        candidates += outline_values.get(_model_key(record["model"]), [])
    if not candidates:
        return None, "", "缺少明确的%s参数，不能按流量点顺序或比例猜测" % key
    values = [_decimal(candidate["value"]) for candidate in candidates]
    if any(value is None or value <= 0 for value in values) or len(set(values)) != 1:
        return None, "", "%s参数缺失、无效或冲突，请核对原记录与大纲" % key
    return values[0], "；".join(dict.fromkeys(candidate["source"] for candidate in candidates)), ""


def _target_flow(requirement, record, outline_values):
    text = _compact(requirement).lower()
    symbols = list(re.finditer(r"(?<![a-z])(" + NUMBER + r")?(qmax|qmin|qt)(?![a-z])", text))
    direct = list(re.finditer(r"(" + NUMBER + r")" + FLOW_UNIT, text, re.I))
    if len(symbols) + len(direct) != 1:
        return None, "", "技术要求没有唯一明确的流量点（qmax、倍率qmax、qt、qmin或带单位数值）"
    if direct:
        target = _decimal(direct[0][1])
        if target is None or target <= 0:
            return None, "", "目标流量必须为明确正数，才能按5%相对差异匹配"
        return target, "", ""
    match = symbols[0]
    factor = _decimal(match[1]) if match[1] else Decimal(1)
    parameter, source, issue = _explicit(record, match[2], outline_values)
    if issue:
        return None, "", issue
    if factor is None or factor <= 0 or (match[2] != "qmax" and factor != 1):
        return None, "", "不支持或无效的流量倍率要求"
    with localcontext() as ctx:
        ctx.prec = len(parameter.as_tuple().digits) + len(factor.as_tuple().digits) + 10
        return parameter * factor, source, ""


def _flow_matches(flow, target):
    actual = _decimal(flow)
    if actual is None:
        return False
    with localcontext() as ctx:
        ctx.prec = (max(actual.adjusted(), target.adjusted())
                    - min(actual.as_tuple().exponent, target.as_tuple().exponent) + 4)
        return abs(actual - target) <= target * Decimal("0.05")


def _match(row, records, name, outline_values):
    sample_id = row.get("sample_id", "")
    if not sample_id:
        label = row.get("sample_label", "")
        if label and any(record["sample_id"] == label for record in records):
            sample_id = label
    if not sample_id:
        return "", "", "", "缺少真实样机编号对应关系，请核对样机代号映射"
    matches = [record for record in records if record["sample_id"] == sample_id]
    if len(matches) != 1:
        return "", "", "", ("同一样机匹配多个测量批次，不能自动选择" if matches
                            else "未找到该真实样机编号的项目记录：" + sample_id)
    record = matches[0]
    requirement = row.get("requirement", "")
    if name == "耐压强度" and record.get("parser", "").startswith("gas-flowmeter-excel-"):
        from gas_flowmeter_excel import PRESSURE, pressure_requirement

        candidates = [value for value in record["rows"] if value["metric"] == "耐压强度"]
        pressure = list(PRESSURE.finditer(requirement))
        if len(candidates) != 1 or len(pressure) != 1:
            return "", "", "", "耐压结果或技术要求中的最大工作压力不唯一，不能自动采用"
        value = candidates[0]
        try:
            resolved = pressure_requirement(requirement, value["maximum_working_pressure"])
        except ValueError as exc:
            return "", "", "", str(exc)
        row["working_pressure"] = {"value": value["maximum_working_pressure"], "source": value["source"]}
        row["resolved_requirement"] = resolved
        return value["value"], value["source"], sample_id, ""
    error_metric = "最大允许误差" in requirement or "平均示值误差" in requirement
    repeat_metric = "测量重复性" in requirement or "重复性" in requirement
    if name == RETEST:
        metric = "平均示值误差"
    elif error_metric == repeat_metric:
        return "", "", "", "技术要求未唯一指定最大允许误差或测量重复性"
    else:
        metric = "平均示值误差" if error_metric else "重复性"
    flow_requirement = requirement
    if record.get("parser", "").startswith("gas-flowmeter-excel-") and "|" in requirement:
        parent, detail = requirement.rsplit("|", 1)
        parent = re.sub(r"(?:qmin|qt)[≤<]q[≤<](?:qt|qmax)", "", _compact(parent).lower())
        points = re.findall(r"(?<![a-z])(?:" + NUMBER + r")?(?:qmax|qmin|qt)(?![a-z])|"
                            + NUMBER + FLOW_UNIT, parent, re.I)
        if not points or (name == RETEST and all(point == _compact(detail).lower() for point in points)):
            flow_requirement = detail
    target, parameter_source, issue = _target_flow(flow_requirement, record, outline_values)
    if issue:
        return "", "", "", issue
    is_qt = bool(re.search(r"(?<![a-z])qt(?![a-z])", _compact(flow_requirement).lower()))
    candidates = [value for value in record["rows"]
                  if value["metric"] == metric and _flow_matches(value["flow"], target)]
    if len(candidates) != 1:
        actual = "、".join(dict.fromkeys(value["flow"] for value in record["rows"]))
        return "", "", "", (
            "目标流量%s m3/h的±5%%范围内匹配多个流量点或多个汇总值，不能自动选择" % target
            if candidates else
            "目标流量%s m3/h的±5%%范围内与记录不匹配或缺少%s；实际流量：%s"
            % (target, metric, actual))
    value = candidates[0]
    raw = value["value"].strip()
    if _decimal(raw.rstrip("%％")) is None or value.get("unit") not in ("%", "％"):
        return "", "", "", "源汇总值格式或百分数单位不明确，不能计算替代"
    display = raw if raw.endswith(("%", "％")) else raw + "%"
    source = "；".join(filter(None, (
        value["source"], "实际流量%s m3/h（%s）" % (value["flow"], value["flow_source"]),
        "%s目标流量%s m3/h，按|实际流量-目标流量|≤目标流量×5%%匹配" % ("qt" if is_qt else "", target),
        parameter_source,
    )))
    return display, source, sample_id, ""


def apply_records(project, parsed, outline=None, outline_source="当前大纲"):
    """Refresh one Project-compatible dict, preserving manual values and clears."""
    result = deepcopy(project)
    name = _project_name(project["name"])
    parsed = [parsed] if isinstance(parsed, dict) else parsed
    records = [record for data in parsed for record in data.get("records", [])
               if _project_name(record.get("project", "")) == name]
    outline_values = _outline_qt(outline, outline_source)
    old_summary = result.get("record_summary", {})
    for key, value in old_summary.get("values", {}).items():
        if result.get(key) == value:
            result[key] = ""
    values, issues = documents._record_conditions(records) if records else ({}, [])
    result.update(values)
    result["record_summary"] = {
        "values": values, "issues": issues,
        "records": [{key: record.get(key, "") for key in (
            "source", "sample_id", "model", "start", "end", "temperature", "humidity",
        )} for record in records],
    }
    for evidence, record in zip(result["record_summary"]["records"], records):
        for key in ("temperature", "humidity"):
            evidence[key] = " ".join(filter(None, (record.get(key, ""), record.get(key + "_unit", ""))))
    result["source"] = "；".join(dict.fromkeys(record["source"] for record in records))
    notes = result.setdefault("notes", [])
    notes[:] = [note for note in notes if not note.startswith("流量计记录：")]
    if not records:
        notes.append("流量计记录：缺失原始记录：" + project["name"])
    for row in result.get("rows", []):
        row.update(working_pressure=None, resolved_requirement="")
        value, source, sample_id, issue = _match(row, records, name, outline_values)
        manual = _manual_result(row) or any(word in row.get("source", "") for word in ("人工", "手工"))
        if manual:
            if name == "耐压强度":
                row["match_issue"] = issue
            if row.get("matched_sample_id") and row["matched_sample_id"] != row.get("sample_id"):
                row["match_issue"] = "样机对应已改变，保留的人工结果属于旧编号，请核对后重新保存"
            continue
        row.update(value=value, auto_value=value, source=source, matched_sample_id=sample_id, match_issue=issue)
        if value and not row.get("sample_id"):
            row["sample_id"] = sample_id
        if issue:
            warning = "流量计记录：" + issue
            if warning not in notes:
                notes.append(warning)
    result["confirmed"] = False
    return result
