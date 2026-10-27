"""V3 rebuilds observation rosters from the current outline, never placeholders."""

from copy import deepcopy
import re

from docx.oxml.ns import qn

import documents as d
import observation_pagination as op
from observation_documents import fill_observation_remarks
from separation_experimental_documents import _plain_cell, _remove_ephemeral_bookmarks
from separation_experimental_format import basic_labels
from separation_experimental_layout import observation_tables
from report_types import INDEPENDENT_METER_TYPES


def _lines(cell):
    return [line.strip() for paragraph in d._xp(cell, "./w:p")
            for line in op._text(paragraph).splitlines() if line.strip()]


def source_layout(doc):
    candidates = []
    for table in d._tables(doc):
        rows = d._rows(table)
        for index, row in enumerate(rows):
            cells = d._cells(row)
            if len(cells) == 4 and d._norm(d._text(cells[0])) == "规格型号":
                candidates.append((table, index, "four"))
                break
            if [d._norm(d._text(c)) for c in cells] == [
                    "序号", "型号", "口径", "准确度等级", "样机编号"]:
                candidates.append((table, index, "five"))
                break
    if len(candidates) != 1:
        raise ValueError("V3观察基本信息的型号表缺失或不唯一，不能按文件名猜测。")
    table, start, shape = candidates[0]
    rows, groups, model_rows = d._rows(table), [], []
    if start != 1 or len(d._cells(rows[0])) != 4:
        raise ValueError("V3观察基本信息不是申请单位行之后的连续型号表。")
    for row in rows[start + (shape == "five"):]:
        cells = d._cells(row)
        if shape == "four":
            if (len(cells) != 4 or d._norm(d._text(cells[0])) not in {"规格型号", ""}
                    or d._norm(d._text(cells[2])) not in {"样机编号", ""}):
                raise ValueError("V3观察型号行混有其他信息，不能安全重建。")
            models, specs, numbers = _lines(cells[1]), {}, _lines(cells[3])
            details = [re.fullmatch(r"([^：:]+)[：:](.+)", line) for line in models[1:]]
            if details and all(details):
                specs = {match[1]: match[2] for match in details}
                models = models[:1]
            model = "\n".join(models)
        else:
            if len(cells) != 5:
                raise ValueError("V3流量计型号表混有非五列数据。")
            model, specs, numbers = "\n".join(_lines(cells[1])), {
                "口径": "\n".join(_lines(cells[2])),
                "准确度等级": "\n".join(_lines(cells[3])),
            }, _lines(cells[4])
        if not model or not numbers:
            raise ValueError("V3源型号或样机编号为空，无法确认项目内旧编号的所属范围。")
        if d._xp(row, "./w:tc/w:tcPr/w:vMerge"):
            raise ValueError("V3源型号行存在纵向合并，不能安全映射。")
        if shape == "four" and len(models) > 1:
            prefixes = [re.fullmatch(r"([A-Z]+)\d{3}", number) for number in numbers]
            order = list(dict.fromkeys(match[1] for match in prefixes if match))
            if (not all(prefixes) or order != [chr(65 + i) for i in range(len(models))]):
                raise ValueError("V3源同一格有多个型号，但编号没有明确的A/B分组，不能猜测型号数量或归属。")
            groups.extend({"model": value, "specs": {},
                           "numbers": [number for number, match in zip(numbers, prefixes) if match[1] == prefix]}
                          for value, prefix in zip(models, order))
        else:
            groups.append({"model": model, "specs": specs, "numbers": numbers})
        model_rows.append(row)
    numbers = [number for group in groups for number in group["numbers"]]
    if not groups or len(numbers) != len(set(numbers)):
        raise ValueError("V3源型号表没有有效分组或样机编号重复。")
    return table, shape, model_rows, groups


def target_roster(outline, task):
    schema, _table = d.outline_samples(outline)
    if task.get("report_type") in INDEPENDENT_METER_TYPES:
        from flowmeter_tasks import reconcile_samples

        task = reconcile_samples(deepcopy(task), schema, report_type=task["report_type"])
    else:
        from water_results import reconcile_task

        task = reconcile_task(deepcopy(task), schema)
    schema = task["sample_schema"]
    headers = [d._norm(header) for header in schema["headers"]]
    model_columns = [i for i, header in enumerate(headers) if header in {"型号", "型号规格", "规格型号"}]
    if len(model_columns) != 1:
        raise ValueError("V3本次大纲样机表必须有唯一型号列。")
    model_column = model_columns[0]
    samples = {sample["label"]: sample for sample in task["samples"]}
    labels = [label for group in schema["groups"] for label in group["labels"]]
    if len(samples) != len(task["samples"]) or set(samples) != set(labels):
        raise ValueError("V3本次样机清单与大纲型号分组不一致。")
    groups, excluded = [], []
    for group in schema["groups"]:
        values = group["values"]
        model = values[model_column].strip()
        if not model:
            raise ValueError("V3本次大纲存在未填写型号，不能保留旧型号代替。")
        specs = {header: value.strip() for i, (header, value) in enumerate(zip(headers, values))
                 if i not in {schema["quantity_column"], model_column} and value.strip()}
        selected = [samples[label] for label in group["labels"] if not samples[label].get("disassembly")]
        excluded.extend(label for label in group["labels"] if samples[label].get("disassembly"))
        groups.append({"model": model, "specs": specs,
                       "labels": [sample["label"] for sample in selected],
                       "numbers": [sample.get("number", "").strip() for sample in selected]})
    numbers = [number for group in groups for number in group["numbers"] if number]
    if len(numbers) != len(set(numbers)):
        raise ValueError("V3本次真实样机编号重复，不能建立编号/叉号对应。")
    if any(re.search(r"[\r\n\v]", number) for number in numbers):
        raise ValueError("V3真实样机编号含换行，不能把一个编号误判为多台样机。")
    return task, groups, excluded


def _put_numbers(cell, numbers):
    _plain_cell(cell)
    template = deepcopy(d._xp(cell, "./w:p")[0])
    for paragraph in d._xp(cell, "./w:p"):
        cell.remove(paragraph)
    for number in numbers or [""]:
        paragraph = deepcopy(template)
        d._put(paragraph, number)
        cell.append(paragraph)


def _four_grid(table):
    from docx.oxml import OxmlElement

    columns = d._xp(table, "./w:tblGrid/w:gridCol")
    widths = [int(column.get(qn("w:w"))) for column in columns]
    top = d._cells(d._rows(table)[0])
    spans = [int(next(iter(d._xp(cell, "./w:tcPr/w:gridSpan/@w:val")), "1")) for cell in top]
    if len(widths) != 5 or spans != [1, 2, 1, 1]:
        raise ValueError("V3流量计首行与五列网格不一致，不能安全转换。")
    widths = [widths[0], widths[1] + widths[2], widths[3], widths[4]]
    for column in columns:
        table.tblGrid.remove(column)
    for width in widths:
        column = OxmlElement("w:gridCol")
        column.set(qn("w:w"), str(width))
        table.tblGrid.append(column)
    for cell, width in zip(top, widths):
        for span in d._xp(cell, "./w:tcPr/w:gridSpan"):
            span.getparent().remove(span)
        cell.get_or_add_tcPr().get_or_add_tcW().set(qn("w:w"), str(width))
        cell.tcPr.tcW.set(qn("w:type"), "dxa")


def rebuild(observe, outline, task):
    _remove_ephemeral_bookmarks(observe)
    table, shape, rows, source = source_layout(observe)
    for row in rows:
        for cell in d._cells(row):
            _plain_cell(cell)
    task, groups, excluded = target_roster(outline, task)
    warnings = []
    fields = {field["key"]: field.get("value", "").strip() for field in task.get("overview", [])}
    if shape == "five":
        for group in groups:
            if not any("口径" in key or "通径" in key or key == "规格" for key in group["specs"]):
                raise ValueError("V3流量计本次大纲没有可映射的口径/规格，不能沿用旧型号的口径。")
            if "准确度等级" not in group["specs"]:
                accuracy = fields.get("accuracy_class", "")
                if not accuracy:
                    raise ValueError("V3流量计本次大纲缺少准确度等级，不能沿用旧样机参数。")
                group["specs"]["准确度等级"] = accuracy
    old_numbers = {number for group in source for number in group["numbers"]}
    numbers = [number for group in groups for number in group["numbers"]]
    project_tables = observation_tables(observe)
    for project_table in project_tables:
        if any(op._text(cell).strip() for row in d._rows(project_table)[1:]
               for _a, _b, cell in list(op._columns(row))[-2:-1]):
            raise ValueError("V3源观察表的−栏非空，不能自动改为＋结论。")
        headers = list(op._columns(d._rows(project_table)[0]))
        if len(headers) != 6:
            continue
        sample_column = headers[2][0]
        for row in d._rows(project_table)[1:]:
            matches = [cell for start, _stop, cell in op._columns(row) if start == sample_column]
            if len(matches) != 1:
                raise ValueError("V3观察项目样机编号列无法唯一定位。")
            cell = matches[0]
            existing = _lines(cell)
            if len(existing) != len(set(existing)) or set(existing) != old_numbers:
                raise ValueError("V3源项目的样机编号不是完整源样机清单，不能猜测子集与新型号的对应关系。")
            _put_numbers(cell, numbers)
    if shape == "five":
        _four_grid(table)
        template = deepcopy(d._rows(table)[0])
        for row in d._rows(table)[1:]:
            table.remove(row)
    else:
        template = deepcopy(rows[0])
        for row in rows:
            table.remove(row)
    for group in groups:
        row = deepcopy(template)
        cells = d._cells(row)
        texts = ["规格型号", "\n".join([group["model"], *(
            f"{key}：{value}" for key, value in group["specs"].items())]), "样机编号"]
        for cell, text in zip(cells, texts):
            d._put(cell, text)
        _put_numbers(cells[3], group["numbers"])
        for cell in cells:
            cell.get_or_add_tcPr().get_or_add_vAlign().set(qn("w:val"), "center")
        for height in d._xp(row, "./w:trPr/w:trHeight"):
            height.set(qn("w:hRule"), "atLeast")
        table.append(row)
    for index, key in ((1, "manufacturer"), (3, "instrument_name")):
        if fields.get(key):
            d._put(d._cells(d._rows(table)[0])[index], fields[key])
        else:
            warnings.append(f"V3观察基本信息/{key}：本次总览未填写，保留源文字，请核对。")
    basic_labels(observe)
    warnings.extend(fill_observation_remarks(observe, outline, task, fill_basics=False)["issues"])
    missing = [label for group in groups for label, number in zip(group["labels"], group["numbers"]) if not number]
    if missing:
        warnings.append("V3观察样机真实编号未填写，编号格留空且不填×：" + "、".join(missing))
    warnings.append(f"V3观察重建：源{len(source)}组型号/规格、{len(old_numbers)}台；"
                    f"本次{len(groups)}组、观察样机{len(numbers)}台，排除拆样{len(excluded)}台。")
    return {"warnings": warnings, "source_groups": source, "target_groups": groups,
            "excluded_labels": excluded}
