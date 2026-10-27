"""Apply only the outline-dependent parameters in the four requested result projects."""

import re
from copy import deepcopy
from decimal import Decimal

import documents as d
from observation_documents import _label_pattern


PROJECT_FIELDS = {
    "水压": ("maximum_pressure",),
    "静压力": ("maximum_pressure",),
    "安装条件": ("upstream", "downstream"),
    "压力损失": ("pressure_loss", "loss_limit"),
}
TARGETS = {
    "maximum_pressure": ("最大允许压力", r"\d+(?:\.\d+)?\s*MPa"),
    "upstream": ("上游流场敏感度等级", r"U\s*\d+"),
    "downstream": ("下游流场敏感度等级", r"D\s*\d+"),
    "pressure_loss": ("压力损失等级", r"[△Δ]\s*p\s*\d+(?:\.\d+)?"),
}
ISSUE_PREFIX = "大纲技术参数/"


def parameter_values(parameters):
    values = parameters.get("values", {})
    result = {}
    pressure = re.fullmatch(r"\s*(\d+(?:\.\d+)?)\s*(MPa|kPa|Pa)\s*",
                            values.get("maximum_pressure", ""), re.I)
    if pressure and Decimal(pressure[1]) > 0:
        if pressure[2].lower() == "mpa":
            result["maximum_pressure"] = pressure[1] + "MPa"
        else:
            divisor = Decimal(1000 if pressure[2].lower() == "kpa" else 1000000)
            result["maximum_pressure"] = format((Decimal(pressure[1]) / divisor).normalize(), "f") + "MPa"
    flow = re.fullmatch(r"\s*(U\s*\d+)\s*[/／]\s*(D\s*\d+)\s*",
                        values.get("flow_sensitivity", ""), re.I)
    if flow:
        result.update(upstream=re.sub(r"\s+", "", flow[1]).upper(),
                      downstream=re.sub(r"\s+", "", flow[2]).upper())
    loss = re.fullmatch(r"\s*([△Δ])\s*p\s*(\d+(?:\.\d+)?)\s*",
                        values.get("pressure_loss", ""), re.I)
    if loss and Decimal(loss[2]) > 0:
        result["pressure_loss"] = loss[1] + "p" + loss[2]
        limit = Decimal(loss[2]) / 1000
        result["loss_limit"] = format(limit, ".3f") if limit == limit.quantize(Decimal(".001")) else format(limit, "f")
    return result


def fill_result_requirements(doc, name, parameters):
    name = d._project_name(name)
    if name not in PROJECT_FIELDS:
        return []
    cells = []
    for table in d._tables(doc):
        rows = d._rows(table)
        if not rows or not all(word in d._text(rows[0]) for word in ("技术要求", "样机编号", "试验结果")):
            continue
        for _physical, effective in list(d._effective_rows(table))[1:]:
            for cell in effective[2:-2]:
                if cell not in cells:
                    cells.append(cell)
    targets = {key: [] for key in PROJECT_FIELDS[name]}
    for cell in cells:
        raw = "".join(d._xp(cell, ".//w:t/text()"))
        for key in targets:
            if key == "loss_limit":
                pattern = re.compile(r"应\s*不\s*超\s*过\s*(\d+(?:\.\d+)?)\s*MPa", re.I)
            else:
                label, number = TARGETS[key]
                pattern = re.compile(_label_pattern([label]).pattern + "(" + number + ")", re.I)
            for match in pattern.finditer(raw):
                targets[key].append((cell, *match.span(match.lastindex)))
    for key, slots in targets.items():
        expected = 2 if name == "水压" else 1
        if len(slots) != expected:
            label = "压力损失限值" if key == "loss_limit" else TARGETS[key][0]
            raise ValueError(f"{name}素材/{label}填写位置缺失、重复或格式不支持，无法安全替换。")
    values = parameter_values(parameters)
    issues = []
    edits = []
    for key, slots in targets.items():
        if key not in values:
            label = "压力损失限值" if key == "loss_limit" else TARGETS[key][0]
            issues.append(f"{ISSUE_PREFIX}{name}/{label}：当前大纲缺失、冲突或值格式无法识别，保留素材原文，请核对。")
            continue
        edits.extend((cell, start, end, values[key]) for cell, start, end in slots)
    for cell in cells:
        raw = "".join(d._xp(cell, ".//w:t/text()"))
        for _cell, start, end, value in sorted(
                (edit for edit in edits if edit[0] is cell), key=lambda edit: edit[1], reverse=True):
            raw = raw[:start] + value + raw[end:]
        d._replace_result_field_text(cell, raw)
    return issues


def sync_project_requirements(root, project, parameters):
    if d._project_name(project["name"]) not in PROJECT_FIELDS:
        return project
    result = deepcopy(project)
    doc = d._load(d._asset(root, project["asset_id"], "result")["_path"])
    issues = fill_result_requirements(doc, project["name"], parameters)
    from water_results import project_schema

    schema = project_schema(doc, result)
    rows = d._validate_rows(schema, result.get("rows", []), project["name"])
    for source in schema:
        row = rows[source["id"]]
        if row.get("requirement") == source["requirement"]:
            continue
        row["requirement"] = source["requirement"]
        result["confirmed"] = False
        if row.get("auto_value") and row.get("value") == row["auto_value"]:
            row.update(value="", auto_value="", source="", matched_sample_id="")
            row["match_issue"] = "大纲技术要求已变化，旧自动结果已失效，请重新解析本次资料。"
        elif row.get("value"):
            row["match_issue"] = "大纲技术要求已变化，人工结果保留，请核对后重新保存。"
    result["notes"] = [note for note in result.get("notes", []) if not note.startswith(ISSUE_PREFIX)] + issues
    return result
