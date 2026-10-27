"""Single-model water materials describe conditions, not the sample count."""

from copy import deepcopy
import re

from docx.oxml.ns import qn

import documents as d


PER_GROUP = frozenset(("反向流", "水温", "水压", "静磁场", "安装条件", "压力损失"))
ISSUE_PREFIX = "水表实测明细/"


def selectable(project):
    return project.get("variant", "").strip().lower() in ("1 model", "b级", "o级", "e1级", "e2级")


def enabled(project):
    return (project.get("variant", "").strip().lower() == "1 model"
            or (selectable(project) and project.get("water_output_labels") is not None))


def reconcile_task(task, schema, strict=True, parsed=None):
    from water_samples import reconcile_samples

    return reconcile_samples(task, schema, parsed=parsed, strict=strict)


def _blocks(doc, *, dynamic=False):
    blocks, current, seen, previous = [], [], set(), None
    for row in d._result_schema(doc):
        physical = row["_sample_cell"].getparent()
        key = (physical.getparent(), row["requirement"])
        if current and (key != previous or row["sample_label"] in seen):
            blocks.append(current)
            current, seen = [], set()
        current.append(row)
        seen.add(row["sample_label"])
        previous = key
    if current:
        blocks.append(current)
    for block in blocks:
        pattern = r"[A-Z]+\d{3}" if dynamic else r"A\d{3}"
        if any(not re.fullmatch(pattern, row["sample_label"]) for row in block):
            raise ValueError("单型号水表结果素材必须使用明确的A组样机代号，不能猜测真实编号。")
        for row in block:
            if any(d._xp(row[key], "./w:tcPr/w:vMerge") for key in ("_sample_cell", "_value_cell")):
                raise ValueError("单型号水表结果素材的样机/结果列含纵向合并，无法安全调整样机数量。")
    return blocks


def schema_for(doc, labels, *, dynamic=False):
    if (not labels and not dynamic) or len(set(labels)) != len(labels):
        raise ValueError("水表实测明细参试代号缺失或重复。")
    source_blocks = _blocks(doc, dynamic=dynamic)
    bookmark_ids = set()
    bookmark_names = {start.get(qn("w:id")): start.get(qn("w:name"), "")
                      for start in d._xp(doc.element, ".//w:bookmarkStart")}
    for block in source_blocks:
        for row in block:
            for mark in d._xp(row["_sample_cell"].getparent(), ".//w:bookmarkStart | .//w:bookmarkEnd"):
                identity = mark.get(qn("w:id"))
                name = bookmark_names.get(identity, "")
                if (not name.startswith(("_Hlk", "_GoBack", "OLE_LINK"))
                        or any(name in instruction for instruction in d._xp(doc.element, ".//w:instrText/text()"))):
                    raise ValueError("样机数据行含自定义书签，不能自动增减行。")
                bookmark_ids.add(identity)
    for mark in d._xp(doc.element, ".//w:bookmarkStart | .//w:bookmarkEnd"):
        if mark.get(qn("w:id")) in bookmark_ids:
            mark.getparent().remove(mark)
    for block in source_blocks:
        first = block[0]["_sample_cell"].getparent()
        source_rows = [row["_sample_cell"].getparent() for row in block]
        table = first.getparent()
        insertion = table.index(first)
        owners = list(d._effective_rows(table))
        effective = {tuple(physical): cells for physical, cells in owners}
        first_effective = effective[tuple(d._cells(first))]
        for offset, label in enumerate(labels or [""]):
            original = next((row for row in block if row["sample_label"] == label), None)
            source = original or block[min(offset, len(block) - 1)]
            physical = source["_sample_cell"].getparent()
            copied = deepcopy(physical)
            cells = d._cells(copied)
            if d._xp(copied, ".//w:bookmarkStart | .//w:bookmarkEnd | .//w:drawing | .//w:sectPr"):
                raise ValueError("样机数据行含书签、图片或节结构，不能复制为另一台样机。")
            source_effective = effective[tuple(d._cells(physical))]
            for column, cell in enumerate(cells[:-2]):
                merge = cell.find("./" + qn("w:tcPr") + "/" + qn("w:vMerge"))
                if merge is None:
                    if len(labels) <= 1:
                        continue
                    merge = cell.get_or_add_tcPr().get_or_add_vMerge()
                if offset == 0:
                    # Retain a continuation only when its owner lies before this block.
                    if first_effective[column] is d._cells(first)[column]:
                        merge.set(qn("w:val"), "restart")
                        d._put(cell, d._text(first_effective[column]))
                    else:
                        merge.attrib.pop(qn("w:val"), None)
                        d._put(cell, "")
                else:
                    if d._text(source_effective[column]) != d._text(first_effective[column]):
                        raise ValueError("同一试验条件内的技术要求合并关系不一致，不能安全增减行。")
                    merge.attrib.pop(qn("w:val"), None)
                    d._put(cell, "")
            d._put(cells[-2], label)
            if dynamic or original is None:
                d._put(cells[-1], "")
            table.insert(insertion + offset, copied)
        for row in source_rows:
            table.remove(row)
    result = d._result_schema(doc, allow_empty_samples=not labels)
    if len(result) != len(source_blocks) * len(labels):
        raise ValueError("水表实测明细调整后的行数不完整。")
    for index, row in enumerate(result):
        row["id"] = f"water:{index // len(labels)}:{row['sample_label']}"
    return result


def output_rows(project):
    selected = project.get("water_output_labels")
    return [row for row in project["rows"] if selected is None or row["sample_label"] in selected]


def validate_selection(project, task):
    selected = project.get("water_output_labels")
    if selected is None:
        return
    if task.get("report_type", "水表") != "水表" or not enabled(project):
        raise ValueError("样机勾选目前仅适用于水表单型号结果素材及B/O/E1/E2等级素材，不支持2 models。")
    eligible = {sample["label"] for sample in task["samples"] if not sample.get("disassembly")}
    invalid = set(selected) - eligible
    if invalid:
        raise ValueError(project["name"] + "：勾选的样机已不在本次参试清单或已选拆样，请取消勾选：" +
                         "、".join(sorted(invalid)))


def output_schema(doc, project):
    selected = project.get("water_output_labels")
    if selected is None:
        return d._result_schema(doc)
    labels = project.get("water_result_labels", [])
    invalid = set(selected) - set(labels)
    if invalid:
        raise ValueError(project["name"] + "：勾选样机尚无可映射结果行，请保存或重新解析：" +
                         "、".join(sorted(invalid)))
    return schema_for(doc, [label for label in labels if label in selected], dynamic=True)


def project_schema(doc, project):
    labels = project.get("water_result_labels", [])
    return schema_for(doc, labels) if labels else d._result_schema(doc)


def _manual(row):
    return bool(row.get("manual_override") or "实测结果人工修改" in row.get("source", "")
                or (d._has_text(row.get("value")) and row.get("value") != row.get("auto_value")))


def sync_project(root, project, task, parsed, parameters, refresh_conditions=False):
    if not enabled(project) or not task.get("sample_schema"):
        return project
    result = deepcopy(project)
    doc = d._load(d._asset(root, project["asset_id"], "result")["_path"])
    from result_requirements import fill_result_requirements

    fill_result_requirements(doc, project["name"], parameters)
    blocks = _blocks(doc)
    records = [record for data in parsed for record in data.get("records", [])
               if d._project_name(record.get("project", "")) == d._project_name(project["name"])]
    samples = [sample for sample in task["samples"] if not sample.get("disassembly")]
    by_label = {sample["label"]: sample for sample in samples}
    known = {sample["number"] for sample in samples if sample["number"]}
    issues = [ISSUE_PREFIX + "原始记录样机不在本次参试清单中：" + number
              for number in dict.fromkeys(record.get("sample_id", "") for record in records)
              if number and number not in known]
    all_samples = any(len(block) > 1 for block in blocks)
    observed = {record.get("sample_id") for record in records}
    if all_samples:
        chosen = samples
    elif d._project_name(project["name"]) in PER_GROUP:
        chosen = []
        for group in task["sample_schema"]["groups"]:
            eligible = [by_label[label] for label in group["labels"] if label in by_label]
            found = [sample for sample in eligible if sample["number"] and sample["number"] in observed]
            chosen.extend(found or eligible[:1])
    else:
        chosen = [sample for sample in samples if sample["number"] and sample["number"] in observed]
        chosen = chosen or samples[:1]
    if project.get("water_output_labels") is not None:
        # Keep unselected results in the task so toggling never discards manual data.
        chosen = samples
    labels = [sample["label"] for sample in chosen]
    if not labels:
        if any(_manual(row) for row in project["rows"]):
            raise ValueError(project["name"] + "：没有参试样机，不能移除已手填或主动清空的结果；未删除人工内容。")
        result["water_result_labels"] = []
        result["rows"] = [{**{key: value for key, value in row.items() if not key.startswith("_")},
                           "sample_id": "", "value": "", "auto_value": "", "matched_sample_id": "",
                           "source": "", "match_issue": "尚未识别参试样机，请在样机清单添加或选择有效记录"}
                          for row in d._result_schema(doc)]
        result["notes"] = [note for note in result.get("notes", []) if not note.startswith(ISSUE_PREFIX)]
        result["notes"].append(ISSUE_PREFIX + "尚无参试样机，结果素材代号仅为待补录位置，不代表本次识别的样机。")
        result["confirmed"] = False
        return result
    source_ids = {row["id"]: (index, row["sample_label"])
                  for index, block in enumerate(blocks) for row in block}
    old = {}
    for row in project["rows"]:
        if row["id"].startswith("water:"):
            match = re.fullmatch(r"water:(\d+):([A-Z]+\d{3})", row["id"])
            if (match is None or int(match[1]) >= len(blocks)
                    or match[2] != row["sample_label"]):
                raise ValueError(project["name"] + "：动态结果行与试验条件或样机代号不一致。")
            key = row["id"]
        elif row["id"] in source_ids:
            index, label = source_ids[row["id"]]
            key = f"water:{index}:{label}"
        else:
            raise ValueError(project["name"] + "：旧结果行与单型号素材不一致，不能丢弃未映射数据。")
        if key in old:
            raise ValueError(project["name"] + "：实测结果行标识重复。")
        old[key] = row
    result["water_result_labels"] = labels
    schema = schema_for(doc, labels)
    expected = {row["id"] for row in schema}
    if any(_manual(row) for key, row in old.items() if key not in expected):
        raise ValueError(project["name"] + "：参试清单变化将移除已手填或主动清空的结果，请先核对相关样机；未删除人工内容。")
    mapping = {sample["label"]: sample["number"] for sample in chosen}
    updated = []
    for source in schema:
        value = deepcopy(old.get(source["id"], {}))
        value.update(id=source["id"], sample_label=source["sample_label"],
                     requirement=source["requirement"])
        number = mapping[source["sample_label"]]
        old_number = value.get("matched_sample_id") or value.get("sample_id", "")
        if _manual(value):
            if old_number and old_number != number:
                value["matched_sample_id"] = old_number
                value["match_issue"] = "样机对应已改变，保留的人工结果属于旧编号，请核对后重新保存"
        else:
            measured, location, issue = d._match_measurement(d._project_name(project["name"]),
                {**value, "sample_id": number}, records)
            value.update(value=measured, auto_value=measured, source=location, match_issue=issue,
                         matched_sample_id=number if measured else "", manual_override=False)
        value["sample_id"] = number
        updated.append(value)
    result["rows"] = updated
    result["notes"] = [note for note in result.get("notes", []) if not note.startswith(ISSUE_PREFIX)] + issues
    if refresh_conditions:
        summary_project = deepcopy(result)
        summary_project["rows"] = []
        summarized = d.apply_records(summary_project, parsed)
        for key in ("start", "end", "temperature", "humidity", "source", "record_summary"):
            result[key] = summarized[key]
    if result["rows"] != project["rows"]:
        result["confirmed"] = False
    return result
