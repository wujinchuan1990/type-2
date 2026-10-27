"""Separate the water outline's sample quantities from the actual task roster."""

from copy import deepcopy
import re

import documents as d


def _shape(schema):
    result = deepcopy(schema)
    for group in result.get("groups", []):
        group.pop("labels", None)
    return result


def _prefix(group):
    match = re.fullmatch(r"([A-Z]+)\d{3}", group["labels"][0])
    if match is None:
        raise ValueError("大纲样机组缺少有效的报告代号，不能猜测型号规格分组。")
    return match[1]


def _model(value):
    return d._norm(re.sub(r"DN\s*\d+", "", value, flags=re.I))


def _reconcile(task, schema, strict):
    from water_results import enabled

    projects = task.get("projects", [])
    task["projects"] = [project for project in projects if not enabled(project)]
    task["sample_schema"] = schema
    try:
        return d.reconcile_samples(task, schema, strict=strict)
    finally:
        task["projects"] = projects


def reconcile_samples(task, schema, parsed=None, strict=True, *, quantity_differences=None, roster_issues=None):
    current = deepcopy(task)
    migrating = not bool(current.get("water_sample_source_schema"))
    source = current.get("water_sample_source_schema") or current.get("sample_schema", {})
    samples = current.get("samples", [])
    if source and source != schema:
        if (any(sample.get("number", "").strip() or sample.get("disassembly") or sample.get("sealed")
                for sample in samples)
                or any(row.get("sample_id", "").strip() for project in current.get("projects", [])
                       for row in project["rows"])):
            raise ValueError("大纲样机表已改变。请先清空旧样机对应关系，再选择新大纲，避免编号对应错型号。")
        samples = []
        current.update(sample_schema={}, samples=[], water_sample_next_numbers={})
    if not schema:
        current.update(water_sample_source_schema={}, water_sample_next_numbers={},
                       sample_candidates={}, sample_mapping_issues=[])
        return _reconcile(current, {}, strict)
    if parsed is None and not current.get("water_sample_source_schema"):
        return _reconcile(current, schema, strict)
    active = current.get("sample_schema", {})
    if active:
        labels = [label for group in active.get("groups", []) for label in group["labels"]]
        if (_shape(active) != _shape(schema) or [sample["label"] for sample in samples] != labels
                or len(labels) != len(set(labels))):
            raise ValueError("样机清单与大纲型号规格分组不一致，请核对本次样机。")
        for original, group in zip(schema["groups"], active["groups"]):
            if any(not re.fullmatch(re.escape(_prefix(original)) + r"\d{3}", label)
                   or label.endswith("000") for label in group["labels"]):
                raise ValueError("样机报告代号与大纲分组不一致，不能跨型号规格添加。")
    elif samples:
        raise ValueError("样机清单缺少大纲分组，不能猜测编号所属型号规格。")
    current["water_sample_source_schema"] = deepcopy(schema)
    if parsed is None:
        return _reconcile(current, active, strict)

    filtered, extra_issues, identities = deepcopy(parsed), [], {}
    headers = schema["headers"]
    models = [value for group in schema["groups"] for header, value in zip(headers, group["values"])
              if "型号" in header]
    for data in filtered:
        retained = []
        for record in data.get("records", []):
            number, model = record.get("sample_id", "").strip(), record.get("model", "")
            if not number or (models and not any(_model(model) == _model(value) for value in models)):
                extra_issues.append("原始编号%s（%s）与大纲型号不一致或身份缺失，未自动分配样机。" % (number, model))
                continue
            identities.setdefault(number, set()).add((d._norm(model), record.get("certificate", "")))
            retained.append(record)
        data["records"] = retained
    if any(len(values) > 1 for values in identities.values()):
        extra_issues.append("同一原始编号存在不同型号规格或证书身份，不能自动合并样机。")
    candidates, issues = d.sample_record_candidates(schema, filtered, match_model=True)
    issues = list(dict.fromkeys(issues + extra_issues))
    scopes = {record.get("certificate") for data in filtered for record in data.get("records", [])
              if record.get("certificate")}
    blocked = len(scopes) > 1 or any(len(values) > 1 for values in identities.values())
    active = deepcopy(schema)
    next_numbers = current.get("water_sample_next_numbers", {}).copy()
    by_label = {sample["label"]: sample for sample in samples}
    old_candidates = current.get("sample_candidates", {})
    reserved = set()
    for sample in samples:
        sample["number"] = sample.get("number", "").strip()
        automatic = sample.get("number_origin") == "auto" and sample["number"] == sample.get("auto_number", "")
        if (sample.get("number_origin") == "manual" or sample.get("added_manually")
                or (sample["number"] and not automatic and sample["number"] != sample["label"])):
            if not sample.get("record_number") and automatic:
                sample["record_number"] = sample["number"]
            if not sample.get("record_number") and not sample.get("added_manually"):
                for group in schema["groups"]:
                    if re.fullmatch(re.escape(_prefix(group)) + r"\d{3}", sample["label"]):
                        choices = old_candidates.get(group["labels"][0], [])
                        index = int(re.search(r"\d{3}$", sample["label"])[0]) - 1
                        if index < len(choices):
                            sample["record_number"] = choices[index]["sample_id"]
            sample.update(number_origin="manual", auto_number="")
            reserved.update(value for value in (sample["number"], sample.get("record_number", "")) if value)
        elif sample.get("sealed") or sample.get("disassembly"):
            sample.update(number_origin="manual", auto_number="")
            reserved.add(sample["number"])
    updated, actual_candidates = [], {}
    for index, (original, group) in enumerate(zip(schema["groups"], active["groups"])):
        prefix = _prefix(original)
        old_groups = current.get("sample_schema", {}).get("groups", [])
        old_group = old_groups[index] if old_groups else None
        labels = old_group["labels"] if old_group else []
        choices = candidates.get(original["labels"][0], [])
        by_number = {choice["sample_id"]: choice for choice in choices}
        used = reserved | {by_label[label]["number"] for label in labels
                           if by_label[label]["number"] in by_number}
        available = iter(choice for choice in choices if choice["sample_id"] not in used)
        kept = []
        for label in labels:
            sample = by_label[label]
            if sample.get("number_origin") == "manual":
                kept.append(sample)
            elif not blocked and sample["number"] in by_number:
                choice = by_number[sample["number"]]
                sample.update(number_origin="auto", auto_number=sample["number"],
                              record_number=sample["number"], number_source=choice["source"])
                kept.append(sample)
            elif migrating and not blocked and not sample["number"]:
                choice = next(available, None)
                if choice:
                    sample.update(number=choice["sample_id"], number_origin="auto", auto_number=choice["sample_id"],
                                  record_number=choice["sample_id"], number_source=choice["source"])
                    kept.append(sample)
        used = reserved | {sample["number"] for sample in kept if sample["number"]}
        minimum = max([int(sample["label"][len(prefix):]) + 1 for sample in kept] or [1])
        counter = next_numbers.get(prefix, minimum)
        if type(counter) is not int or not 1 <= counter <= 1000:
            raise ValueError("样机报告代号计数无效，请重新核对样机清单。")
        counter = max(counter, minimum)
        if not blocked:
            for choice in choices:
                number = choice["sample_id"]
                if number in used:
                    continue
                if counter > 999:
                    raise ValueError("该型号规格组的样机报告代号已用尽，不能重用旧样机代号。")
                kept.append({"label": f"{prefix}{counter:03d}", "number": number, "number_origin": "auto",
                             "auto_number": number, "record_number": number, "number_source": choice["source"],
                             "disassembly": False, "sealed": False, "added_manually": False})
                counter += 1
                used.add(number)
        next_numbers[prefix] = counter
        group["labels"] = [sample["label"] for sample in kept]
        updated.extend(kept)
        actual_candidates.update({sample["label"]: choices for sample in kept})
        if len(kept) != len(original["labels"]):
            message = "规格组%s：本次样机%d台，大纲原数量%d台；按本次清单输出，不预留空白样机。" % (
                prefix, len(kept), len(original["labels"]))
            issues.append(message)
            if quantity_differences is not None:
                quantity_differences.append({
                    "group_index": index, "outline_count": len(original["labels"]),
                    "actual_count": len(kept), "issue_index": issues.index(message),
                })
        if not kept:
            issues.append("规格组%s未识别到可唯一对应的样机，可在该组手工添加并填写编号。" % prefix)
            if roster_issues is not None:
                roster_issues.append({"kind": "empty_group", "group_index": index, "issue_index": len(issues) - 1})
        for sample in kept:
            if not sample["number"]:
                issues.append("%s编号待手工填写，主动清空不会被自动覆盖。" % sample["label"])
                if roster_issues is not None:
                    roster_issues.append({"kind": "missing_number", "label": sample["label"], "issue_index": len(issues) - 1})
            elif sample.get("number_origin") == "manual" and sample["number"] not in by_number:
                suffix = "已选拆样，不计入参试样机。" if sample.get("disassembly") else "项目实测结果待补录。"
                issues.append("%s人工编号%s没有对应原始记录，%s" % (sample["label"], sample["number"], suffix))
                if roster_issues is not None:
                    roster_issues.append({"kind": "manual_without_record", "label": sample["label"], "issue_index": len(issues) - 1})
    if updated != samples:
        current["samples_confirmed"] = False
    current.update(samples=updated, sample_candidates=actual_candidates,
                   sample_mapping_issues=list(dict.fromkeys(issues)), water_sample_next_numbers=next_numbers)
    return _reconcile(current, active, strict)
