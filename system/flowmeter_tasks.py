"""Flowmeter-only sample reconciliation for code and real-number result templates."""

import re
from copy import deepcopy

import documents


def _model(value):
    return re.sub(r"\s+", "", re.sub(r"DN\s*\d+", "", value, flags=re.I))


def reconcile_samples(task, schema, parsed=None, strict=True, report_type="流量计（液体）", *, match_model=False, roster_issues=None):
    current = deepcopy(task)
    projects = current.pop("projects", [])
    if current.get("sample_schema") and current["sample_schema"] != schema and any(
            row.get("sample_id", "").strip() for project in projects for row in project["rows"]):
        raise ValueError("大纲样机表已改变，项目中仍有旧样机编号。请先清空样机对应关系。")
    current = documents.reconcile_samples(current, schema, strict=strict)
    extra_issues = []
    if schema and parsed is not None:
        candidate_schema = deepcopy(schema)
        candidate_schema["headers"] = [
            "规格" if documents._norm(header) == "口径" else header for header in schema["headers"]
        ]
        models = {
            _model(value)
            for group in schema["groups"]
            for header, value in zip(schema["headers"], group["values"])
            if "型号" in header and value.strip()
        }
        candidates = deepcopy(parsed)
        if models:
            for data in candidates:
                retained = []
                for record in data.get("records", []):
                    if _model(record.get("model", "")) in models:
                        retained.append(record)
                    else:
                        extra_issues.append(
                            "%s原始编号%s（%s）与大纲型号不一致，未自动分配样机编号。"
                            % (report_type, record["sample_id"], record.get("model", ""))
                        )
                data["records"] = retained
        current["sample_schema"] = candidate_schema
        current = documents.fill_sample_numbers(current, candidates, match_model=match_model, roster_issues=roster_issues)
        current["sample_schema"] = schema
        current = documents.reconcile_samples(current, schema, strict=strict)
    elif not schema:
        current["sample_candidates"] = {}
        current["sample_mapping_issues"] = []

    current["sample_mapping_issues"] = list(dict.fromkeys(
        current.get("sample_mapping_issues", []) + extra_issues))
    mappings = {sample["label"]: sample.get("number", "") for sample in current["samples"]}
    known = {value for value in mappings.values() if value} if schema else {
        record["sample_id"] for data in (parsed or []) for record in data.get("records", [])
        if record.get("sample_id")
    }
    issues = list(current.get("sample_issues", []))
    for project in projects:
        for row in project["rows"]:
            label = row.get("sample_label", "").strip()
            number = mappings.get(label, label if label in known else "")
            previous = row.get("sample_id", "")
            if previous != number:
                project["confirmed"] = False
                if not row.get("manual_override") and row.get("auto_value") and row.get("value") == row["auto_value"]:
                    row.update(value="", source="", auto_value="", matched_sample_id="")
                elif row.get("value", "").strip() or row.get("manual_override"):
                    row["matched_sample_id"] = row.get("matched_sample_id") or previous
            row["sample_id"] = number
            if not number and label not in mappings:
                message = "%s结果素材编号%s未在本次样机中找到；不按旧编号或表格行序猜配。" % (report_type, label)
                row["match_issue"] = message
                issues.append(message)
    current.update(projects=projects, sample_issues=list(dict.fromkeys(issues)))
    return current
