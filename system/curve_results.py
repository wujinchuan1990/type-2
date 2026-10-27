"""Refresh saved curve results from current records without replacing manual edits."""

from copy import deepcopy

import documents


def _manual_result(row):
    if row.get("manual_override") or "实测结果人工修改" in row.get("source", ""):
        return True
    value, automatic = row.get("value", ""), row.get("auto_value", "")
    return value != automatic if automatic else bool(value.strip())


def refresh_curve_results(project, parsed, include_conditions=False):
    if documents._project_name(project["name"]) != "示值误差曲线":
        return documents.apply_records(project, parsed) if include_conditions else project

    result = documents.apply_records(project, parsed) if include_conditions else deepcopy(project)
    result["rows"] = deepcopy(project["rows"])
    records = [
        record for data in parsed for record in data.get("records", [])
        if documents._project_name(record.get("project", "")) == "示值误差曲线"
    ]
    for row in result["rows"]:
        if _manual_result(row):
            if row.get("matched_sample_id") and row["matched_sample_id"] != row.get("sample_id"):
                row["match_issue"] = "样机对应已改变，保留的人工结果属于旧编号，请核对后重新保存"
            continue
        value, source, issue = documents._match_measurement("示值误差曲线", row, records)
        row.update(
            value=value, auto_value=value, source=source, match_issue=issue,
            matched_sample_id=row.get("sample_id", "") if value else "",
        )
    if result["rows"] != project["rows"]:
        result["confirmed"] = False
    return result
