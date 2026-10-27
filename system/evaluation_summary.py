"""Derive report location and period from this task, never template history."""

from contracts import DEFAULT_EQUIPMENT_STORAGE_LOCATION
from validation import parse_date


def summarize_evaluation(task):
    values, issues = {}, []
    equipment = task.get("equipment", [])
    locations = []
    for index, item in enumerate(equipment, 1):
        location = item.get("storage_location", DEFAULT_EQUIPMENT_STORAGE_LOCATION).strip()
        if not location:
            issues.append(f"型式评价地点：本次设备{index}未填写存放地点，请补齐设备信息并刷新快照。")
        elif location not in locations:
            locations.append(location)
    if not equipment:
        issues.append("型式评价地点：尚未选择本次仪器设备。")
    if equipment and not issues:
        values["test_location"] = "；".join(locations) + "；"

    start = None
    try:
        start = parse_date(task.get("sample_delivery_date", ""))
    except ValueError as exc:
        issues.append(f"型式评价时间：请填写有效的送样日期（{exc}）。")
    projects = task.get("projects", [])
    ends = []
    for project in projects:
        try:
            ends.append(parse_date(project.get("end", "")))
        except ValueError as exc:
            issues.append(f"型式评价时间：{project.get('name', '未命名项目')}截止日期无法识别（{exc}）。")
    if not projects:
        issues.append("型式评价时间：尚未选择试验项目，无法确定结束日期。")
    if ends and len(ends) == len(projects):
        end = max(ends)
        values["evaluation_end"] = end.isoformat()
        if start:
            if end < start:
                issues.append("型式评价时间：项目最晚截止日期早于送样日期，请核对。")
            else:
                values["evaluation_period"] = f"从{start:%Y年%m月%d日}到{end:%Y年%m月%d日}"
    return {"values": values, "issues": issues}
