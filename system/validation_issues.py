import hashlib
import json
import re


class IssueMessages(list):
    def __init__(self, issues, severity):
        super().__init__()
        self.issues = issues
        self.severity = severity
        self.target = {"step": 0}

    def append(self, message):
        super().append(message)
        target = dict(self.target)
        identity = hashlib.sha256(json.dumps([self.severity, message, target],
                                             sort_keys=True, ensure_ascii=False).encode()).hexdigest()[:16]
        if not any(issue["id"] == identity for issue in self.issues):
            self.issues.append({"id": identity, "severity": self.severity,
                                "message": message, "target": target})

    def extend(self, messages):
        for message in messages:
            self.append(message)


def consistency_warnings(task, warnings, parse_date, component_text=""):
    notes = [*task.supplementary_notes, *[note.text for note in task.custom_supplementary_notes]]
    for note in notes:
        for old, new in re.findall(r"将\s*([A-Za-z][A-Za-z0-9._/-]{1,40})\s*(?:改为|更改为|变更为|更换为)\s*([A-Za-z][A-Za-z0-9._/-]{1,40})", note):
            tokens = set(re.findall(r"[A-Za-z][A-Za-z0-9._/-]{1,40}", component_text))
            if old in tokens and new not in tokens:
                warnings.target = {"step": 1}
                warnings.append(f"补充说明写明将{old}改为{new}，但本次大纲零部件表仍含{old}且未找到{new}。请核对原资料；未改写大纲或型号。")
    for field in task.overview:
        candidates = {candidate.value.strip() for candidate in field.candidates if candidate.value.strip()}
        if field.manual_override and field.value.strip() and candidates and field.value.strip() not in candidates:
            warnings.target = {"step": 1, "overview_key": field.key}
            warnings.append(f"总览「{field.label}」人工填写值与本次来源候选不同，请核对；未自动替换。")
    for project in task.projects:
        for row in project.rows:
            if row.manual_override and row.auto_value and row.value.strip() != row.auto_value.strip():
                warnings.target = {"step": 4, "project_id": project.id, "row_id": row.id}
                warnings.append(f"{project.name}：样机{row.sample_id}人工实测值与保留的自动值不同，请核对来源；人工值未覆盖。")
        warnings.target = {"step": 4, "project_id": project.id, "field": "end"}
        try:
            start, end = parse_date(project.start), parse_date(project.end)
        except ValueError:
            continue
        if start.year != end.year:
            warnings.append(f"{project.name}：试验起止日期跨年（{project.start} → {project.end}），请核对年份；跨年不等于错误。")
        if project.start == project.end and ":" in project.start:
            warnings.append(f"{project.name}：开始和结束时间完全相同，请核对是否为实际记录。")
        if task.sample_delivery_date:
            try:
                delivery = parse_date(task.sample_delivery_date)
            except ValueError:
                continue
            if start < delivery:
                warnings.target = {"step": 4, "project_id": project.id, "field": "start"}
                warnings.append(f"{project.name}：开始日期早于本次送样日期，请核对日期或送样信息。")
