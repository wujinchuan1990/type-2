import re
from datetime import date, datetime

from contracts import (AUTOMATIC_OVERVIEW_FIELDS, MEASUREMENT_FIELDS, OBSERVATION_MANUAL_FIELDS, REPORT_ENVIRONMENT_FIELDS,
                       Task, TEMPLATE_FALLBACK_WARNING, active_photo_slots, photo_slot_group)
from environment import aggregate_project_environment
from classification import CLASSIFICATION_ISSUE, resolve_classification
from validation_issues import IssueMessages, consistency_warnings

ORDINARY = {"吴锦川", "姚新红"}
GAS_STAFF = {"陈超", "沈琼", "朱碧玉", "刘夷平"}
CLIMATE = {"张宇", "黄薛青"}
ELECTROMAGNETIC = {"刘雨宸", "秦夏臻", "杨先望"}
STAFF_GROUPS = {
    "液体组": ORDINARY,
    "气体组": GAS_STAFF,
    "环境试验组": CLIMATE,
    "电磁兼容组": ELECTROMAGNETIC,
}


def staff_group(name):
    if any(word in name for word in ("贮存环境性能", "高温", "低温", "湿热")):
        return CLIMATE
    if any(word in name for word in ("电磁兼容", "静电", "辐射")):
        return ELECTROMAGNETIC
    return set()


def required_staff_groups(task, project):
    report_type = task.get("report_type", "水表") if isinstance(task, dict) else task.report_type
    equipment = task.get("equipment", []) if isinstance(task, dict) else task.equipment
    equipment_ids = project.get("equipment_ids", []) if isinstance(project, dict) else project.equipment_ids
    required = set()
    if report_type == "水表":
        required.add("液体组")
    elif report_type == "燃气表":
        required.add("气体组")
    for item in equipment:
        identity = item.get("id", "") if isinstance(item, dict) else item.id
        group = item.get("trial_group", "") if isinstance(item, dict) else item.trial_group
        if identity in equipment_ids and group in STAFF_GROUPS:
            required.add(group)
    return [(name, STAFF_GROUPS[name]) for name in STAFF_GROUPS if name in required]


def parse_date(value):
    text = value.strip().replace("年", "-").replace("月", "-").replace("日", "")
    text = text.replace("/", "-").replace(".", "-")
    match = re.fullmatch(
        r"(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?", text
    )
    if not match:
        raise ValueError("日期须为 YYYY-MM-DD 或完整日期时间")
    year, month, day, hour, minute, second = match.groups()
    if hour is not None:
        datetime(int(year), int(month), int(day), int(hour), int(minute), int(second or 0))
    return date(int(year), int(month), int(day))


def validate_task(task: Task, component_text=""):
    from evaluation_summary import summarize_evaluation
    from observation_documents import copy_warning
    from water_results import output_rows

    issues = []
    errors, warnings = IssueMessages(issues, "error"), IssueMessages(issues, "warning")
    warnings.target = {"step": 6}
    warnings.append(TEMPLATE_FALLBACK_WARNING)
    errors.target = {"step": 1}
    if not resolve_classification(task.model_dump()):
        errors.append(CLASSIFICATION_ISSUE)
    errors.extend(summarize_evaluation(task.model_dump())["issues"])
    errors.target = {"step": 5}
    errors.extend(task.photo_issues)
    errors.target = {"step": 1}
    errors.extend(task.construction_issues)
    errors.extend(task.production_basis_issues)
    errors.extend(task.test_basis_issues)
    errors.extend(task.sealing_issues)
    errors.target = {"step": 1, "overview_group": "observation"}
    errors.extend(task.observation_certificates.issues)
    for key, certificate in (("explosion", task.observation_certificates.explosion),
                             ("protection", task.observation_certificates.protection)):
        errors.target = {"step": 1, "certificate_key": key}
        errors.extend(certificate.issues)
    if task.report_type == "水表":
        errors.target = {"step": 1, "overview_group": "observation"}
        errors.extend(task.observation_summary.issues)
    errors.target = {"step": 1}
    errors.extend(task.supplementary_issues)
    errors.target = {"step": 4}
    errors.extend(aggregate_project_environment([project.model_dump() for project in task.projects])["issues"])
    errors.target = {"step": 1}
    for index, parameter in enumerate(task.measurement_parameters or [], 1):
        for key, label in MEASUREMENT_FIELDS:
            if not getattr(parameter, key).strip():
                errors.append(f"计量器具测量参数第{index}行「{label}」尚未填写；原有行空值保留模板原文。")
        if task.include_volume_display and not parameter.volume_display.strip():
            errors.append(f"计量器具测量参数第{index}行「{task.volume_display_label}」尚未填写；新增列空值保持空白。")
    if task.include_volume_display and task.measurement_parameters is None:
        errors.append(f"计量器具测量参数「{task.volume_display_label}」尚未填写；新增列空值保持空白。")
    for slot, (label, _width) in active_photo_slots(task.display_types, task.report_type).items():
        errors.target = {"step": 5, "photo_slot": slot}
        if slot not in {photo_slot_group(key) for key in task.photos}:
            errors.append("附件2缺少" + label + "；草稿保留空白照片位置。")
    for component in task.photo_components:
        errors.target = {"step": 5, "photo_slot": component["id"]}
        if component["id"] not in task.photos:
            errors.append("附件2缺少零部件照片：" + component["name"])
    errors.target = {"step": 2}
    if not task.items_asset_id:
        errors.append("请选择型式评价项目一览表方案。")
    if not task.observe_asset_id:
        errors.append("请选择观察项目方案。")
    errors.target = {"step": 0}
    if not task.outline_id:
        errors.append("请上传并选择本次大纲。")
    errors.target = {"step": 2}
    if not task.projects:
        errors.append("请至少选择一个试验项目。")
    errors.target = {"step": 1}
    if not task.samples or any(not s.number.strip() for s in task.samples):
        errors.append("样机清单尚未逐台填写编号；草稿未填写编号处使用对应报告代号。")
    if task.samples and not any(s.disassembly for s in task.samples):
        errors.append("请在样机清单中选择拆样样机。")
    if task.legacy_sample_fields:
        errors.append("旧任务的样机字段尚未核对，请在样机清单下核对并处理旧值。")
    errors.extend(task.sample_issues)
    if not task.overview:
        errors.append("请填写基本信息总览。")
    for field in task.overview:
        errors.target = warnings.target = {"step": 1, "overview_key": field.key}
        if field.key in REPORT_ENVIRONMENT_FIELDS | AUTOMATIC_OVERVIEW_FIELDS:
            continue
        if task.report_type != "水表" and field.key in OBSERVATION_MANUAL_FIELDS:
            continue
        if not field.value.strip():
            errors.append(f"总览「{field.label}」尚未填写；不适用项请明确填写不适用。")
        if len({candidate.value for candidate in field.candidates}) > 1:
            warnings.append(f"总览「{field.label}」存在不同来源候选，请确认所选值。")
        if (field.key == "test_basis" and field.value.strip() and field.candidates
                and not any(candidate.value == field.value for candidate in field.candidates)):
            warnings.append("试验依据与本次大纲第2项技术依据不同，已保留人工填写值，请核对。")
    errors.target = {"step": 3}
    if not task.equipment:
        errors.append("请加入本次报告使用的仪器设备。")
    equipment = {item.id: item for item in task.equipment}
    if len(equipment) != len(task.equipment) or "" in equipment:
        errors.append("本次设备清单存在重复或未保存的设备。")
    for item in task.equipment:
        if not all(value.strip() for value in (
                item.name, item.model, item.performance, item.number, item.valid_until, item.trial_group)):
            errors.append(f"设备「{item.name}」信息不完整。")
        try:
            parse_date(item.valid_until)
        except ValueError:
            errors.append(f"设备「{item.name}」证书有效期不是有效日期。")
    seen = set()
    sample_mappings = {}
    for number, project in enumerate(task.projects, 1):
        errors.target = warnings.target = {"step": 4, "project_id": project.id}
        label = f"{number}. {project.name} / {project.variant}"
        if task.include_atmospheric_pressure:
            warnings.extend(f"{label}：{issue}" for issue in project.pressure_issues)
            if not project.atmospheric_pressure.strip():
                errors.append(f"{label} 大气压力未填写或未识别，请在项目列表补录。")
        if project.id in seen:
            errors.append(f"{label} 项目实例标识重复。")
        seen.add(project.id)
        for key, title in (
            ("start", "试验开始时间"), ("end", "试验结束时间"),
            ("temperature", "环境温度"), ("humidity", "环境湿度"),
        ):
            errors.target = {"step": 4, "project_id": project.id, "field": key}
            if not getattr(project, key).strip():
                errors.append(f"{label} 缺少{title}。")
        errors.target = {"step": 4, "project_id": project.id, "field": "end"}
        dates = []
        for key, value in (("start", project.start), ("end", project.end)):
            errors.target = {"step": 4, "project_id": project.id, "field": key}
            try:
                dates.append(parse_date(value))
            except ValueError:
                if value:
                    errors.append(f"{label} 试验日期格式错误。")
        if len(dates) == 2:
            if dates[0] > dates[1]:
                errors.append(f"{label} 开始日期晚于结束日期。")
            try:
                start = datetime.fromisoformat(project.start)
                end = datetime.fromisoformat(project.end)
                if start > end and dates[0] == dates[1]:
                    errors.append(f"{label} 开始时间晚于结束时间。")
            except ValueError:
                pass  # Date-only and Chinese date formats were validated above.
        groups = required_staff_groups(task, project)
        for people, role in ((project.evaluators, "评价"), (project.reviewers, "复核")):
            errors.target = {"step": 4, "project_id": project.id}
            selected = set(people)
            if (len(people) != len(groups) or len(selected) != len(groups)
                    or any(len(selected & names) != 1 for _name, names in groups)
                    or selected - set().union(*(names for _name, names in groups))):
                expected = " + ".join(name + "1名" for name, _names in groups) or "按项目所选设备确定人员组"
                errors.append(f"{label} {role}人员须配置：{expected}。")
        if set(project.evaluators) & set(project.reviewers):
            errors.append(f"{label} 评价人员和复核人员不能是同一人。")
        if not project.rows:
            errors.append(f"{label} 未找到可填写结果行，需检查素材。")
        row_ids = [row.id for row in project.rows]
        if len(row_ids) != len(set(row_ids)):
            errors.append(f"{label} 结果行标识重复。")
        selected_ids = {row["id"] for row in output_rows(project.model_dump())}
        for index, row in enumerate(project.rows, 1):
            if row.id not in selected_ids:
                continue
            errors.target = warnings.target = {"step": 4, "project_id": project.id, "row_id": row.id}
            if row.match_issue:
                warnings.append(f"{label} 第{index}行：{row.match_issue}")
                if "属于旧编号" in row.match_issue:
                    errors.append(f"{label} 第{index}行：{row.match_issue}")
            if not row.sample_id.strip() or not row.value.strip():
                errors.append(f"{label} 第{index}行缺少样机编号或试验结果。")
            if row.sample_label and row.sample_id:
                old = sample_mappings.setdefault(row.sample_label, row.sample_id)
                if old != row.sample_id:
                    errors.append(f"样机代号 {row.sample_label} 在不同结果行中对应了不同实际编号。")
        errors.target = warnings.target = {"step": 4, "project_id": project.id}
        if not project.equipment_ids:
            errors.append(f"{label} 未关联本项目的标准器。")
        for identity in project.equipment_ids:
            if identity not in equipment:
                errors.append(f"{label} 引用了不在本次报告清单中的设备。")
                continue
            if dates:
                try:
                    if parse_date(equipment[identity].valid_until) < max(dates):
                        errors.append(f"{label} 设备「{equipment[identity].name}」证书在试验时已过有效期。")
                except ValueError:
                    pass  # Equipment validation already reports the invalid date.
        warnings.extend(f"{label}：{note}" for note in project.notes)
        warnings.extend(f"{label}：{issue}" for issue in project.record_summary.issues)
    warnings.target = {"step": 1}
    warnings.extend(task.sample_mapping_issues)
    consistency_warnings(task, warnings, parse_date, component_text)
    warnings.target = {"step": 6}
    warnings.append("观察记录仅更新指定水表备注，其他内容及已有勾选未由系统核验；项目一览表仍原样复制，请人工核对。"
                    if task.report_type == "水表" else copy_warning(task.report_type))
    warnings.append("姓名为人员配置展示，不代表签字或批准；设备有效期只校核截止日期。")
    warnings.append("各试验项目的结论及异常情况固定保留所选素材原文，不代表已核实本次试验。")
    return {"errors": list(dict.fromkeys(errors)), "warnings": list(dict.fromkeys(warnings)), "issues": issues}
