import re
from decimal import Decimal


def environment_bounds(value, unit, key):
    text = re.sub(r"[\s（）()]", "", value)
    unit = re.sub(r"[\s（）()]", "", unit).upper()
    suffix = r"(?:℃|°C)" if key == "temperature" else r"(?:%RH|%)"
    found = re.search(suffix + r"$", text, re.I)
    if found:
        embedded = found[0].upper()
        if unit and unit not in ({"℃", "°C"} if key == "temperature" else {"%", "%RH"}):
            raise ValueError("单位不一致")
        unit = embedded
        text = text[:found.start()]
        text = re.sub(suffix, "", text, flags=re.I)
    if unit not in ({"℃", "°C"} if key == "temperature" else {"%", "%RH"}):
        raise ValueError("缺少单位或单位不支持")
    number = r"[+-]?(?:\d+(?:\.\d*)?|\.\d+)"
    match = re.fullmatch(r"(" + number + r")(?:[~～至—–-](" + number + r"))?", text)
    if not match:
        raise ValueError("不是明确的数值或最小～最大范围")
    low, high = Decimal(match[1]), Decimal(match[2] or match[1])
    if low > high or (key == "humidity" and (low < 0 or high > 100)):
        raise ValueError("范围上下限无效")
    return low, high


def aggregate_project_environment(projects):
    values, issues = {}, []
    if not projects:
        return {"values": values, "issues": ["未选择试验项目，无法汇总报告环境温湿度。"]}
    for key, title, unit, output_unit in (
        ("temperature", "环境温度", "℃", "℃"),
        ("humidity", "环境湿度", "%RH", "％"),
    ):
        bounds = []
        for index, project in enumerate(projects, 1):
            text = project.get(key, "").strip()
            try:
                if not text:
                    raise ValueError("未填写")
                text = text.replace("−", "-").replace("－", "-").replace("％", "%")
                bounds.append(environment_bounds(text, unit, key))
            except ValueError as exc:
                issues.append(f"第{index}项「{project['name']}」{title}无法汇总（{exc}）。")
        if len(bounds) == len(projects):
            low, high = min(value[0] for value in bounds), max(value[1] for value in bounds)
            low_text = format(low.normalize(), "f") if low else "0"
            high_text = format(high.normalize(), "f") if high else "0"
            values["environment_" + key] = f"从{low_text}{output_unit}到{high_text}{output_unit}；"
    return {"values": values, "issues": issues}
