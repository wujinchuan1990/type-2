"""Classify legacy upload listings without rewriting uploads or task snapshots."""


from report_types import REPORT_TYPES, normalize_report_type


def upload_report_type(upload):
    explicit = normalize_report_type(upload.get("report_type"))
    if explicit in REPORT_TYPES:
        return explicit
    names = [candidate.get("value", "") for candidate in
             upload.get("fields", {}).get("instrument_name", [])]
    matches = {kind for name in names for kind in REPORT_TYPES if kind in name}
    if any("流量计" in name and not any(kind in name for kind in
           ("流量计（气体）", "流量计（液体）")) for name in names):
        matches.add("流量计（液体）")
    return next(iter(matches)) if len(matches) == 1 else ""


def classified_upload(upload):
    result = dict(upload)
    result["report_type"] = upload_report_type(upload)
    result["classification_issue"] = (
        "" if result["report_type"] else
        "旧资料没有明确的报告对象类型，器具名称缺失或存在歧义；请在对应报告类型下重新上传，原件未删除。"
    )
    return result
