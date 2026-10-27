"""User-supplied instrument classification table and conservative name lookup."""

import re


CLASSIFICATIONS = [
    {"code": "12180000", "name": "流量计", "reference_code": "90261000", "aliases": [], "exact": True},
    {"code": "12180500", "name": "差压式流量计", "reference_code": "90261000", "aliases": []},
    {"code": "12181000", "name": "速度式流量计", "reference_code": "90261000", "aliases": ["涡街流量计"]},
    {"code": "12181500", "name": "液体容积式流量计", "reference_code": "90261000", "aliases": []},
    {"code": "12182000", "name": "转子流量计", "reference_code": "90261000", "aliases": []},
    {"code": "12182500", "name": "靶式流量变送器", "reference_code": "90261000", "aliases": []},
    {"code": "12183000", "name": "临界流流量计", "reference_code": "90261000", "aliases": []},
    {"code": "12183500", "name": "质量流量计", "reference_code": "90261000", "aliases": []},
    {"code": "12184000", "name": "气体层流流量传感器", "reference_code": "90261000", "aliases": []},
    {"code": "12184500", "name": "气体容积式流量计（气体腰轮流量计）", "reference_code": "90261000",
     "aliases": ["气体容积式流量计", "气体腰轮流量计"]},
    {"code": "12185000", "name": "明渠堰槽流量计", "reference_code": "90261000", "aliases": []},
    {"code": "12200000", "name": "水表", "reference_code": "90282010", "aliases": [], "exact": True},
    {"code": "12200500", "name": "冷水表", "reference_code": "90282010", "aliases": ["冷水水表"]},
    {"code": "12201000", "name": "热水表", "reference_code": "90282010", "aliases": ["热水水表"]},
    {"code": "12220000", "name": "燃气表", "reference_code": "90281010", "aliases": [], "exact": True},
    {"code": "12220500", "name": "膜式燃气表（膜式煤气表）", "reference_code": "90281010",
     "aliases": ["膜式燃气表", "膜式煤气表"]},
    {"code": "12240000", "name": "热能表", "reference_code": "90281090", "aliases": []},
]
CLASSIFICATION_CODES = frozenset(item["code"] for item in CLASSIFICATIONS)
CLASSIFICATION_ISSUE = "分类编码未能按器具名称唯一匹配，请在基本信息总览选择；草稿显示“待补录”，不沿用模板旧编码。"


def match_classification(name):
    name = re.sub(r"\s+", "", name)
    matches = {
        item["code"] for item in CLASSIFICATIONS
        if any(name == alias if item.get("exact") else alias in name
               for alias in [item["name"], *item["aliases"]])
    }
    return next(iter(matches)) if len(matches) == 1 else ""


def resolve_classification(task):
    selected = task.get("classification_code", "")
    if selected:
        if selected not in CLASSIFICATION_CODES:
            raise ValueError("分类编码须从已录入的对照表中选择")
        return selected
    name = next((field.get("value", "") for field in task.get("overview", [])
                 if field["key"] == "instrument_name"), "")
    return match_classification(name)
