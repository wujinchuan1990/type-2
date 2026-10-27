"""Local report families and compatibility with existing liquid-flowmeter data."""

from typing import Annotated, Literal

from pydantic import BeforeValidator


def normalize_report_type(value):
    return "流量计（液体）" if value == "流量计" else value


ReportType = Annotated[
    Literal["水表", "流量计（气体）", "流量计（液体）", "燃气表"],
    BeforeValidator(normalize_report_type),
]
REPORT_TYPES = ("水表", "流量计（气体）", "流量计（液体）", "燃气表")
FLOWMETER_TYPES = ("流量计（气体）", "流量计（液体）", "流量计")
INDEPENDENT_METER_TYPES = (*FLOWMETER_TYPES, "燃气表")
REPORT_TYPE_FOLDERS = {
    "水表": "water meter",
    "流量计（气体）": "gas flowmeter",
    "流量计（液体）": "flowmeter",
    "燃气表": "gas meter",
}


def normalize_report_payload(value):
    if isinstance(value, list):
        return [normalize_report_payload(item) for item in value]
    if isinstance(value, dict):
        return {key: normalize_report_type(item) if key == "report_type"
                else normalize_report_payload(item) for key, item in value.items()}
    return value
