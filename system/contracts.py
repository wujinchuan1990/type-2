"""Shared JSON shapes for the local report workbench (Python 3.9+)."""

import re
from typing import Any, Dict, List, Literal, Optional, Tuple

from pydantic import BaseModel, ConfigDict, Field, StrictInt, field_validator, model_validator
from classification import CLASSIFICATION_CODES
from equipment_typography import equipment_identity
from report_types import ReportType, REPORT_TYPE_FOLDERS


OUTLINE_OVERVIEW_FIELDS = frozenset(("instrument_name", "manufacturer", "application_number", "production_basis", "test_basis"))
OBSERVATION_MANUAL_FIELDS = {
    "meter_manufacture_month": "制造年月",
    "battery_replacement_deadline": "更换电池的最后期限",
}
REPORT_ENVIRONMENT_FIELDS = frozenset(("environment_temperature", "environment_humidity"))
AUTOMATIC_OVERVIEW_FIELDS = frozenset(("test_location", "evaluation_period"))
TEMPLATE_DEFAULT_FIELDS = frozenset((
    "commissioned_unit", "commissioned_contact",
    "measurement_parameter", "parameter_unit", "measurement_range", "display_digits",
    "metrological_performance", "display_type", "other_environment",
    "technical_review", "evaluation_conclusion",
))
SAMPLE_OVERVIEW_FIELDS = frozenset(("model", "q3", "q3_q1", "accuracy_class",
                                  "sample_number", "disassembly_sample"))
HIDDEN_OVERVIEW_FIELDS = TEMPLATE_DEFAULT_FIELDS | SAMPLE_OVERVIEW_FIELDS | {
    "applicant", "classification_code", "installation",
}
TEMPLATE_FALLBACK_WARNING = (
    "未填写的字段保留原模板/素材内容，包括原有结果、日期、人员及结论；"
    "这些内容未确认为本次数据，设备序号及样机编号也须与本次清单核对。"
)
PHOTO_SLOTS = {
    "sample": ("样机照片", 12),
    "nameplate": ("铭牌照片", 8),
    "sealed_overall": ("整体封样照片", 12),
    "sealed_detail": ("封样特写照片", 8),
    "internal": ("样机内部结构照片", 8),
    "display": ("样机显示部分照片", 8),
    "display_mechanical": ("机械部分照片", 8),
    "display_electronic": ("电子部分照片", 8),
}
MULTI_PHOTO_SLOTS = frozenset(("sample", "display", "display_mechanical", "display_electronic", "sealed_overall"))
GAS_UNUSED_PHOTO_SLOTS = frozenset(("nameplate", "sealed_detail"))
GAS_PHOTO_REPORT_TYPES = frozenset(("燃气表", "流量计（气体）"))
PHOTO_HEADINGS = {
    "heading:overall": "样机的整体外形照片",
    "heading:internal": "样机的内部结构照片",
    "heading:display": "样机的显示部分照片",
    "heading:components": "样机的关键零部件照片",
}


def photo_slot_group(slot):
    base, separator, suffix = slot.partition(":")
    if base in MULTI_PHOTO_SLOTS and separator and re.fullmatch(r"[a-f0-9]{32}", suffix):
        return base
    return slot


def ordered_photo_slots(photos, order, group):
    keys = list(dict.fromkeys([*(order or []), *photos]))
    return [key for key in keys if key in photos and photo_slot_group(key) == group]

DISPLAY_TYPES = ("机械", "电动机械", "电子")
MEASUREMENT_FIELDS = (
    ("measurement_parameter", "测量参数"), ("parameter_unit", "测量参数单位"),
    ("measurement_range", "测量区间"), ("display_digits", "显示位数"),
    ("metrological_performance", "计量性能指标"),
)


def photo_default_width(slot, report_type="水表"):
    if slot.startswith("component:"):
        return 8
    group = photo_slot_group(slot)
    if report_type in GAS_PHOTO_REPORT_TYPES:
        return 10 if group == "internal" else 12
    return PHOTO_SLOTS[group][1]


def active_photo_slots(display_types=None, report_type="水表"):
    paired = {"机械", "电子"}.issubset(display_types or [])
    excluded = {"display"} if paired else {"display_mechanical", "display_electronic"}
    if report_type in GAS_PHOTO_REPORT_TYPES:
        excluded |= GAS_UNUSED_PHOTO_SLOTS
    return {key: (value[0], photo_default_width(key, report_type))
            for key, value in PHOTO_SLOTS.items() if key not in excluded}


def photo_layout_ids(photos, components, display_types=None, report_type="水表"):
    slots = set(active_photo_slots(display_types, report_type)) | set(photos) | {item["id"] for item in components}
    return set(PHOTO_HEADINGS) | {"photo:" + slot for slot in slots} | {
        "heading:" + item["id"] for item in components
    }


class PhotoLayout(BaseModel):
    model_config = ConfigDict(extra="forbid")

    blank_lines: Optional[StrictInt] = Field(default=None, ge=0)
    # Retained for old task payloads; appendix photos now always flow naturally.
    page_break_before: Optional[bool] = Field(default=None, strict=True)


class MeasurementParameter(BaseModel):
    model_config = ConfigDict(extra="forbid")

    measurement_parameter: str = ""
    parameter_unit: str = ""
    measurement_range: str = ""
    display_digits: str = ""
    metrological_performance: str = ""
    volume_display: str = ""


class EquipmentRun(BaseModel):
    model_config = ConfigDict(extra="forbid", strict=True)

    text: str
    italic: bool = False
    script: Literal["normal", "sup", "sub"] = "normal"


DEFAULT_EQUIPMENT_STORAGE_LOCATION = "张衡路1500号"
EQUIPMENT_TRIAL_GROUPS = ("液体组", "气体组", "环境试验组", "电磁兼容组")
EQUIPMENT_TRIAL_GROUP_BY_IDENTITY = {
    ("步入式高低温试验箱", "GDW-HS15.6C", "202104-2"): "气体组",
    ("差压变送器", "HP610-(±500)PaD-N-8-5", "712457"): "气体组",
    ("静电放电模拟器", "ESS-B3011A", "ESS2042260/容-608-004"): "电磁兼容组",
    ("对数周期天线", "ATR80M6G", "0554492/容-029-30"): "电磁兼容组",
    ("高增益对数周期天线", "HL046E", "100317/容-607-013"): "电磁兼容组",
    ("差压变送器", "HP610-(±500)PaD-N-8-5", "712456"): "气体组",
    ("跌落试验台", "DLJ-150", "2314040/环-345"): "环境试验组",
    ("功率放大器", "BBA150-BC1000", "103963/容-607-008"): "电磁兼容组",
    ("钟罩式气体流量标准装置", "LJQ-100", "2816"): "气体组",
    ("音速喷嘴式燃气表检验装置", "MJ2210-03", "140301"): "气体组",
    ("电动振动试验系统", "DC-1000-15", "230611056/环-346"): "环境试验组",
    ("步入式高低温湿热实验室", "FW10065VWAF", "22177009/环-294"): "环境试验组",
    ("精密压力表", "(0~100)kPa", "2006-10-002"): "环境试验组",
    ("信号发生器", "SMB100B", "101897/容-607-001"): "电磁兼容组",
    ("差压变送器", "HP610-(±500)PaD-N-8-5", "712455"): "气体组",
    ("直流稳压电源", "ADC-0600033-15", "D211030002/安-154"): "环境试验组",
    ("电动振动试验系统", "DC-3200-36", "111040/环-151"): "环境试验组",
    ("压力表", "Y-100", "YB12813840"): "液体组",
    ("1.5立方高低温试验箱", "EW1570", "19175020/环-276"): "环境试验组",
    ("静态质量法水流量标准装置", "DN4~DN300", "001"): "液体组",
    ("静态质量法水流量标准装置", "DN2~DN25", "20140930"): "液体组",
    ("水压影响试验装置", "LJS-P", "2002-124"): "液体组",
    ("水温影响试验装置", "LJS-T", "2002-123"): "液体组",
    ("差压计", "（0~100）kPa", "020516"): "液体组",
    ("磁铁", "70mm×32mm", "001"): "液体组",
    ("高低温湿热试验箱", "EW0470", "15113844/环-213"): "环境试验组",
    ("信号发生器", "SMB 100B", "103743/容-620-001"): "电磁兼容组",
    ("功率放大器", "BBL200-A10000", "101638/容-621-001"): "电磁兼容组",
    ("功率放大器", "NTWPA-1060400200PE", "20019001/容-607-009"): "电磁兼容组",
    ("双锥天线", "VHBD9134-7/16+BBFA9146", "00189/容-618-017"): "电磁兼容组",
}

EQUIPMENT_TRIAL_GROUP_BY_IDENTITY = {
    equipment_identity(*identity): group for identity, group in EQUIPMENT_TRIAL_GROUP_BY_IDENTITY.items()
}


class Equipment(BaseModel):
    id: str = ""
    name: str
    model: str = ""
    performance: str = ""
    number: str
    valid_until: str
    storage_location: str = DEFAULT_EQUIPMENT_STORAGE_LOCATION
    trial_group: Literal["", "液体组", "气体组", "环境试验组", "电磁兼容组"] = ""
    # Plain strings remain authoritative; formatting is optional per field.
    formatting: Optional[Dict[Literal["name", "model", "performance", "number"],
                              List[EquipmentRun]]] = None

    @model_validator(mode="after")
    def matching_formatting(self):
        if not self.trial_group:
            self.trial_group = EQUIPMENT_TRIAL_GROUP_BY_IDENTITY.get(
                equipment_identity(self.name, self.model, self.number), "")
        for key, runs in (self.formatting or {}).items():
            if "".join(run.text for run in runs) != getattr(self, key):
                raise ValueError(f"设备 {key} 的格式化文字必须与原文字完全一致。")
        return self

    def trimmed(self):
        """Trim library fields and their runs together, after exact-text validation."""
        data = self.model_dump()
        for key in ("name", "model", "performance", "number", "valid_until", "storage_location", "trial_group"):
            original = data[key]
            data[key] = original.strip()
            if self.formatting is None or key not in self.formatting:
                continue
            start = len(original) - len(original.lstrip())
            end = len(original.rstrip())
            offset = 0
            trimmed_runs = []
            for run in self.formatting[key]:
                left, right = max(start - offset, 0), min(end - offset, len(run.text))
                if left < right:
                    trimmed_runs.append(dict(run.model_dump(), text=run.text[left:right]))
                offset += len(run.text)
            data["formatting"][key] = trimmed_runs
        return Equipment.model_validate(data)


class Candidate(BaseModel):
    value: str
    source: str


class OverviewField(BaseModel):
    key: str
    label: str
    old: str = ""
    value: str = ""
    confirmed: bool = False
    manual_override: bool = False
    candidates: List[Candidate] = Field(default_factory=list)
    observation_manual_override: bool = False


class ResultRow(BaseModel):
    id: str
    requirement: str = ""
    sample_label: str = ""
    sample_id: str = ""
    value: str = ""
    source: str = ""
    match_issue: str = ""
    auto_value: str = ""
    matched_sample_id: str = ""
    manual_override: bool = False
    working_pressure: Optional[Candidate] = None
    resolved_requirement: str = ""


class RecordMatch(BaseModel):
    source: str = ""
    sample_id: str = ""
    model: str = ""
    start: str = ""
    end: str = ""
    temperature: str = ""
    humidity: str = ""


class RecordSummary(BaseModel):
    values: Dict[str, str] = Field(default_factory=dict)
    records: List[RecordMatch] = Field(default_factory=list)
    issues: List[str] = Field(default_factory=list)


class ReportEnvironment(BaseModel):
    values: Dict[str, str] = Field(default_factory=dict)
    issues: List[str] = Field(default_factory=list)


class ObservationSummary(ReportEnvironment):
    sources: Dict[str, str] = Field(default_factory=dict)


class GasCurvePoint(BaseModel):
    flow: str
    error: str


class GasCurve(BaseModel):
    sample_id: str
    source: str
    model: str = ""
    points: List[GasCurvePoint]


class Project(BaseModel):
    id: str
    asset_id: str
    name: str
    variant: str
    rows: List[ResultRow] = Field(default_factory=list)
    water_result_labels: List[str] = Field(default_factory=list)
    water_output_labels: Optional[List[str]] = None
    start: str = ""
    end: str = ""
    temperature: str = ""
    humidity: str = ""
    atmospheric_pressure: str = ""
    pressure_manual_override: bool = Field(default=False, strict=True)
    pressure_auto_value: str = ""
    pressure_sources: List[str] = Field(default_factory=list)
    pressure_issues: List[str] = Field(default_factory=list)
    conclusion: str = ""
    abnormal: str = ""
    evaluators: List[str] = Field(default_factory=list)
    reviewers: List[str] = Field(default_factory=list)
    equipment_ids: List[str] = Field(default_factory=list)
    confirmed: bool = False
    notes: List[str] = Field(default_factory=list)
    source: str = ""
    record_summary: RecordSummary = Field(default_factory=RecordSummary)
    gas_curves: List[GasCurve] = Field(default_factory=list)

    @field_validator("water_output_labels")
    @classmethod
    def valid_output_labels(cls, values):
        if values is not None and (len(values) != len(set(values))
                or any(not re.fullmatch(r"[A-Z]+\d{3}", value) for value in values)):
            raise ValueError("参试样机勾选须使用不重复的A001、B001等报告代号")
        return values

    @field_validator("atmospheric_pressure")
    @classmethod
    def valid_pressure_value(cls, value):
        from atmospheric_pressure import pressure_input_value

        return pressure_input_value(value)

    @field_validator("conclusion", "abnormal", mode="before")
    @classmethod
    def retain_source_text(cls, value):
        return ""


class Sample(BaseModel):
    label: str
    number: str = ""
    number_origin: Literal["", "auto", "manual"] = ""
    auto_number: str = ""
    number_source: str = ""
    record_number: str = ""
    added_manually: bool = Field(default=False, strict=True)
    disassembly: bool = False
    sealed: bool = False


class SampleSchemeProject(BaseModel):
    model_config = ConfigDict(extra="forbid")
    asset_id: str = Field(min_length=1)
    labels: List[str] = Field(default_factory=list)

    @field_validator("labels")
    @classmethod
    def valid_labels(cls, values):
        return Project.valid_output_labels(values)


class SampleOutputScheme(BaseModel):
    model_config = ConfigDict(extra="forbid")
    id: str = ""
    name: str = Field(min_length=1, max_length=100)
    report_type: Literal["水表"] = "水表"
    projects: List[SampleSchemeProject] = Field(min_length=1)
    revision: int = Field(default=0, ge=0)

    @field_validator("name")
    @classmethod
    def nonblank_name(cls, value):
        if not value.strip():
            raise ValueError("样机方案名称不能为空")
        return value.strip()

    @field_validator("projects")
    @classmethod
    def unique_projects(cls, values):
        if len({value.asset_id for value in values}) != len(values):
            raise ValueError("样机方案不能重复保存同一结果素材")
        return values


class ResultMaterialScheme(BaseModel):
    model_config = ConfigDict(extra="forbid")
    id: str = ""
    name: str = Field(min_length=1, max_length=100)
    report_type: ReportType = "水表"
    asset_ids: List[str] = Field(min_length=1)
    revision: int = Field(default=0, ge=0)

    @field_validator("name")
    @classmethod
    def nonblank_name(cls, value):
        if not value.strip():
            raise ValueError("素材方案名称不能为空")
        return value.strip()

    @field_validator("asset_ids")
    @classmethod
    def unique_assets(cls, values):
        if any(not value.strip() for value in values):
            raise ValueError("素材方案不能包含空素材标识")
        if len(values) != len(set(values)):
            raise ValueError("素材方案不能重复保存同一结果素材")
        return values


class SampleCandidate(BaseModel):
    sample_id: str
    model: str = ""
    certificate: str = ""
    source: str = ""


class PhotoEdit(BaseModel):
    model_config = ConfigDict(extra="forbid")
    rotation: Literal[0, 90, 180, 270] = 0
    crop: Optional[Tuple[int, int, int, int]] = None
    width_cm: Optional[float] = Field(default=None, ge=0.1, le=50, allow_inf_nan=False)

    @field_validator("crop", mode="before")
    @classmethod
    def valid_crop(cls, value):
        if value is not None:
            if (not isinstance(value, (tuple, list)) or len(value) != 4
                    or any(type(number) is not int for number in value)):
                raise ValueError("裁剪区域须为四个整数像素坐标")
            left, top, right, bottom = value
            if left < 0 or top < 0 or right <= left or bottom <= top:
                raise ValueError("裁剪区域须非空且位于照片内")
        return value


class ObservationCertificateCandidate(Candidate):
    text: str = ""


class ObservationCertificate(BaseModel):
    value: str = ""
    manual_override: bool = False
    targets: List[str] = Field(default_factory=list)
    originals: List[str] = Field(default_factory=list)
    candidates: List[ObservationCertificateCandidate] = Field(default_factory=list)
    source_issues: List[str] = Field(default_factory=list)
    issues: List[str] = Field(default_factory=list)


class ObservationCertificates(BaseModel):
    explosion: ObservationCertificate = Field(default_factory=ObservationCertificate)
    protection: ObservationCertificate = Field(default_factory=ObservationCertificate)
    issues: List[str] = Field(default_factory=list)


class SupplementaryNote(BaseModel):
    id: str = Field(pattern=r"^custom:[A-Za-z0-9-]+$")
    text: str = ""


class Task(BaseModel):
    id: str = ""
    name: str = "型式评价报告"
    report_type: ReportType = "水表"
    include_atmospheric_pressure: bool = Field(default=False, strict=True)
    overview: List[OverviewField] = Field(default_factory=list)
    is_new_type: Optional[bool] = Field(default=None, strict=True)
    measurement_parameters: Optional[List[MeasurementParameter]] = Field(default=None, min_length=1)
    include_volume_display: bool = Field(default=False, strict=True)
    volume_display_label: str = "体积量显示"
    display_types: Optional[List[Literal["机械", "电动机械", "电子"]]] = None
    items_asset_id: str = ""
    observe_asset_id: str = ""
    outline_id: str = ""
    record_ids: List[str] = Field(default_factory=list)
    photos: Dict[str, str] = Field(default_factory=dict)
    photo_order: List[str] = Field(default_factory=list)
    photo_layout: Dict[str, PhotoLayout] = Field(default_factory=dict)
    gas_photo_layout_version: StrictInt = Field(default=0, ge=0, le=2)
    photo_edits: Dict[str, PhotoEdit] = Field(default_factory=dict)
    photo_components: List[Dict[str, str]] = Field(default_factory=list)
    photo_issues: List[str] = Field(default_factory=list)
    construction_usage: Literal["", "用水", "液体", "气体", "燃气", "其他"] = ""
    construction_custom: str = ""
    construction_instrument: str = ""
    construction_issues: List[str] = Field(default_factory=list)
    production_basis_issues: List[str] = Field(default_factory=list)
    test_basis_issues: List[str] = Field(default_factory=list)
    sealing_mode: Literal["auto", "manual"] = "auto"
    sealing_model: str = ""
    sealing_model_source: str = ""
    sealing_issues: List[str] = Field(default_factory=list)
    other_notes: List[str] = Field(default_factory=list)
    supplementary_notes: List[str] = Field(default_factory=list)
    supplementary_note_ids: List[str] = Field(default_factory=list)
    custom_supplementary_notes: List[SupplementaryNote] = Field(default_factory=list)
    other_information_order: List[str] = Field(default_factory=list)
    supplementary_source: str = ""
    supplementary_issues: List[str] = Field(default_factory=list)
    other_sealing_number: int = 1
    record_receipt_location: str = "张衡路1500号"
    record_test_conclusion: str = ""
    sample_delivery_date: str = ""
    evaluation_summary: ReportEnvironment = Field(default_factory=ReportEnvironment)
    report_environment: ReportEnvironment = Field(default_factory=ReportEnvironment)
    observation_summary: ObservationSummary = Field(default_factory=ObservationSummary)
    observation_certificates: ObservationCertificates = Field(default_factory=ObservationCertificates)
    equipment: List[Equipment] = Field(default_factory=list)
    projects: List[Project] = Field(default_factory=list)
    samples: List[Sample] = Field(default_factory=list)
    sample_schema: Dict[str, Any] = Field(default_factory=dict)
    water_sample_source_schema: Dict[str, Any] = Field(default_factory=dict)
    water_sample_next_numbers: Dict[str, StrictInt] = Field(default_factory=dict)
    sample_table_excluded_columns: List[StrictInt] = Field(default_factory=list)
    classification_code: str = ""
    sample_issues: List[str] = Field(default_factory=list)
    samples_confirmed: bool = False
    sample_candidates: Dict[str, List[SampleCandidate]] = Field(default_factory=dict)
    sample_mapping_issues: List[str] = Field(default_factory=list)
    legacy_sample_fields: List[OverviewField] = Field(default_factory=list)
    # Legacy row edits remain readable in old tasks but do not affect copied modules.
    observations: List[Dict[str, Any]] = Field(default_factory=list)
    items: List[Dict[str, Any]] = Field(default_factory=list)
    revision: int = 0

    @field_validator("volume_display_label")
    @classmethod
    def nonblank_volume_display_label(cls, value):
        if not value.strip():
            raise ValueError("体积量显示列名不能为空")
        return value

    @field_validator("other_information_order")
    @classmethod
    def unique_other_information_order(cls, values):
        if len(values) != len(set(values)):
            raise ValueError("说明条款顺序不能包含重复条目")
        return values

    @field_validator("custom_supplementary_notes")
    @classmethod
    def unique_custom_notes(cls, values):
        if len({value.id for value in values}) != len(values):
            raise ValueError("新增补充说明的标识不能重复")
        return values

    @field_validator("sample_table_excluded_columns")
    @classmethod
    def valid_sample_table_exclusions(cls, values):
        if any(value < 0 for value in values) or len(values) != len(set(values)):
            raise ValueError("样机信息表隐藏列须为不重复的非负列索引")
        return sorted(values)

    @field_validator("classification_code")
    @classmethod
    def valid_classification_code(cls, value):
        if value and value not in CLASSIFICATION_CODES:
            raise ValueError("分类编码须从已录入的对照表中选择")
        return value

    @field_validator("display_types")
    @classmethod
    def unique_display_types(cls, values):
        if values is not None and len(values) != len(set(values)):
            raise ValueError("显示型式不能重复选择")
        return values

    @model_validator(mode="after")
    def fixed_record_cover(self):
        self.record_receipt_location = "张衡路1500号"
        self.record_test_conclusion = ""
        return self

    @model_validator(mode="after")
    def gas_photo_defaults(self):
        if self.report_type == "流量计（气体）":
            self.gas_photo_layout_version = 2
        elif self.report_type == "燃气表":
            if self.gas_photo_layout_version == 0:
                self.photo_layout = {}
            if self.gas_photo_layout_version < 2:
                self.photo_edits = {
                    slot: edit.model_copy(update={"width_cm": None}) for slot, edit in self.photo_edits.items()
                }
                self.gas_photo_layout_version = 2
        elif any(slot.startswith("sealed_overall:") for slot in self.photos):
            raise ValueError("仅燃气表和流量计（气体）支持多张封样照片，请先移除追加的封样照片或新建任务")
        return self

    @model_validator(mode="after")
    def clear_inactive_display_photos(self):
        inactive = set(PHOTO_SLOTS) - set(active_photo_slots(self.display_types))
        self.photos = {slot: identity for slot, identity in self.photos.items() if photo_slot_group(slot) not in inactive}
        self.photo_edits = {slot: edit for slot, edit in self.photo_edits.items() if photo_slot_group(slot) not in inactive}
        if len(set(self.photo_order)) != len(self.photo_order):
            raise ValueError("照片排序不能包含重复位置")
        self.photo_order = [slot for slot in self.photo_order if photo_slot_group(slot) not in inactive]
        self.photo_layout = {
            target: layout for target, layout in self.photo_layout.items()
            if not (target.startswith("photo:") and photo_slot_group(target[6:]) in inactive)
        }
        if set(self.photo_order) - set(self.photos):
            raise ValueError("照片排序包含不存在的照片位置")
        if set(self.photo_layout) - photo_layout_ids(self.photos, self.photo_components, self.display_types, self.report_type):
            raise ValueError("照片排版设置包含不存在的照片或标题位置，请移除失效设置")
        return self

    @model_validator(mode="before")
    @classmethod
    def preserve_legacy_sealing(cls, data):
        if isinstance(data, dict) and "sealing_mode" not in data:
            data = dict(data)
            fields = [f.model_dump() if isinstance(f, OverviewField) else f
                      for f in data.get("overview", [])]
            data["sealing_mode"] = "manual" if any(
                f.get("key") == "sealing" and f.get("value", "").strip() for f in fields
            ) else "auto"
        return data

    @field_validator("photos")
    @classmethod
    def valid_photo_references(cls, photos):
        for slot, identity in photos.items():
            if photo_slot_group(slot) not in PHOTO_SLOTS and not re.fullmatch(r"component:[0-9a-f]{32}:[1-9][0-9]*", slot):
                raise ValueError("无法识别的附件2照片位置：" + slot)
            if not re.fullmatch(r"[0-9a-f]{32}", identity):
                raise ValueError("照片必须引用本机上传记录，不能使用路径或外部网址。")
        return photos

    @model_validator(mode="before")
    @classmethod
    def preserve_legacy_samples(cls, data):
        if isinstance(data, dict):
            data = dict(data)
            if not isinstance(data.get("overview", []), (list, tuple)) or not isinstance(
                    data.get("legacy_sample_fields", []), (list, tuple)):
                raise ValueError("总览及历史样机字段必须为列表。")
            legacy = list(data.get("legacy_sample_fields", []))
            for field in data.get("overview", []):
                item = field.model_dump() if isinstance(field, OverviewField) else field
                if not isinstance(item, dict) or not isinstance(item.get("key"), str) or not isinstance(item.get("value", ""), str):
                    raise ValueError("总览字段必须包含有效的字段标识和文字值。")
                if item["key"] in SAMPLE_OVERVIEW_FIELDS and item.get("value", "").strip():
                    if item not in legacy:
                        legacy.append(item)
            data["legacy_sample_fields"] = legacy
        return data

    @field_validator("overview")
    @classmethod
    def editable_overview(cls, fields: List[OverviewField]) -> List[OverviewField]:
        for field in fields:
            if field.key == "sealing":
                field.label = "其他"
        return [field for field in fields if field.key not in HIDDEN_OVERVIEW_FIELDS]


class ResultMaterialRequest(BaseModel):
    model_config = ConfigDict(extra="forbid")
    version: str = Field(pattern=r"^[0-9a-f]{64}$")
    asset_ids: List[str] = Field(min_length=1, max_length=24)

    @field_validator("asset_ids")
    @classmethod
    def unique_assets(cls, values):
        if len(values) != len(set(values)) or any(not value.strip() or len(value) > 100 for value in values):
            raise ValueError("结果素材标识不能为空、重复或过长。")
        return values


class RecoveryRequest(BaseModel):
    task: Task
    result: Optional[Project] = None
    equipment: Optional[Equipment] = None

    @model_validator(mode="after")
    def matching_result_rows(self):
        if self.result:
            project = next((p for p in self.task.projects if p.id == self.result.id), None)
            if project is None or [r.id for r in project.rows] != [r.id for r in self.result.rows]:
                raise ValueError("暂存结果与项目行结构不一致，未恢复。请下载暂存内容核对。")
        return self


class GenerateRequest(BaseModel):
    task: Task
    mode: str = "auto"
    native_word: bool = Field(False, description="是否使用本机Word执行V4；默认false直接合成下载。")
    pagination_strategy: Literal["standard", "optimized_v4"] = "standard"

    @field_validator("pagination_strategy", mode="before")
    @classmethod
    def reject_retired_pagination(cls, value: Any) -> Any:
        if value in ("optimized", "optimized_v2", "optimized_v3"):
            raise ValueError("优化自动重排V1、V2、V3已移除，请刷新页面并选择直接生成Word或V4。")
        return value

    @model_validator(mode="after")
    def compatible_pagination_strategy(self):
        if self.pagination_strategy == "standard" and self.native_word:
            raise ValueError("自动重排后生成Word已移除，请刷新页面并选择直接生成Word或V4。")
        if self.pagination_strategy == "optimized_v4" and not self.native_word:
            raise ValueError("优化自动重排需启用本机Word，不能与直接生成同时选择。")
        return self


# Document-module interface, implemented by documents.py:
# catalog(root: Path) -> {"assets": [{id, kind, name, project, variant, folders}],
#                        "overview": [OverviewField dict],
#                        "hidden_overview_fields": [str], "warnings": [str]}
# kind is "items", "observe", or "result"; ids are opaque stable strings.
# project_from_asset(root, asset_id) -> Project dict (no historical results).
# section_rows(root, asset_id) -> [{id, label, value, note}] for observe/items.
# parse_upload(path: Path, kind: str) -> {
#   "fields": {overview_key: [{"value": str, "source": str}]},
#   "records": [{project: str, sample_id: str, model: str,
#                start: str, end: str, temperature: str, humidity: str,
#                rows: [{label: str, value: str}], source: str}],
#   "warnings": [str], "outline": {"components": str, "coverage": str}
# }
# apply_records(project: dict, parsed: list[dict]) -> Project dict.
# Never overwrite nonempty user data; ambiguous candidates become notes.
# order_projects(root: Path, items_asset_id: str, projects: list[dict]) -> list[dict]
# Uses the selected items document order, NOT click order (latest user decision).
# compose_report(root: Path, task: dict, outline_path: Optional[Path],
#                output: Path, draft: bool) -> {"warnings": [str]}
# Sections, rows and relationships are mapped by document module.
# Report file remains editable; master headers are authoritative.
