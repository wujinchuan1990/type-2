"""Fill agreed observation content and fit basic-info label columns."""

import re
from copy import deepcopy
from difflib import SequenceMatcher
from math import ceil

from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.table import Table

import documents as d
from contracts import OBSERVATION_MANUAL_FIELDS
from word_fonts import text_width


LABELS = {
    "accuracy_class": "准确度等级", "q3": "Q3", "q3_q1": "Q3/Q1",
    "manufacturer": "制造商名称", "maximum_pressure": "最大允许压力",
    "installation": "安装方式", "pressure_loss": "压力损失等级",
    "flow_sensitivity": "流场敏感度等级", "environment_class": "环境等级",
    "electromagnetic_class": "电磁环境等级",
    **OBSERVATION_MANUAL_FIELDS,
}
SOURCE_LABELS = {
    "最大允许压力": "maximum_pressure", "安装方式": "installation", "方位限制": "installation",
    "压力损失等级": "pressure_loss", "上游流场敏感度等级": "upstream",
    "下游流场敏感度等级": "downstream", "流场敏感度等级": "flow_sensitivity",
    "气候和机械环境等级": "environment_class", "环境等级": "environment_class",
    "电磁环境等级": "electromagnetic_class",
}
REMARK_LABELS = [*LABELS.values(), "型式批准标志和编号", "计量单位", "编号", "流动方向", "Q2/Q1", "Q4/Q3"]
COPY_WARNING = (
    "观察素材首部按本次总览和样机分组填写；仅更新指定水表备注：大纲参数（安装方式亦识别方位限制）及总览手填制造年月/电池更换期限；"
    "防爆及防护证书类项目另按素材与章节中的证书编号填写，封印和防护备注不改。"
    "未填写保留对应原文并提示。其余技术要求和已有勾选保留，未由系统核验。项目一览表同步填写同类证书编号，其余原样复制。"
)


def copy_warning(report_type="水表"):
    if report_type == "水表":
        return COPY_WARNING
    return (
        f"{report_type}观察素材首部型号及项目内样机编号按本次总览和样机分组填写，拆样样机不列入观察；"
        "防爆及防护证书类项目另填证书编号，不填充水表备注参数，"
        "不套用水表条款检查。"
        + ("燃气表流量范围要求按本次型号及总览最小/最大流量替换，分界流量=0.1×最大流量；"
           if report_type == "燃气表" else "")
        + "其余技术要求、备注和原结论保留，统一英文Arial、中文宋体；"
        "素材内容不代表本次已核实，项目一览表同步填写同类证书编号，其余原样复制。"
    )


def _label_pattern(labels):
    return re.compile("(" + "|".join(
        r"\s*".join(re.escape(char) for char in label)
        for label in sorted(labels, key=len, reverse=True)
    ) + r")\s*[：:]\s*")


def outline_values(outline, source="当前大纲", *, report_type="水表"):
    result = {"values": {}, "sources": {}, "issues": []}
    if report_type != "水表":
        return result

    def accept(key, entries):
        distinct = list(dict.fromkeys(value.strip() for value, _where in entries if value.strip()))
        if len(distinct) != 1:
            result["issues"].append(f"观察备注/{LABELS.get(key, key)}：大纲来源缺失或有冲突，保留素材原文，请核对。")
            return
        result["values"][key] = distinct[0]
        result["sources"][key] = "；".join(dict.fromkeys(where for _value, where in entries))

    if outline is None:
        result["issues"].append("观察备注：尚未选择本次大纲，自动参数保留素材原文。")
        return result
    manufacturers = d._field_candidates(outline, source).get("manufacturer", [])
    accept("manufacturer", [(c["value"], c["source"]) for c in manufacturers])
    try:
        schema, _table = d.outline_samples(outline)
        headers = [re.sub(r"\s+", "", value) for value in schema["headers"]]
        for key, pattern in (("q3", r"^Q3(?:[（(]|$)"), ("q3_q1", r"^Q3/Q1$"),
                             ("accuracy_class", r"^准确度(?:等级)?$")):
            columns = [i for i, text in enumerate(headers) if re.match(pattern, text, re.I)]
            if len(columns) != 1:
                accept(key, [])
                continue
            values = [group["values"][columns[0]].strip() for group in schema["groups"]]
            if not all(values):
                accept(key, [])
                continue
            accept(key, [("、".join(dict.fromkeys(values)), f"{source} · 第5项实际提供样机表/{headers[columns[0]]}")])
    except ValueError as exc:
        result["issues"].append("观察备注/样机参数：" + str(exc))
    try:
        body = outline.element.body
        start = d._unique_anchor(body, "提供试验样机")
        end = d._unique_anchor(body, "次型式评价项目及所用设备一览表")
        paragraphs = [p for p in list(body)[body.index(start) + 1:body.index(end)] if p.tag == qn("w:p")]
    except ValueError as exc:
        result["issues"].append("观察备注/第5项参数：" + str(exc))
        return result
    entries = {key: [] for key in SOURCE_LABELS.values()}
    pattern = _label_pattern(SOURCE_LABELS)
    for paragraph in paragraphs:
        raw = d.Paragraph(paragraph, outline).text
        matches = list(pattern.finditer(raw))
        for index, match in enumerate(matches):
            label = re.sub(r"\s+", "", match.group(1))
            stop = matches[index + 1].start() if index + 1 < len(matches) else len(raw)
            value = raw[match.end():stop].strip(" \t\r\n，,；;。")
            entries[SOURCE_LABELS[label]].append((value, f"{source} · 第5项/{label}"))
    for key in ("maximum_pressure", "installation", "pressure_loss", "environment_class", "electromagnetic_class"):
        accept(key, entries[key])
    direct = entries["flow_sensitivity"]
    upstream = list(dict.fromkeys(value for value, _ in entries["upstream"] if value))
    downstream = list(dict.fromkeys(value for value, _ in entries["downstream"] if value))
    if entries["upstream"] or entries["downstream"]:
        if len(upstream) == len(downstream) == 1:
            direct = [*direct, (upstream[0] + "/" + downstream[0], f"{source} · 第5项/上游及下游流场敏感度等级")]
        else:
            direct = []
    accept("flow_sensitivity", direct)
    return result


def _slots(cell):
    raw = "".join(d._xp(cell, ".//w:t/text()"))
    matches = list(_label_pattern(REMARK_LABELS).finditer(raw))
    slots = {}
    for index, match in enumerate(matches):
        label = re.sub(r"\s+", "", match.group(1))
        # Read-only labels delimit editable values but never become replacement slots.
        if label not in LABELS.values():
            continue
        if label in slots:
            raise ValueError("观察备注同一单元格字段重复，无法安全替换：" + label)
        stop = matches[index + 1].start() if index + 1 < len(matches) else len(raw)
        value = raw[match.end():stop]
        end = stop - (len(value) - len(value.rstrip()))
        slots[label] = (match.end(), end, raw[match.end():end])
    return slots


def _targets(observe):
    targets = []
    found = set()
    required = {
        "5.3": LABELS,
        "6.1": {"q3": "Q3", "q3_q1": "Q3/Q1"},
        "7.7.2": {"environment_class": "环境等级"},
        "7.7.3": {"electromagnetic_class": "电磁环境等级"},
        "7.7.4": {"battery_replacement_deadline": OBSERVATION_MANUAL_FIELDS["battery_replacement_deadline"]},
    }
    seen_bare = set()
    for row in d._section_schema(observe, "observe"):
        cell = row["_outputs"][-1]
        label = re.sub(r"\s+", "", row["label"])
        clause = next((code for code in ("5.3", "6.1", "7.7.2", "7.7.3", "7.7.4")
                       if re.search(r"(?<![\d.])" + re.escape(code) + r"(?![\d.])", label)), None)
        if clause is None:
            continue
        if clause in ("7.7.2", "7.7.3"):
            key = "environment_class" if clause == "7.7.2" else "electromagnetic_class"
            if not d._text(cell) and clause in seen_bare:
                continue
            seen_bare.add(clause)
            slots = _slots(cell)
            start, end, old = slots.get(LABELS[key], (0, len("".join(d._xp(cell, ".//w:t/text()"))), d._text(cell)))
            targets.append((cell, start, end, key, old))
            found.add((clause, key))
            continue
        slots = _slots(cell)
        for key, name in required[clause].items():
            if name in slots:
                start, end, old = slots[name]
                targets.append((cell, start, end, key, old))
                found.add((clause, key))
    missing = [clause + "/" + label for clause, fields in required.items() for key, label in fields.items()
               if (clause, key) not in found]
    if missing:
        raise ValueError("观察素材缺少指定水表备注填写位置：" + "、".join(missing))
    return targets


def manual_originals(observe):
    targets = _targets(observe)
    return {key: "\n".join(dict.fromkeys(old for _cell, _start, _end, field, old in targets if field == key))
            for key in OBSERVATION_MANUAL_FIELDS}


def observation_values(outline, task, source="当前大纲"):
    result = outline_values(outline, source, report_type=task.get("report_type", "水表"))
    if task.get("report_type", "水表") != "水表":
        return result
    fields = {field["key"]: field for field in task.get("overview", [])}
    for key, label in OBSERVATION_MANUAL_FIELDS.items():
        field = fields.get(key, {})
        value = field.get("value", "").strip()
        if value:
            result["values"][key] = value
            result["sources"][key] = "基本信息总览 · 手工填写"
        else:
            result["issues"].append(f"观察备注/{label}：尚未手工填写，保留素材原文，请核对。")
    return result


def _basic_table(observe):
    tables = [table for table in d._xp(observe.element.body, ".//w:tbl")
              if all(label in d._text(table) for label in ("申请单位", "计量器具名称", "规格型号", "样机编号"))]
    if len(tables) != 1:
        raise ValueError("观察素材首部基本信息表缺失或不唯一，无法安全填写。")
    return tables[0]


def _fit_basic_labels(observe):
    table = _basic_table(observe)
    columns = d._xp(table, "./w:tblGrid/w:gridCol")
    rows = d._rows(table)
    if len(columns) != 4 or any(len(d._cells(row)) != 4 for row in rows):
        raise ValueError("观察基本信息表不是四列结构，无法安全调整标签列宽。")
    widths = [int(column.get(qn("w:w"))) for column in columns]
    table_styles = []
    style = Table(table, observe).style
    while style is not None:
        table_styles.append(style.element)
        style = style.base_style
    for index, labels in ((0, {"申请单位", "规格型号"}), (2, {"计量器具名称", "样机编号"})):
        needed = widths[index]
        for row in rows:
            cell = d._cells(row)[index]
            if d._norm(d._text(cell)).rstrip("：:") not in labels:
                raise ValueError("观察基本信息标签无法识别：" + d._text(cell))
            margins = 0
            for side in ("left", "right"):
                candidates = d._xp(cell, f"./w:tcPr/w:tcMar/w:{side}/@w:w")
                candidates += d._xp(table, f"./w:tblPr/w:tblCellMar/w:{side}/@w:w")
                for table_style in table_styles:
                    candidates += d._xp(table_style, f"./w:tblPr/w:tblCellMar/w:{side}/@w:w")
                margins += int(candidates[0]) if candidates else 108
            paragraphs = d._xp(cell, "./w:p")
            if len(paragraphs) != 1 or d._xp(cell, ".//w:br | .//w:tab"):
                raise ValueError("观察基本信息标签含手动换行或制表符，无法仅靠列宽单行显示。")
            paragraph = d.Paragraph(paragraphs[0], observe)
            points = 0
            for run in paragraph.runs:
                size = run.font.size
                style = paragraph.style
                while size is None and style is not None:
                    size = style.font.size
                    style = style.base_style
                defaults = d._xp(observe.styles.element, "./w:docDefaults/w:rPrDefault/w:rPr/w:sz/@w:val")
                if size is None and not defaults:
                    raise ValueError("观察基本信息标签字号无法确定。")
                font_size = size.pt if size is not None else int(defaults[0]) / 2
                scale = d._xp(run._r, "./w:rPr/w:w/@w:val")
                spacing = d._xp(run._r, "./w:rPr/w:spacing/@w:val")
                points += text_width(run.text, font_size) * (int(scale[0]) / 100 if scale else 1)
                points += len(run.text) * (int(spacing[0]) / 20 if spacing else 0)
            indent = paragraphs[0].pPr.ind if paragraphs[0].pPr is not None else None
            inset = sum(int(indent.get(qn("w:" + side), "0")) for side in ("left", "right", "firstLine")) if indent is not None else 0
            needed = max(needed, ceil(points * 20) + margins + max(0, inset) + 20)
        growth = needed - widths[index]
        if widths[index + 1] <= growth:
            raise ValueError("观察基本信息标签过长，原表总宽内没有足够的填写列宽。")
        widths[index] = needed
        widths[index + 1] -= growth
    for column, width in zip(columns, widths):
        column.set(qn("w:w"), str(width))
    for row in rows:
        for cell, width in zip(d._cells(row), widths):
            cell.get_or_add_tcPr().get_or_add_tcW().set(qn("w:w"), str(width))
            cell.tcPr.tcW.set(qn("w:type"), "dxa")
    table.tblPr.get_or_add_tblLayout().set(qn("w:type"), "fixed")


def _fill_basic_lines(cell, values, *, models=False):
    paragraphs = d._xp(cell, "./w:p")
    if not models:
        if len(paragraphs) == 1:
            d._put(paragraphs[0], "\n".join(values))
        else:
            for index, value in enumerate(values):
                paragraph = paragraphs[index] if index < len(paragraphs) else deepcopy(paragraphs[-1])
                if index >= len(paragraphs):
                    cell.append(paragraph)
                d._put(paragraph, value)
            for paragraph in paragraphs[len(values):]:
                cell.remove(paragraph)
            if not values:
                paragraph = deepcopy(paragraphs[0])
                d._put(paragraph, "")
                cell.append(paragraph)
        return
    slots = [p for p in paragraphs if d._text(p).strip()]
    if not slots:
        raise ValueError("观察素材规格型号缺少可识别的填写段落。")
    values = [value for value in values if value]
    for index, value in enumerate(values):
        if index < len(slots):
            paragraph = slots[index]
        else:
            paragraph = deepcopy(slots[-1])
            separator = slots[-1].getnext()
            if separator is not None and separator.tag == qn("w:p") and not d._text(separator).strip():
                blank = deepcopy(separator)
                slots[-1].addnext(blank)
                blank.addnext(paragraph)
            else:
                slots[-1].addnext(paragraph)
            slots.append(paragraph)
        d._put(paragraph, value)
    for paragraph in slots[len(values):]:
        previous = paragraph.getprevious()
        if previous is not None and previous.tag == qn("w:p") and not d._text(previous).strip():
            cell.remove(previous)
        cell.remove(paragraph)


def fill_observation_basics(observe, task):
    table = _basic_table(observe)
    _fit_basic_labels(observe)
    issues = []
    fields = {field["key"]: field.get("value", "").strip() for field in task.get("overview", [])}
    for label, key in (("申请单位", "manufacturer"), ("计量器具名称", "instrument_name")):
        targets = []
        for row in d._rows(table):
            cells = d._cells(row)
            targets.extend(cells[i + 1] for i, cell in enumerate(cells[:-1])
                           if d._text(cell).strip().rstrip("：:") == label)
        if len(targets) != 1:
            raise ValueError("观察素材首部字段缺失或重复：" + label)
        if fields.get(key):
            d._put(targets[0], fields[key])
        else:
            issues.append(f"观察基本信息/{label}：本次总览未填写，保留素材原文，请核对。")
    schema = task.get("sample_schema", {})
    for row in d._rows(table):
        cells = d._cells(row)
        if len(cells) == 4 and "规格型号" in d._text(cells[0]):
            d._set_font_size(cells[1], 12)
    groups = schema.get("groups", [])
    if not groups:
        issues.append("观察基本信息/规格型号、样机编号：缺少本次样机分组，保留素材原文，请核对。")
        return issues
    headers = [re.sub(r"\s+", "", value) for value in schema["headers"]]
    if headers.count("型号") != 1 or headers.count("规格") != 1:
        raise ValueError("观察基本信息的本次样机表必须有唯一的型号、规格列。")
    source_rows = [row for row in d._rows(table) if "规格型号" in d._text(d._cells(row)[0])]
    if len(source_rows) != 1 or len(d._cells(source_rows[0])) != 4:
        raise ValueError("观察素材规格型号/样机编号行结构无法安全映射。")
    by_label = {sample["label"]: sample.get("number", "").strip() for sample in task.get("samples", [])}
    disassembly_labels = {sample["label"] for sample in task.get("samples", []) if sample.get("disassembly")}
    labels = [label for group in groups for label in group["labels"]]
    if len(by_label) != len(task.get("samples", [])) or set(by_label) != set(labels):
        raise ValueError("观察基本信息样机清单与分组不一致，请重新核对。")
    original = {}
    original_models = []
    for row in source_rows:
        original_models.extend(line.strip() for paragraph in d._xp(d._cells(row)[1], "./w:p")
                               for line in d.Paragraph(paragraph, observe).text.splitlines() if line.strip())
        for paragraph in d._xp(d._cells(row)[3], "./w:p"):
            for value in d.Paragraph(paragraph, observe).text.splitlines():
                value = value.strip()
                if re.fullmatch(r"[A-Z]+\d{3}", value):
                    original[value] = value
    if len(source_rows) == len(groups):
        for group, row in zip(groups, source_rows):
            previous_numbers = [line for paragraph in d._xp(d._cells(row)[3], "./w:p")
                                for line in d.Paragraph(paragraph, observe).text.split("\n")]
            previous_labels = group["labels"]
            if len(previous_numbers) != len(previous_labels):
                previous_labels = [label for label in previous_labels if label not in disassembly_labels]
            if len(previous_numbers) == len(previous_labels):
                original.update(zip(previous_labels, previous_numbers))
    cells = d._cells(source_rows[0])
    models, numbers = [], []
    for index, group in enumerate(groups):
        model = group["values"][headers.index("型号")].strip()
        specification = group["values"][headers.index("规格")].strip()
        if model and specification:
            models.append(model + " " + specification)
        else:
            models.append(original_models[index] if index < len(original_models) else "")
            issues.append(f"观察基本信息/第{index + 1}组规格型号：本次值缺失，保留素材原文，请核对。")
        for label in group["labels"]:
            if label in disassembly_labels:
                continue
            number = by_label[label]
            if not number:
                issues.append(f"观察基本信息/样机编号{label}：真实编号未填写，保留对应原文；无对应原位则留空，请核对。")
            numbers.append(number or original.get(label, ""))
    _fill_basic_lines(cells[1], models, models=True)
    d._set_font_size(cells[1], 12)
    _fill_basic_lines(cells[3], numbers)
    return issues


def fill_observation_identities(observe, task):
    from optimized_observation import _lines, _put_numbers, source_layout
    from separation_experimental_documents import _plain_cell, _remove_ephemeral_bookmarks
    import observation_pagination as op

    _remove_ephemeral_bookmarks(observe)
    table, shape, rows, original = source_layout(observe)
    schema = task.get("sample_schema", {})
    headers = [d._norm(value) for value in schema.get("headers", [])]
    model_columns = [i for i, value in enumerate(headers) if value in {"型号", "型号规格", "规格型号"}]
    if len(model_columns) != 1 or not schema.get("groups"):
        raise ValueError("观察基本信息缺少本次型号分组或唯一型号列，不能沿用素材占位值。")
    samples = {sample["label"]: sample for sample in task.get("samples", [])}
    labels = [label for group in schema["groups"] for label in group["labels"]]
    if len(samples) != len(task.get("samples", [])) or set(samples) != set(labels):
        raise ValueError("观察基本信息样机清单与本次型号分组不一致。")
    groups, issues = [], []
    for index, group in enumerate(schema["groups"], 1):
        values = group["values"]
        model = values[model_columns[0]].strip()
        if not model:
            raise ValueError(f"观察基本信息第{index}组型号未填写，不能保留素材型号。")
        selected = [samples[label] for label in group["labels"] if not samples[label].get("disassembly")]
        numbers = [sample.get("number", "").strip() for sample in selected]
        if any(re.search(r"[\r\n\v]", number) for number in numbers):
            raise ValueError("观察真实样机编号含换行，不能把一个编号拆为多台。")
        issues.extend(f"观察样机{sample['label']}：真实编号未填写，对应编号及叉号留空，请核对。"
                      for sample, number in zip(selected, numbers) if not number)
        specifications = {header: value.strip() for header, value in zip(headers, values)}
        if shape == "four":
            model = " ".join(value for value in (model, specifications.get("规格", "")) if value)
            texts = ("规格型号", model, "样机编号")
        else:
            diameter = specifications.get("口径") or specifications.get("规格", "")
            accuracy = specifications.get("准确度等级", "")
            if not diameter or not accuracy:
                raise ValueError("观察五列型号表缺少本次口径或准确度等级，不能保留旧规格。")
            texts = (str(index), model, diameter, accuracy)
        groups.append((texts, numbers))
    numbers = [number for _texts, values in groups for number in values]
    actual = [number for number in numbers if number]
    if len(actual) != len(set(actual)):
        raise ValueError("观察真实样机编号重复，不能建立编号与叉号对应。")
    old_numbers = {number for group in original for number in group["numbers"]}
    projects, replacements = {}, []
    for project_table in d._tables(observe):
        project_rows = d._rows(project_table)
        if not project_rows or "型式评价大纲章节号及项目名称" not in d._text(project_rows[0]):
            continue
        columns = list(op._columns(project_rows[0]))
        op._layout(project_table)
        if len(columns) != 6:
            continue
        if d._norm(d._text(columns[2][2])) != "样机编号":
            raise ValueError("观察项目第三列不是样机编号，不能猜填。")
        sample_span = columns[2][:2]
        mark_spans = op._mark_columns(project_table)
        for _table, row, identity in op._project_rows([project_table]):
            cells = {(start, stop): cell for start, stop, cell in op._columns(row)}
            if sample_span not in cells:
                raise ValueError("观察项目样机编号列无法唯一定位。")
            sample_cell = cells[sample_span]
            previous = _lines(sample_cell)
            if not previous:
                continue
            if len(previous) != len(set(previous)) or not set(previous) <= old_numbers:
                raise ValueError("观察项目样机编号不属于源样机清单，不能猜测本次样机对应范围。")
            marks = []
            for span in mark_spans:
                mark = cells[span]
                if op._is_mark(mark):
                    if len("".join(d._text(mark).split())) != len(previous):
                        raise ValueError("观察项目原编号与叉号数量不一致，不能猜填结论。")
                    symbol = "".join(d._text(mark).split())[0]
                    marks.append((span, mark, symbol))
            projects.setdefault(identity[1][1], []).append((sample_cell, previous, marks))
    for entries in projects.values():
        partial = [entry for entry in entries if set(entry[1]) != old_numbers]
        if partial:
            previous = [number for _cell, values, _marks in partial for number in values]
            if (len(previous) != len(set(previous)) or set(previous) != old_numbers
                    or len(partial) != len(entries)
                    or len({tuple((span, symbol) for span, _cell, symbol in marks)
                            for _sample, _values, marks in partial}) != 1):
                raise ValueError("观察跨页编号片段不构成完整且一致的源样机清单，不能猜测对应关系。")
        for index, (sample_cell, _previous, marks) in enumerate(entries):
            values = [] if partial and index else numbers
            changes = [(sample_cell, values)]
            changes.extend((mark, [symbol if number else "" for number in values])
                           for _span, mark, symbol in marks)
            for cell, _values in changes:
                _plain_cell(cell)
            replacements.extend(changes)
    for row in rows:
        for cell in d._cells(row):
            _plain_cell(cell)
    template = deepcopy(rows[0])
    for row in rows:
        table.remove(row)
    for texts, values in groups:
        row = deepcopy(template)
        cells = d._cells(row)
        for cell, text in zip(cells, texts):
            d._put(cell, text)
        _put_numbers(cells[-1], values)
        table.append(row)
    fields = {field["key"]: field.get("value", "").strip() for field in task.get("overview", [])}
    for index, key in ((1, "manufacturer"), (3, "instrument_name")):
        if fields.get(key):
            d._put(d._cells(d._rows(table)[0])[index], fields[key])
        else:
            issues.append(f"观察基本信息/{key}：本次总览未填写，保留素材原文，请核对。")
    for cell, values in replacements:
        _put_numbers(cell, values)
    return issues


def prepare_observation_export(observe):
    import observation_pagination as op
    from separation_experimental_documents import _plain_cell
    from separation_experimental_layout import _prepare_requirement_flow, _verify_observation_content

    original = deepcopy(observe)
    tables = [table for table in d._tables(observe) if d._rows(table)
              and "型式评价大纲章节号及项目名称" in d._text(d._rows(table)[0])]
    for table in tables:
        if table.getparent() is None:
            continue
        while True:
            separator = table.getnext()
            following = separator.getnext() if separator is not None else None
            if following not in tables or separator.tag != qn("w:p"):
                break
            left, right = d._rows(table)[-1], d._rows(following)[1]
            name_end, remarks_start = op._layout(table)
            left_columns, right_columns = list(op._columns(left)), list(op._columns(right))
            if (d.etree.tostring(table.tblGrid, method="c14n", exclusive=True)
                    != d.etree.tostring(following.tblGrid, method="c14n", exclusive=True)
                    or d._text(d._rows(table)[0]) != d._text(d._rows(following)[0])
                    or [(a, b) for a, b, _c in left_columns] != [(a, b) for a, b, _c in right_columns]
                    or any(not d._text(c).strip() or d._norm(d._text(c)) != d._norm(d._text(other))
                           for (a, _b, c), (_x, _y, other) in zip(left_columns, right_columns) if a < name_end)
                    or any(d._text(c).strip() for a, _b, c in right_columns
                           if name_end < a < remarks_start)):
                break
            blank = deepcopy(separator)
            for page_break in d._xp(blank, ".//w:br[@w:type='page'] | ./w:pPr/w:pageBreakBefore"):
                page_break.getparent().remove(page_break)
            if not d._outline_spacing_paragraph(blank):
                break
            for _a, _b, cell in left_columns + right_columns:
                _plain_cell(cell)
            for (start, _stop, target), (_a, _b, source) in zip(left_columns, right_columns):
                if start not in (name_end, remarks_start):
                    continue
                paragraphs, continuation = d._xp(target, "./w:p"), d._xp(source, "./w:p")
                if (paragraphs and continuation
                        and paragraphs[-1].get(qn("w14:paraId"))
                        and paragraphs[-1].get(qn("w14:paraId")) == continuation[0].get(qn("w14:paraId"))):
                    restored = deepcopy(paragraphs[-1])
                    for child in list(continuation[0]):
                        if child.tag != qn("w:pPr"):
                            restored.append(child)
                    source.replace(continuation[0], restored)
                    target.remove(paragraphs[-1])
                    if not d._xp(target, "./w:p"):
                        target.append(OxmlElement("w:p"))
            for row in d._rows(following)[1:]:
                table.append(row)
            following.getparent().remove(following)
            separator.getparent().remove(separator)
    _prepare_requirement_flow([table for table in tables if table.getparent() is not None])
    _verify_observation_content(original, observe)


def _compact_pressure_loss(cell):
    paragraphs = d._xp(cell, "./w:p")
    label_pattern = _label_pattern(["压力损失等级"])
    starts = [p for p in paragraphs if label_pattern.search(d._text(p))]
    if len(starts) != 1:
        raise ValueError("压力损失等级备注段落缺失或重复。")
    paragraph = starts[0]
    index = paragraphs.index(paragraph)
    # The source puts the label and Δp expression in separate paragraphs.
    while not re.search(r"[△Δ]\s*p", d._text(paragraph)):
        index += 1
        if index >= len(paragraphs):
            return
        following = paragraphs[index]
        if d._text(following).strip() and not re.match(r"\s*[△Δ]\s*p", d._text(following)):
            return
        for child in list(following):
            if child.tag != qn("w:pPr"):
                paragraph.append(child)
        cell.remove(following)
    raw = "".join(d._xp(paragraph, ".//w:t/text()"))
    match = re.search(r"压力\s*损\s*失\s*等\s*级\s*[：:](\s*)(?=[△Δ]\s*p)", raw)
    if not match:
        return
    start, end = match.span(1)
    offset = 0
    for node in d._xp(paragraph, ".//w:t | .//w:br | .//w:tab"):
        if node.tag != qn("w:t"):
            if start <= offset <= end:
                node.getparent().remove(node)
            continue
        text = node.text or ""
        stop = offset + len(text)
        if stop > start and offset < end:
            d._write_slot_text(node, text[:max(0, start - offset)] + text[max(0, end - offset):])
        offset = stop


def _replace_pressure_value(cell, start, old, value):
    # Keep unchanged technical symbols in their original styled runs.
    for operation, left, right, new_left, new_right in reversed(
            SequenceMatcher(None, old, value).get_opcodes()):
        if operation == "equal":
            continue
        begin, end = start + left, start + right
        replacement = value[new_left:new_right]
        offset = 0
        inserted = False
        for node in d._xp(cell, ".//w:t"):
            text = node.text or ""
            stop = offset + len(text)
            if (stop > begin and offset < end) or (begin == end and offset <= begin <= stop):
                prefix = text[:max(0, begin - offset)]
                suffix = text[max(0, end - offset):]
                d._write_slot_text(node, prefix + (replacement if not inserted else "") + suffix)
                inserted = True
                if begin == end:
                    break
            offset = stop


def fill_observation_remarks(observe, outline, task, *, fill_basics=True):
    result = observation_values(outline, task)
    if task.get("report_type") == "燃气表":
        from gas_observation_flow import fill

        fill(observe, task)
    if task.get("report_type", "水表") != "水表":
        return result
    # Offsets are computed on the unchanged cell; edit from the end backwards.
    for cell, start, end, key, _old in sorted(_targets(observe), key=lambda item: item[1], reverse=True):
        value = result["values"].get(key, "")
        if not value:
            continue
        if start == end:
            offset = 0
            for node in d._xp(cell, ".//w:t"):
                text = node.text or ""
                if offset <= start <= offset + len(text):
                    split = start - offset
                    d._write_slot_text(node, text[:split] + value + text[split:])
                    break
                offset += len(text)
            else:
                d._put(cell, value)
        elif key == "pressure_loss":
            _replace_pressure_value(cell, start, _old, value)
        else:
            d._replace_text_span(cell, start, end, value)
    for cell, _start, _end, key, _old in _targets(observe):
        if key == "pressure_loss":
            _compact_pressure_loss(cell)
    if fill_basics:
        result["issues"].extend(fill_observation_basics(observe, task))
    return result
