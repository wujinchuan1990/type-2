"""Read component grids and fill only the report's existing photo appendix.

Inputs are already validated uploads; source documents are never saved here.
Missing assignments leave empty paragraphs. Inactive display assignments remain
in the task and produce a warning; stale/unknown assignments raise.
"""

import re
from copy import deepcopy
from pathlib import Path
from io import BytesIO

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Emu, Pt
from docx.text.paragraph import Paragraph
from PIL import Image
from contracts import (PHOTO_SLOTS, PHOTO_HEADINGS, PhotoEdit, PhotoLayout, active_photo_slots,
                       ordered_photo_slots, photo_slot_group, photo_layout_ids, photo_default_width,
                       GAS_PHOTO_REPORT_TYPES)
from photo_uploads import edited_photo_bytes, edited_photo_size


_GROUP_NAMES = {"测量传感器", "计算器", "指示", "指示装置", "辅助装置"}
_NAME_HEADERS = {"名称", "部件名称", "零部件名称"}


def _text(element):
    return "".join(element.itertext(qn("w:t"))).strip()


def _compact(value):
    return re.sub(r"\s+", "", value)


def _grid_rows(table):
    """Resolve physical cells to grid columns, retaining vertical owner identity."""
    grid = table.find(qn("w:tblGrid"))
    width = len(grid) if grid is not None else 0
    if not width:
        raise ValueError("关键零部件表缺少明确的表格网格")
    previous = {}
    rows = []
    for row_index, row in enumerate(table.findall(qn("w:tr")), 1):
        before = row.find("./" + qn("w:trPr") + "/" + qn("w:gridBefore"))
        column = int(before.get(qn("w:val"), "0")) if before is not None else 0
        current, physical = {}, []
        for cell in row.findall(qn("w:tc")):
            span = cell.find("./" + qn("w:tcPr") + "/" + qn("w:gridSpan"))
            span = int(span.get(qn("w:val"), "1")) if span is not None else 1
            if span < 1 or column < 0 or column + span > width:
                raise ValueError("关键零部件表第%d行网格跨度无效" % row_index)
            merge = cell.find("./" + qn("w:tcPr") + "/" + qn("w:vMerge"))
            if merge is not None and merge.get(qn("w:val")) != "restart":
                owners = [previous.get(c) for c in range(column, column + span)]
                if any(owner is None for owner in owners) or any(
                        owner is not owners[0] for owner in owners):
                    raise ValueError("关键零部件表第%d行纵向合并缺少唯一的起点" % row_index)
                owner = owners[0]
                if _text(cell) and _text(cell) != _text(owner):
                    raise ValueError("关键零部件表第%d行纵向合并名称冲突" % row_index)
            else:
                owner = cell
            for c in range(column, column + span):
                current[c] = owner
            physical.append((cell, column, column + span - 1))
            column += span
        # Only an explicit restart/continuation can own the next merged row.
        previous = {
            c: owner for cell, start, end in physical
            if cell.find("./" + qn("w:tcPr") + "/" + qn("w:vMerge")) is not None
            for c, owner in current.items() if start <= c <= end
        }
        rows.append((current, physical))
    return rows


def extract_component_names(tables) -> list[str]:
    """Return ordered leaf names under the rightmost grid column of 名称.

    Distinct cells with identical names remain distinct components. A vertically
    merged leaf is emitted once; a vertical group never replaces an empty leaf.
    Every supplied table must have an unambiguous name header and nonempty data.
    """
    result = []
    for table_index, table in enumerate(tables, 1):
        rows = _grid_rows(table)
        headers = [(r, start, end) for r, (_, cells) in enumerate(rows)
                   for cell, start, end in cells if _compact(_text(cell)) == "名称"]
        if not headers:
            raise ValueError("关键零部件表%d缺少“名称”表头" % table_index)
        columns = {end for _, _, end in headers}
        if len(columns) != 1:
            raise ValueError("关键零部件表%d的“名称”列不明确" % table_index)
        column = columns.pop()
        seen = set()
        names = []
        for row_index, (grid, _) in enumerate(rows[headers[0][0] + 1:],
                                              headers[0][0] + 2):
            owner = grid.get(column)
            if owner is None:
                raise ValueError("关键零部件表%d第%d行缺少名称单元格" %
                                 (table_index, row_index))
            name = _text(owner)
            if _compact(name) in _NAME_HEADERS:
                continue
            if not name or _compact(name) in _GROUP_NAMES:
                raise ValueError("关键零部件表%d第%d行的叶级名称为空或不明确" %
                                 (table_index, row_index))
            if owner not in seen:
                names.append(name)
                seen.add(owner)
        if not names:
            raise ValueError("关键零部件表%d没有叶级名称" % table_index)
        result.extend(names)
    if not result:
        raise ValueError("未提供可识别的关键零部件名称表")
    return result


def _unique_paragraph(body, text):
    matches = [p for p in body if p.tag == qn("w:p")
               and re.sub(r"^[一二三]、", "", _compact(_text(p))) == text]
    if len(matches) != 1:
        raise ValueError("附录2缺少或存在重复段落：" + text)
    return matches[0]


def _number_id(paragraph):
    return paragraph.xpath("./w:pPr/w:numPr/w:numId/@w:val")


def _appendix_region(doc):
    body = doc.element.body
    overall = _unique_paragraph(body, "样机的整体外形照片")
    display = _unique_paragraph(body, "样机的显示部分照片")
    components = _unique_paragraph(body, "样机的关键零部件照片")
    indices = [body.index(p) for p in (overall, display, components)]
    if not indices[0] < indices[1] < indices[2]:
        raise ValueError("附录2照片标题顺序不正确")
    between = list(body)[indices[0] + 1:indices[1]]
    captions = []
    for name in ("铭牌", "封样"):
        matches = [p for p in between if p.tag == qn("w:p") and _text(p) == name]
        if len(matches) != 1:
            raise ValueError("附录2缺少或存在重复的原始图注：" + name)
        captions.append(matches[0])
    if body.index(captions[0]) >= body.index(captions[1]):
        raise ValueError("附录2原始图注顺序不正确")
    first = components.getnext()
    if first is None or first.tag != qn("w:p") or not _number_id(first):
        raise ValueError("附录2缺少原始编号零部件段落")
    number_id = _number_id(first)
    end = indices[2] + 1
    for element in list(body)[end:]:
        if element.tag != qn("w:p") or element.xpath("./w:pPr/w:sectPr"):
            break
        if _text(element) and _number_id(element) != number_id:
            break
        end += 1
    region = list(body)[indices[0]:end]
    if any(p.xpath(".//w:sectPr") for p in region):
        raise ValueError("附录2照片区域内部包含分节符，不能安全替换")
    # Non-photo text in the fixed slots is not a disposable placeholder.
    known = {overall, display, components, first, *captions}
    for element in region:
        if element.tag != qn("w:p"):
            raise ValueError("附录2照片区域包含非段落内容")
        if (element not in known and body.index(element) < indices[2]
                and _text(element)):
            raise ValueError("附录2照片区域存在不能覆盖的内容：" + _text(element))
    blank = overall.getnext()
    if blank is None or blank.tag != qn("w:p") or _text(blank):
        raise ValueError("附录2缺少原始照片空白段落")
    return region, overall, display, components, captions, first, blank


def _clone_paragraph(source, doc, text=None):
    element = deepcopy(source)
    if text is not None:
        styles = source.xpath("./w:r/w:rPr")
        for child in list(element):
            if child.tag != qn("w:pPr"):
                element.remove(child)
        if text:
            run = OxmlElement("w:r")
            if styles:
                run.append(deepcopy(styles[0]))
            value = OxmlElement("w:t")
            value.text = text
            run.append(value)
            element.append(run)
    for marker in element.xpath(".//w:lastRenderedPageBreak"):
        marker.getparent().remove(marker)
    return Paragraph(element, doc._body)


def _new_numbering(doc, original):
    """Restart a new list without changing any list outside this appendix."""
    root = doc.part.numbering_part.element
    source = next((n for n in root.findall(qn("w:num"))
                   if n.get(qn("w:numId")) == original), None)
    if source is None:
        raise ValueError("附录2零部件编号定义不存在")
    existing_override = source.find("./" + qn("w:lvlOverride") + '[@' + qn("w:ilvl") + '="0"]')
    level = existing_override.find(qn("w:lvl")) if existing_override is not None else None
    if level is None:
        reference = source.find(qn("w:abstractNumId"))
        abstract = next((n for n in root.findall(qn("w:abstractNum"))
                         if reference is not None and
                         n.get(qn("w:abstractNumId")) == reference.get(qn("w:val"))), None)
        level = (next((n for n in abstract.findall(qn("w:lvl"))
                       if n.get(qn("w:ilvl")) == "0"), None) if abstract is not None else None)
    if level is None or level.find(qn("w:lvlText")) is None:
        raise ValueError("附录2零部件编号缺少第一级格式")
    level = deepcopy(level)
    suffix = level.find(qn("w:suff"))
    if suffix is None:
        suffix = OxmlElement("w:suff")
        level.insert(list(level).index(level.find(qn("w:lvlText"))), suffix)
    suffix.set(qn("w:val"), "space")
    properties = level.find(qn("w:pPr"))
    if properties is None:
        properties = OxmlElement("w:pPr")
        run_properties = level.find(qn("w:rPr"))
        level.insert(list(level).index(run_properties) if run_properties is not None else len(level), properties)
    for old in properties.findall(qn("w:ind")):
        properties.remove(old)
    indentation = OxmlElement("w:ind")
    indentation.set(qn("w:left"), "0")
    indentation.set(qn("w:hanging"), "0")
    properties.append(indentation)
    new_id = max(int(n.get(qn("w:numId"))) for n in root.findall(qn("w:num"))) + 1
    number = deepcopy(source)
    number.set(qn("w:numId"), str(new_id))
    for override in number.findall(qn("w:lvlOverride")):
        if override.get(qn("w:ilvl")) == "0":
            number.remove(override)
    override = OxmlElement("w:lvlOverride")
    override.set(qn("w:ilvl"), "0")
    start = OxmlElement("w:startOverride")
    start.set(qn("w:val"), "1")
    override.append(start)
    override.append(level)
    number.append(override)
    root.append(number)
    return str(new_id)


def _usable_height(region, axis="h"):
    for element in region[-1].itersiblings():
        section = (element if element.tag == qn("w:sectPr")
                   else element.find("./" + qn("w:pPr") + "/" + qn("w:sectPr")))
        if section is not None:
            size, margin = section.find(qn("w:pgSz")), section.find(qn("w:pgMar"))
            if size is not None and margin is not None:
                first, last = ("top", "bottom") if axis == "h" else ("left", "right")
                return (int(size.get(qn("w:" + axis))) -
                        abs(int(margin.get(qn("w:" + first)))) -
                        abs(int(margin.get(qn("w:" + last))))) * 635
    raise ValueError("附录2缺少可确定的页面高度和页边距")


def photo_width_limit(doc):
    # Round down to 0.01 cm so the displayed upper bound is safe to enter.
    return _usable_height(_appendix_region(doc)[0], "w") // 3600 / 100


def fill_attachment_photos(doc, photos, components, photo_paths, photo_edits=None,
                           *, display_types=None, photo_order=None, photo_layout=None,
                           layout_targets=None, report_type="水表") -> list[str]:
    """Fill the existing appendix in-place and return missing/geometry warnings.

    ``photos`` maps fixed slots or component IDs to upload IDs; ``photo_paths``
    maps those IDs to verified JPEG/PNG Paths. Components are ordered ``id/name``
    dictionaries. Unknown slots, malformed components and missing active photo
    files raise before replacing body content. No section or source is saved.
    Image sizes use pixels, never potentially anisotropic image DPI.
    Only a mechanical/electronic selection uses the paired display slots;
    inactive fixed assignments and their edits remain untouched and unrendered.
    """
    gas = report_type in GAS_PHOTO_REPORT_TYPES
    active_slots = active_photo_slots(display_types, report_type)
    if not gas and any(slot.startswith("sealed_overall:") for slot in photos):
        raise ValueError("仅燃气表和流量计（气体）支持多张封样照片")
    layouts = {key: PhotoLayout.model_validate(value) for key, value in (photo_layout or {}).items()}
    if set(layouts) - photo_layout_ids(photos, components, display_types, report_type):
        raise ValueError("照片排版设置包含失效位置，不能应用到其他照片或标题")
    inactive_slots = set(PHOTO_SLOTS) - set(active_slots)
    edits = {slot: PhotoEdit.model_validate(edit)
             for slot, edit in (photo_edits or {}).items()}
    if set(edits) - set(photos):
        raise ValueError("照片编辑设置没有对应照片，请移除旧编辑设置")
    maximum_width = photo_width_limit(doc)
    for slot, edit in edits.items():
        if photo_slot_group(slot) in inactive_slots:
            continue
        if edit.width_cm is not None and edit.width_cm > maximum_width:
            raise ValueError("照片输出宽度超出附件2可用页面宽度（%.2fcm）" % maximum_width)
    component_ids = []
    for component in components:
        identifier, name = component.get("id", ""), component.get("name", "")
        if (not isinstance(identifier, str) or
                not re.fullmatch(r"component:[^:]+:[1-9]\d*", identifier) or
                not isinstance(name, str) or not name.strip()):
            raise ValueError("关键零部件必须包含有效的稳定ID和非空名称")
        if identifier in component_ids:
            raise ValueError("关键零部件ID重复：" + identifier)
        component_ids.append(identifier)
    slots = set(PHOTO_SLOTS) | set(component_ids)
    unknown = {slot for slot in photos if photo_slot_group(slot) not in slots}
    if unknown:
        raise ValueError("未知或已失效的照片位置：" + "、".join(sorted(unknown)))
    paths, dimensions = {}, {}
    for slot, photo_id in photos.items():
        if not isinstance(photo_id, str) or not photo_id:
            raise ValueError("照片位置没有有效的上传ID：" + slot)
        if photo_slot_group(slot) in inactive_slots:
            continue
        if photo_id not in photo_paths:
            raise FileNotFoundError("照片上传文件不存在：" + photo_id)
        path = Path(photo_paths[photo_id])
        if not path.is_file():
            raise FileNotFoundError("照片上传文件不存在：" + photo_id)
        with Image.open(path) as image:
            if image.format not in ("JPEG", "PNG"):
                raise ValueError("照片仅支持JPEG/PNG：" + photo_id)
            dimensions[slot] = edited_photo_size(image.size, edits.get(slot, PhotoEdit()))
        paths[slot] = path
    region, overall, display, heading, captions, name_style, blank = _appendix_region(doc)
    usable = _usable_height(region)
    warnings = [
        "附录2照片%s（%s）不适用于当前报告对象或显示型式；已保留在任务中，本次不输出" %
        (slot, PHOTO_SLOTS[slot][0])
        for slot in PHOTO_SLOTS if slot in inactive_slots and slot in photos
    ]
    output, heights = [], {}
    layout_spacers = set()
    next_drawing_id = doc.part.next_id

    def apply_layout(paragraph, target, label, default_blanks=0, height=18):
        layout = layouts.get(target, PhotoLayout())
        blanks = default_blanks if layout.blank_lines is None else layout.blank_lines
        blank_line_spacing = 1.5 if default_blanks else (2.0 if gas else 1.5)
        for _ in range(blanks):
            spacer = photo(None, 0)
            spacer.paragraph_format.line_spacing = blank_line_spacing
            layout_spacers.add(spacer._p)
        if layout_targets is not None:
            layout_targets.append({
                "id": target, "label": label, "paragraph": paragraph,
                "height": height, "blank_lines": blanks, "page_break_before": False,
                "default_blank_lines": default_blanks, "default_page_break_before": False,
                "blank_line_spacing": blank_line_spacing,
            })

    def photo(slot, width_cm, default_blanks=0):
        nonlocal next_drawing_id
        paragraph = _clone_paragraph(blank, doc, "")
        fmt = paragraph.paragraph_format
        fmt.alignment = WD_ALIGN_PARAGRAPH.CENTER
        fmt.space_before = Pt(0)
        fmt.space_after = Pt(0)
        fmt.line_spacing = 1.0
        fmt.keep_with_next = False
        fmt.keep_together = True
        fmt.page_break_before = False
        for prop in paragraph._p.xpath(
                "./w:pPr/w:numPr | ./w:pPr/w:ind | ./w:pPr/w:snapToGrid"):
            prop.getparent().remove(prop)
        snap = OxmlElement("w:snapToGrid")
        snap.set(qn("w:val"), "0")
        paragraph._p.get_or_add_pPr().insert_element_before(
            snap, "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents",
            "w:suppressOverlap", "w:jc", "w:textDirection", "w:textAlignment",
            "w:textboxTightWrap", "w:outlineLvl", "w:divId", "w:cnfStyle", "w:rPr",
            "w:sectPr", "w:pPrChange")
        if slot is not None:
            if slot not in paths:
                warnings.append("附录2缺少照片：%s（保留空白位置）" % slot)
            else:
                edit = edits.get(slot, PhotoEdit())
                if gas:
                    width_cm = photo_default_width(slot, report_type)
                width_cm = edit.width_cm if edit.width_cm is not None else width_cm
                if width_cm > maximum_width:
                    raise ValueError("照片输出宽度超出附件2可用页面宽度（%.2fcm）" % maximum_width)
                width = int(Cm(width_cm))
                pixel_width, pixel_height = dimensions[slot]
                height = round(width * pixel_height / pixel_width)
                image_source = (BytesIO(edited_photo_bytes(paths[slot], edit))
                                if edit.rotation or edit.crop else str(paths[slot]))
                shape = paragraph.add_run().add_picture(image_source, Emu(width), Emu(height))
                # Detached paragraphs are not visible to python-docx's ID allocator.
                shape._inline.docPr.id = next_drawing_id
                next_drawing_id += 1
                heights[slot] = height
                if height > usable:
                    warnings.append("附录2照片%s按要求宽度的高度超出可用页面高度；"
                                    "未缩小或裁剪，需人工检查分页" % slot)
        if slot is not None:
            group = photo_slot_group(slot)
            label = (PHOTO_SLOTS[group][0] if group in PHOTO_SLOTS else
                     next(item["name"] for item in components if item["id"] == slot))
            if group in ("sample", "display", "display_mechanical", "display_electronic") or (gas and group == "sealed_overall"):
                group_slots = ordered_photo_slots(photos, photo_order, group)
                if slot in group_slots:
                    label += " " + str(group_slots.index(slot) + 1)
            apply_layout(paragraph, "photo:" + slot, label, default_blanks,
                         heights.get(slot, int(Pt(18))) / 12700)
        output.append(paragraph._p)
        return paragraph

    def heading_paragraph(source, target, text=None):
        paragraph = _clone_paragraph(source, doc, text)
        paragraph.paragraph_format.keep_with_next = True
        paragraph.paragraph_format.keep_together = True
        paragraph.paragraph_format.page_break_before = False
        apply_layout(paragraph, target, PHOTO_HEADINGS[target])
        output.append(paragraph._p)

    def original_caption(source):
        paragraph = _clone_paragraph(source, doc)
        paragraph.paragraph_format.keep_with_next = False
        paragraph.paragraph_format.keep_together = True
        output.append(paragraph._p)

    heading_paragraph(overall, "heading:overall")
    for index, slot in enumerate(ordered_photo_slots(photos, photo_order, "sample") or ["sample"]):
        photo(slot, 12, default_blanks=int(index > 0))
    if gas:
        for slot in ordered_photo_slots(photos, photo_order, "sealed_overall") or ["sealed_overall"]:
            photo(slot, 12, default_blanks=1)
    else:
        photo(None, 0)
        photo("nameplate", 8).paragraph_format.keep_with_next = True
        original_caption(captions[0])
        photo("sealed_overall", 12)
        photo(None, 0)
        photo("sealed_detail", 8).paragraph_format.keep_with_next = True
        original_caption(captions[1])
    heading_paragraph(display, "heading:internal", "二、样机的内部结构照片")
    photo("internal", 8)
    heading_paragraph(display, "heading:display", "三、样机的显示部分照片")
    if "display_mechanical" in active_slots:
        display_count = 0
        for group, label in (("display_mechanical", "机械部分"), ("display_electronic", "电子部分")):
            for index, slot in enumerate(ordered_photo_slots(photos, photo_order, group) or [group]):
                image_paragraph = photo(slot, 8, default_blanks=int(index > 0 or (gas and display_count > 0)))
                if slot in paths:
                    display_count += 1
                image_paragraph.paragraph_format.keep_with_next = True
                caption = _clone_paragraph(captions[0], doc, label)
                caption.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
                caption.paragraph_format.keep_with_next = False
                caption.paragraph_format.keep_together = True
                caption.paragraph_format.page_break_before = False
                for prop in caption._p.xpath("./w:pPr/w:numPr | ./w:pPr/w:ind"):
                    prop.getparent().remove(prop)
                output.append(caption._p)
    else:
        for index, slot in enumerate(ordered_photo_slots(photos, photo_order, "display") or ["display"]):
            photo(slot, 8, default_blanks=int(index > 0))
    heading_paragraph(heading, "heading:components", text="四、样机的关键零部件照片")
    number_id = _new_numbering(doc, _number_id(name_style)[0]) if components else None
    for component in components:
        paragraph = _clone_paragraph(name_style, doc, component["name"].strip())
        paragraph._p.xpath("./w:pPr/w:numPr/w:numId")[0].set(qn("w:val"), number_id)
        paragraph.paragraph_format.keep_with_next = True
        paragraph.paragraph_format.keep_together = True
        paragraph.paragraph_format.page_break_before = False
        for indentation in paragraph._p.xpath("./w:pPr/w:ind"):
            indentation.getparent().remove(indentation)
        apply_layout(paragraph, "heading:" + component["id"], component["name"])
        output.append(paragraph._p)
        photo(component["id"], 8)
    # Override template/style and legacy task breaks only inside this appendix.
    # Carry a heading's keep-next chain through any user-inserted blank lines.
    keep_next = False
    for element in output:
        paragraph = Paragraph(element, doc._body)
        paragraph.paragraph_format.page_break_before = False
        for marker in element.xpath(".//w:br[@w:type='page'] | .//w:lastRenderedPageBreak"):
            marker.getparent().remove(marker)
        if element in layout_spacers:
            paragraph.paragraph_format.keep_with_next = keep_next
        keep_next = bool(paragraph.paragraph_format.keep_with_next)
    body = doc.element.body
    index = body.index(region[0])
    for element in region:
        body.remove(element)
    for offset, element in enumerate(output):
        body.insert(index + offset, element)
    return warnings
