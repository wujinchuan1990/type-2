"""Read numbered supplementary paragraphs without importing unnumbered boilerplate."""

import re
import unicodedata

from docx.oxml.ns import qn

import documents as d


NUMBER = re.compile(r"^\s*(?:([0-9０-９]+)[.．、)）](?![0-9０-９])|[（(]([0-9０-９]+)[）)])\s*")


def strip_number(text):
    return NUMBER.sub("", text, count=1)


def renumber_note(text, number):
    match = NUMBER.match(text)
    if match:
        start, end = match.span(1 if match[1] is not None else 2)
        return text[:start] + clause_number(text[start:end], number) + text[end:]
    return f"{number}." + text


def clause_number(original, number):
    digits = str(number)
    return digits.translate(str.maketrans("0123456789", "０１２３４５６７８９")) if all(
        "０" <= char <= "９" for char in original) else digits


def outline_note_ids(task, entries):
    return [f"outline:{task.get('outline_id') or 'current'}:{entry['number']}" for entry in entries]


def clause_order(task, source_ids):
    custom_ids = [note["id"] for note in task.get("custom_supplementary_notes", [])]
    available = [*source_ids, *custom_ids, "sealing"]
    submitted = task.get("other_information_order", [])
    if len(submitted) != len(set(submitted)) or len(available) != len(set(available)):
        raise ValueError("说明条款标识或顺序重复，无法可靠排列。")
    unknown = [key for key in submitted if key not in available and not key.startswith("outline:")]
    if unknown:
        raise ValueError("说明条款顺序包含不存在的条目，请重新打开任务核对。")
    retained = [key for key in submitted if key in available]
    new_sources = [key for key in source_ids if key not in retained]
    return [*new_sources, *retained, *(key for key in available if key not in retained and key not in new_sources)]


def _automatic_number(doc, paragraph, counters):
    properties = d._xp(paragraph, "./w:pPr/w:numPr")
    style = d.Paragraph(paragraph, doc).style
    while not properties and style is not None:
        properties = d._xp(style.element, "./w:pPr/w:numPr")
        style = style.base_style
    if not properties:
        return None
    ids = d._xp(properties[0], "./w:numId/@w:val")
    if not ids or ids == ["0"]:
        return None
    levels = d._xp(properties[0], "./w:ilvl/@w:val")
    level = int(levels[0]) if levels else 0
    root = doc.part.numbering_part.element
    numbers = d._xp(root, f"./w:num[@w:numId='{ids[0]}']")
    if len(numbers) != 1:
        raise ValueError("补充说明的Word自动编号定义缺失或不唯一。")
    abstract = d._xp(numbers[0], "./w:abstractNumId/@w:val")
    definitions = d._xp(root, f"./w:abstractNum[@w:abstractNumId='{abstract[0]}']/w:lvl[@w:ilvl='{level}']") if abstract else []
    override = d._xp(numbers[0], f"./w:lvlOverride[@w:ilvl='{level}']")
    overridden = d._xp(override[0], "./w:lvl") if override else []
    definitions = overridden or definitions
    if len(definitions) != 1:
        raise ValueError("补充说明的Word自动编号层级无法读取。")
    definition = definitions[0]
    fmt = d._xp(definition, "./w:numFmt/@w:val")
    label = d._xp(definition, "./w:lvlText/@w:val")
    starts = d._xp(override[0], "./w:startOverride/@w:val") if override else []
    starts = starts or d._xp(definition, "./w:start/@w:val") or ["1"]
    key = (ids[0], level)
    counters[key] = counters.get(key, int(starts[0]) - 1) + 1
    for child in list(counters):
        if child[0] == ids[0] and child[1] > level:
            del counters[child]
    if fmt != ["decimal"] or not label:
        return None
    marker = "%" + str(level + 1)
    if re.findall(r"%\d+", label[0]) != [marker]:
        raise ValueError("补充说明暂不支持多级组合条款号，请核对大纲编号。")
    prefix = label[0].replace(marker, str(counters[key]))
    if not NUMBER.match(prefix):
        raise ValueError("补充说明的自动条款号格式不支持，请核对大纲。")
    suffix = d._xp(definition, "./w:suff/@w:val") or ["tab"]
    if suffix[0] not in ("tab", "space", "nothing"):
        raise ValueError("补充说明的自动编号分隔符无法读取。")
    return counters[key], prefix + {"tab": "\t", "space": " ", "nothing": ""}[suffix[0]]


def read_supplementary_notes(outline, source="当前大纲"):
    result = {"entries": [], "source": source + " · 第7节补充说明", "issues": []}
    if outline is None:
        return result
    blocks = list(outline.element.body)
    headings = [i for i, block in enumerate(blocks) if block.tag == qn("w:p") and
                re.fullmatch(r"(?:7[.．、]?)?补充说明[：:]?", re.sub(r"\s+", "", d.Paragraph(block, outline).text))]
    if not headings:
        result["issues"].append("其他/大纲中未找到第7节补充说明，未提取附加条款，请核对。")
        return result
    if len(headings) != 1:
        raise ValueError("其他/大纲补充说明章节不唯一，无法确定复制范围。")
    start = headings[0]
    counters = {}
    last = 0
    for index, block in enumerate(blocks):
        if block.tag == qn("w:sectPr"):
            continue
        if block.tag != qn("w:p"):
            if index > start and d._text(block):
                raise ValueError("其他/补充说明含非段落内容，无法按编号条款完整复制，请核对。")
            continue
        text = d.Paragraph(block, outline).text
        if index > start and (
            re.match(r"^\s*(?:8[.．、]\s*(?:批准|签发|参考|附录|附件)|附录|附件)", text)
            or d._xp(block, "./w:pPr/w:outlineLvl")
            or d.Paragraph(block, outline).style.name.startswith("Heading")):
            break
        automatic = _automatic_number(outline, block, counters)
        if index <= start:
            continue
        match = NUMBER.match(text)
        prefix = ""
        if match:
            number = int(unicodedata.normalize("NFKC", match[1] or match[2]))
            content = text[match.end():]
        elif automatic:
            number, prefix = automatic
            content = text
        else:
            continue
        if number <= last or not content.strip():
            raise ValueError("其他/补充说明的条款号重复、倒序或内容为空，请核对大纲。")
        last = number
        result["entries"].append({"number": number, "text": prefix + text,
                                  "prefix": prefix, "element": block})
    return result
