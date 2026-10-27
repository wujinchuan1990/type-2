"""Editable whole-section extraction, verified against native Word pages."""

from copy import deepcopy
import re
from zipfile import ZipFile, ZIP_DEFLATED

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from lxml import etree
from PIL import Image, ImageChops, ImageStat
import pymupdf

import documents as d


FOLDERS = {"observe": "observe", "items": "type evaluation items", "result": "type evaluation results"}
EXPORT_REPORT_NUMBER = "2026E31-40-000001"


def fixed_header_number(data, number):
    root = etree.fromstring(data)
    pattern = re.compile(r"\d{4}E\d+\s*[-－–—:：]\s*\d+\s*[-－–—]\s*\d+(?:[-－–—]\d+)*")
    replaced = 0
    for paragraph in d._xp(root, ".//w:p"):
        pieces, field_depth = [], 0
        for node in paragraph.iter():
            if node.tag == qn("w:fldChar"):
                kind = node.get(qn("w:fldCharType"))
                if kind == "begin":
                    field_depth += 1
                elif kind == "end":
                    field_depth = max(0, field_depth - 1)
            elif node.tag == qn("w:t"):
                value = node.text or ""
                in_field = field_depth or any(parent.tag == qn("w:fldSimple")
                                              for parent in node.iterancestors())
                pieces.append("\0" * len(value) if in_field else value)
        text = "".join(pieces)
        matches = list(pattern.finditer(text))
        if not matches:
            continue
        instructions = d._xp(paragraph, ".//w:instrText/text() | .//w:fldSimple/@w:instr")
        page_fields_only = instructions and all(
            re.fullmatch(r"\s*(?:PAGE|NUMPAGES)(?:\s+\\\*\s+\w+)?\s*", instruction, re.I)
            for instruction in instructions)
        if (d._xp(paragraph, ".//w:ins | .//w:del")
                or (d._xp(paragraph, ".//w:fldChar | .//w:fldSimple") and not page_fields_only)):
            raise ValueError("拆分页眉报告编号含域或修订，不能安全统一编号。")
        for match in reversed(matches):
            d._replace_text_span(paragraph, *match.span(), number)
            replaced += 1
    if re.search(r"(?:报告|证书)\s*编号", d._text(root)) and not replaced:
        raise ValueError("拆分页眉报告编号格式无法识别，停止导出，不保留旧编号。")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def paragraphs(block):
    return [block] if block.tag == qn("w:p") else d._xp(block, ".//w:p")


def numbering_properties(paragraph, styles):
    current = paragraph.find(qn("w:pPr"))
    seen = set()
    while current is not None:
        num = current.find(qn("w:numPr"))
        if num is not None and num.find(qn("w:numId")) is not None:
            return None if num.find(qn("w:numId")).get(qn("w:val")) == "0" else num
        style = current.find(qn("w:pStyle"))
        if style is None:
            return None
        identity = style.get(qn("w:val"))
        if identity in seen:
            raise ValueError("源报告样式继承循环，无法保留自动编号。")
        seen.add(identity)
        source = styles.get(identity)
        if source is None:
            return None
        current = deepcopy(source.find(qn("w:pPr")))
        parent = source.find(qn("w:basedOn"))
        if parent is not None:
            if current is None:
                current = OxmlElement("w:pPr")
            link = OxmlElement("w:pStyle")
            link.set(qn("w:val"), parent.get(qn("w:val")))
            current.append(link)
    return None


def discover(source):
    doc = Document(source)
    blocks = list(doc.element.body)

    def anchor(phrase):
        hits = [i for i, block in enumerate(blocks)
                if block.tag == qn("w:p") and phrase in d._norm(d._text(block))]
        if len(hits) != 1:
            raise ValueError(f"无法唯一定位“{phrase}”，请核对报告章节；不能猜测拆分范围。")
        return hits[0]

    units = []

    def add(kind, start, end, project):
        if start >= end:
            raise ValueError("源报告章节顺序不正确。")
        while end > start + 1 and not d._text(blocks[end - 1]) and not d._xp(
                blocks[end - 1], ".//w:drawing | .//w:pict | .//w:fldChar"):
            end -= 1
        units.append({"id": f"part-{len(units) + 1}", "kind": kind, "project": project,
                      "start_block": start, "end_block": end})

    add("items", anchor("型式评价项目及评价结果一览表"), anchor("审查的技术资料及结论"), "项目一览表")
    add("observe", anchor("样机的基本信息"), anchor("试验项目记录"), "观察记录")
    results = anchor("型式评价结果摘要")
    photo_hits = [i for i, block in enumerate(blocks) if i > results and block.tag == qn("w:p")
                  and re.search(r"样机的整体(?:外形)?照片", d._norm(d._text(block)))]
    if len(photo_hits) != 1:
        raise ValueError("未唯一定位附件2整体照片标题，不能确定结果部分的结束位置。")
    stop = photo_hits[0]
    titles = [i for i in range(results + 1, stop) if blocks[i].tag == qn("w:p")
              and re.match(r"^型式评价结果\s*[-—－–]", d._text(blocks[i]))]
    if not titles:
        raise ValueError("未找到独立的型式评价结果项目标题，不能按表格行猜分项目。")
    for start, end in zip(titles, titles[1:] + [stop]):
        names, numbers = [], []
        for table in blocks[start:end]:
            if table.tag != qn("w:tbl"):
                continue
            rows = list(d._effective_rows(table))
            if not rows or [d._norm(d._text(c)) for c in rows[0][0]][:2] != ["序号", "项目名称"]:
                continue
            for cells, _owners in rows[1:]:
                if len(cells) < 2:
                    continue
                number, name = d._norm(d._text(cells[0])), re.sub(r"\s+", "", d._text(cells[1]))
                if number.isdigit() and number not in numbers:
                    numbers.append(number)
                if name and name != "项目名称" and name not in names:
                    names.append(name)
        if not names:
            raise ValueError("结果项目缺少可识别的项目名称表头，不能生成可靠文件名。")
        add("result", start, end, "、".join(names))
        units[-1]["sequence"] = numbers[0] if numbers else ""
    units[2]["start_block"] = results
    count = None
    for block in blocks[:units[0]["start_block"]]:
        if block.tag != qn("w:tbl"):
            continue
        rows = list(d._effective_rows(block))
        if not rows:
            continue
        head = [d._norm(d._text(c)) for c in rows[0][0]]
        if "编号" in head and any("型号" in c for c in head) and any("准确度" in c for c in head):
            numbers = {d._text(cells[0]) for cells, _ in rows[1:] if d._text(cells[0]).isdigit()}
            if numbers:
                if count is not None and count != len(numbers):
                    raise ValueError("样机组数量存在冲突，请核对源报告。")
                count = len(numbers)
    return units, count


def prepare_marked(source, output, units):
    doc = Document(source)
    blocks = list(doc.element.body)
    styles = {s.get(qn("w:styleId")): s for s in doc.styles.element}
    existing = [int(x) for x in d._xp(doc.element, ".//w:bookmarkStart/@w:id") if x.isdigit()]
    existing_names = set(d._xp(doc.element, ".//w:bookmarkStart/@w:name"))
    next_id = max(existing, default=0) + 1
    names, numbered = [], []

    def mark(p, name, at_end=False, allow_empty=False):
        nonlocal next_id
        if name in existing_names or name in names:
            raise ValueError("源报告已含同名定位书签，不能覆盖或使用含糊的页码定位。")
        begin, end = OxmlElement("w:bookmarkStart"), OxmlElement("w:bookmarkEnd")
        begin.set(qn("w:id"), str(next_id))
        begin.set(qn("w:name"), name)
        end.set(qn("w:id"), str(next_id))
        next_id += 1
        texts = [node for node in d._xp(p, ".//w:t") if node.text]
        if not texts:
            if allow_empty:
                p.append(begin)
                p.append(end)
                names.append(name)
                return
            raise ValueError("自动分页定位段落没有可定位文字，需人工核对。")
        text = texts[-1] if at_end else texts[0]
        run = text.getparent()
        if run.tag != qn("w:r"):
            raise ValueError("自动分页定位遇到不支持的文字容器。")
        character = len(text.text) - 1 if at_end else 0
        pieces = [OxmlElement("w:r") for _ in range(3)]
        style = run.find(qn("w:rPr"))
        for piece in pieces:
            if style is not None:
                piece.append(deepcopy(style))
        index = list(run).index(text)
        for node in list(run)[:index]:
            if node.tag != qn("w:rPr"):
                pieces[0].append(deepcopy(node))
        for piece, content in zip(pieces, (text.text[:character], text.text[character], text.text[character + 1:])):
            node = deepcopy(text)
            node.text = content
            node.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
            piece.append(node)
        for node in list(run)[index + 1:]:
            pieces[2].append(deepcopy(node))
        parent, position = run.getparent(), run.getparent().index(run)
        parent.remove(run)
        for offset, node in enumerate((pieces[0], begin, pieces[1], end, pieces[2])):
            parent.insert(position + offset, node)
        names.append(name)

    for index, unit in enumerate(units):
        selected = blocks[unit["start_block"]:unit["end_block"]]
        ps = [p for b in selected for p in paragraphs(b) if d._text(p) or d._xp(p, ".//w:drawing")]
        if not ps:
            raise ValueError("选定区域没有可编辑内容。")
        unit["start_marker"], unit["end_marker"] = f"sep_s{index}", f"sep_e{index}"
        mark(ps[0], unit["start_marker"])
        mark(ps[-1], unit["end_marker"], at_end=True)
        unit["leading_markers"] = []
        for block_index in range(unit["start_block"] - 1, -1, -1):
            block = blocks[block_index]
            if (block.tag != qn("w:p") or d._text(block).strip()
                    or d._xp(block, ".//w:drawing | .//w:fldChar | .//w:sectPr")):
                break
            name = f"sep_lead{index}_{block_index}"
            mark(block, name, allow_empty=True)
            unit["leading_markers"].append({"block": block_index, "marker": name})
        after = next((p for block in blocks[unit["end_block"]:] for p in paragraphs(block)
                      if d._text(p).strip()), None)
        if after is not None:
            unit["after_marker"] = f"sep_after{index}"
            mark(after, unit["after_marker"])
    for block_index, block in enumerate(blocks):
        for paragraph_index, p in enumerate(paragraphs(block)):
            num = numbering_properties(p, styles)
            if num is not None:
                name = f"sepn_{block_index}_{paragraph_index}"
                mark(p, name)
                level = num.find(qn("w:ilvl"))
                numbered.append({"block": block_index, "paragraph": paragraph_index, "marker": name,
                                 "num_id": num.find(qn("w:numId")).get(qn("w:val")),
                                 "level": int(level.get(qn("w:val"))) if level is not None else 0})
    if len(names) > 1500:
        raise ValueError("报告自动编号定位点过多，请分批核对，不能跳过实际分页。")
    doc.save(output)
    return names, numbered


def default_filename(unit, count):
    suffix = f"{count} {'model' if count == 1 else 'models'}" if count else "待填写型号数量"
    if unit["kind"] == "items":
        return "type evaluation items.docx"
    if unit["kind"] == "observe":
        return f"observe-{suffix}.docx"
    name = re.sub(r'[/\\:*?"<>|\x00-\x1f]', "_", unit["project"])
    return f"{unit.get('sequence') or '待核对序号'}.{name}-{suffix}.docx"


def effective_sections(body):
    sections, previous = [], {}
    begin = 0
    for i, block in enumerate(body):
        props = block if block.tag == qn("w:sectPr") else block.find("./" + qn("w:pPr") + "/" + qn("w:sectPr"))
        if props is None:
            continue
        resolved = deepcopy(props)
        for kind in ("headerReference", "footerReference"):
            own = {ref.get(qn("w:type")): ref for ref in resolved.findall(qn("w:" + kind))}
            for typ, ref in previous.get(kind, {}).items():
                if typ not in own:
                    resolved.insert(0, deepcopy(ref))
            previous[kind] = {ref.get(qn("w:type")): deepcopy(ref) for ref in resolved.findall(qn("w:" + kind))}
        sections.append((begin, i, resolved))
        begin = i + 1
    return sections


def fixed_totals(data, total):
    root = etree.fromstring(data)
    for simple in d._xp(root, ".//w:fldSimple"):
        if simple.get(qn("w:instr"), "").strip().upper().startswith("NUMPAGES"):
            texts = d._xp(simple, ".//w:t")
            if not texts:
                raise ValueError("总页数域缺少显示值，无法保留原页脚。")
            texts[0].text = str(total)
            for text in texts[1:]:
                text.text = ""
            parent, position = simple.getparent(), simple.getparent().index(simple)
            for child in list(simple):
                parent.insert(position, deepcopy(child))
                position += 1
            parent.remove(simple)
    for p in d._xp(root, ".//w:p"):
        active, instruction = None, ""
        for run in p.findall(qn("w:r")):
            marker = run.find(qn("w:fldChar"))
            kind = marker.get(qn("w:fldCharType")) if marker is not None else None
            if kind == "begin":
                if active is not None:
                    raise ValueError("嵌套页脚域需要人工核对，不能扁平化。")
                active, instruction = [], ""
            if active is not None:
                active.append(run)
                instruction += "".join(d._xp(run, "./w:instrText/text()"))
            if kind == "end" and active is not None:
                if instruction.strip().upper().startswith("NUMPAGES"):
                    texts = [text for item in active for text in item.findall(qn("w:t"))]
                    if not texts:
                        raise ValueError("总页数域缺少缓存值，不能保留原页脚。")
                    texts[0].text = str(total)
                    for text in texts[1:]:
                        text.text = ""
                    for item in active:
                        for node in d._xp(item, "./w:fldChar | ./w:instrText"):
                            item.remove(node)
                active = None
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def extract(source, output, start, end, printed_page, total_pages, numbered, bookmark_data,
            freeze_total_pages=True, report_number=None):
    """Copy original package parts; only resolve section/page/list context."""
    with ZipFile(source) as archive:
        root = etree.fromstring(archive.read("word/document.xml"))
        body = root.find(qn("w:body"))
        if body is None or not 0 <= start < end < len(body):
            raise ValueError("选定正文范围无效。")
        sections = effective_sections(body)
        first = next((s for s in sections if s[0] <= start <= s[1]), None)
        last = next((s for s in sections if s[0] <= end - 1 <= s[1]), None)
        if first is None or last is None:
            raise ValueError("无法恢复源报告分节设置。")
        selected = list(body)[start:end]
        trailing_section = selected[-1].find("./" + qn("w:pPr") + "/" + qn("w:sectPr"))
        if trailing_section is not None:
            trailing_section.getparent().remove(trailing_section)
        for node in list(selected[0].iter()):
            if node.tag == qn("w:t") and (node.text or "").strip():
                break
            if node.tag == qn("w:br") and node.get(qn("w:type")) == "page":
                node.getparent().remove(node)
        final = deepcopy(last[2])
        first_props = final
        for section in sections:
            if start <= section[1] < end - 1:
                props = selected[section[1] - start].find("./" + qn("w:pPr") + "/" + qn("w:sectPr"))
                if props is not None:
                    props.getparent().replace(props, deepcopy(section[2]))
                    if section is first:
                        first_props = selected[section[1] - start].find("./" + qn("w:pPr") + "/" + qn("w:sectPr"))
        if start > first[0]:
            title = first_props.find(qn("w:titlePg"))
            if title is not None:
                first_props.remove(title)
        page = first_props.find(qn("w:pgNumType"))
        if page is None:
            page = OxmlElement("w:pgNumType")
            following = next((n for n in first_props if n.tag in {
                qn("w:cols"), qn("w:formProt"), qn("w:vAlign"), qn("w:noEndnote"),
                qn("w:titlePg"), qn("w:textDirection"), qn("w:docGrid")}), None)
            first_props.insert(list(first_props).index(following) if following is not None else len(first_props), page)
        page.set(qn("w:start"), str(printed_page))
        numbering = etree.fromstring(archive.read("word/numbering.xml")) if "word/numbering.xml" in archive.namelist() else None
        next_num = max((int(n.get(qn("w:numId"))) for n in numbering.findall(qn("w:num"))), default=0) + 1 if numbering is not None else 1
        contexts = {}
        for item in numbered:
            original_id, level = item["num_id"], item["level"]
            value = bookmark_data[item["marker"]]["number"]
            contexts.setdefault(original_id, {})[level] = value
            if not start <= item["block"] < end:
                continue
            if numbering is None:
                raise ValueError("源报告缺少自动编号定义。")
            original = next((n for n in numbering.findall(qn("w:num")) if n.get(qn("w:numId")) == original_id), None)
            if original is None or value < 1:
                raise ValueError("源报告自动编号无法核对，停止拆分。")
            num = deepcopy(original)
            num.set(qn("w:numId"), str(next_num))
            for key in list(num.attrib):
                if key != qn("w:numId"):
                    del num.attrib[key]
            for ilvl, start_value in contexts[original_id].items():
                if ilvl > level:
                    continue
                override = next((n for n in num.findall(qn("w:lvlOverride"))
                                 if n.get(qn("w:ilvl")) == str(ilvl)), None)
                if override is None:
                    override = OxmlElement("w:lvlOverride")
                    override.set(qn("w:ilvl"), str(ilvl))
                    num.append(override)
                start_override = override.find(qn("w:startOverride"))
                if start_override is None:
                    start_override = OxmlElement("w:startOverride")
                    override.insert(0, start_override)
                start_override.set(qn("w:val"), str(start_value))
            numbering.append(num)
            p = paragraphs(selected[item["block"] - start])[item["paragraph"]]
            props = p.find(qn("w:pPr"))
            if props is None:
                props = OxmlElement("w:pPr")
                p.insert(0, props)
            old = props.find(qn("w:numPr"))
            position = list(props).index(old) if old is not None else (1 if props.find(qn("w:pStyle")) is not None else 0)
            if old is not None:
                props.remove(old)
            new = OxmlElement("w:numPr")
            for tag, value in (("ilvl", level), ("numId", next_num)):
                child = OxmlElement("w:" + tag)
                child.set(qn("w:val"), str(value))
                new.append(child)
            props.insert(position, new)
            next_num += 1
        for block in list(body):
            body.remove(block)
        for block in selected:
            for marker in d._xp(block, ".//w:lastRenderedPageBreak"):
                marker.getparent().remove(marker)
            body.append(block)
        if body[-1].tag == qn("w:tbl"):
            p, props, spacing = OxmlElement("w:p"), OxmlElement("w:pPr"), OxmlElement("w:spacing")
            for key, value in {"before": "0", "after": "0", "line": "1", "lineRule": "exact"}.items():
                spacing.set(qn("w:" + key), value)
            props.append(spacing)
            p.append(props)
            body.append(p)
        body.append(final)
        with ZipFile(output, "x", compression=ZIP_DEFLATED) as target:
            for info in archive.infolist():
                data = archive.read(info.filename)
                if info.filename == "word/document.xml":
                    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                elif info.filename == "word/numbering.xml" and numbering is not None:
                    data = etree.tostring(numbering, xml_declaration=True, encoding="UTF-8", standalone=True)
                elif freeze_total_pages and re.fullmatch(r"word/(?:footer|header)\d*\.xml", info.filename):
                    data = fixed_totals(data, total_pages)
                if report_number is not None and re.fullmatch(r"word/header\d*\.xml", info.filename):
                    data = fixed_header_number(data, report_number)
                target.writestr(info, data)


def compare_pages(reference, output, start_page, end_page, diagnostic_dir, end_y=None):
    """Compare the selected source region; require omitted regions to remain blank."""
    checks = []
    with pymupdf.open(reference) as source, pymupdf.open(output) as result:
        if len(result) != end_page - start_page + 1:
            raise ValueError(f"拆分后页数{len(result)}与源范围{end_page - start_page + 1}不一致，停止交付。")
        for index, page in enumerate(result):
            original = source[start_page - 1 + index]
            boundary = end_y if index == len(result) - 1 else None
            if boundary is not None and not 0 < boundary < original.rect.height * .85:
                raise ValueError("同页章节边界无效，不能排除无关章节。")
            footer_top = min((b[1] for b in original.get_text("blocks")
                              if b[1] > original.rect.height * .85 and re.search(r"第\s*\d+\s*页", b[4])),
                             default=original.rect.height)
            source_text = original.get_text(sort=True)
            if boundary is not None:
                source_text = (original.get_text(clip=(0, 0, original.rect.width, boundary), sort=True)
                               + original.get_text(clip=(0, footer_top, original.rect.width, original.rect.height), sort=True))
            normalize = lambda text: re.sub(r"\s+", "", text)
            if normalize(source_text) != normalize(page.get_text(sort=True)):
                raise ValueError(f"原报告P{start_page + index}的文字/数据或页眉页脚不一致，停止交付。")
            if original.rect != page.rect:
                raise ValueError("拆分后纸张尺寸改变，停止交付。")
            images = []
            for current in (original, page):
                pix = current.get_pixmap(matrix=pymupdf.Matrix(1, 1), alpha=False)
                image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
                # NUMPAGES becomes fixed text; only its footer line spacing may differ.
                for block in current.get_text("blocks"):
                    if block[1] > current.rect.height * .85 and re.search(r"第\s*\d+\s*页.*共\s*\d+\s*页", block[4], re.S):
                        image.paste("white", (int(block[0]) - 2, int(block[1]) - 2,
                                              int(block[2]) + 3, int(block[3]) + 3))
                images.append(image)
            if boundary is not None:
                images[0].paste("white", (0, int(boundary), images[0].width, int(footer_top) - 2))
            diff = ImageChops.difference(*images)
            mean = sum(ImageStat.Stat(diff).mean) / 3
            checks.append({"source_page": start_page + index, "pixel_mean_difference": round(mean, 6)})
            if mean > .05:
                diagnostic_dir.mkdir(exist_ok=True)
                images[0].save(diagnostic_dir / f"p{start_page + index}-source.png")
                images[1].save(diagnostic_dir / f"p{start_page + index}-output.png")
                diff.save(diagnostic_dir / f"p{start_page + index}-diff.png")
                raise ValueError(f"原报告P{start_page + index}的页面图像发生变化（差异{mean:.3f}），停止交付并保留核对图。")
    return checks
