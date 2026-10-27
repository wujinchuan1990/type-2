"""Task-local certificate remarks shared by observation and summary tables."""

import re
from copy import deepcopy
from typing import List, Tuple

from docx.document import Document
from docx.oxml.ns import qn
from docx.table import Table
from lxml.etree import _Element

import documents as d
from contracts import ObservationCertificateCandidate, ObservationCertificates


LABELS = {"explosion": "防爆证书编号", "protection": "防护证书编号"}
CERTIFICATE = re.compile(r"证\s*书\s*编\s*号\s*[：:]\s*([^\s，,。；;：:（）()]+)")
CERTIFICATE_LABEL = re.compile(r"证\s*书\s*编\s*号\s*[：:]")


def certificate_targets(observe: Document, *, section="观察") -> List[Tuple[str, str, _Element]]:
    targets = []
    seen = {}
    for element in d._tables(observe):
        table = Table(element, observe)
        if not table.rows:
            continue
        headers = [d._norm(cell.text) for cell in table.rows[0].cells]
        names = [i for i, text in enumerate(headers) if "项目名称" in text or text == "评价项目"]
        if not names:
            continue
        remarks = [i for i, text in enumerate(headers) if text == "备注"]
        for row in table.rows[1:]:
            cells = row.cells
            name = " / ".join(dict.fromkeys(cells[i].text for i in names if i < len(cells)))
            compact = re.sub(r"\s+", "", name)
            kinds = []
            if "防爆" in compact:
                kinds.append("explosion")
            if "防护" in compact and "封印" not in compact and (
                    any(word in compact for word in ("防护性能", "防护等级", "外壳防护"))
                    or any(CERTIFICATE_LABEL.search(cells[i].text) for i in remarks if i < len(cells))):
                kinds.append("protection")
            if not kinds:
                continue
            if len(kinds) != 1 or len(remarks) != 1 or remarks[0] >= len(cells):
                raise ValueError(f"{section}证书备注：{name}的项目类别或备注列不唯一，无法安全填写。")
            cell = cells[remarks[0]]._tc
            if d._xp(cell, "./w:tcPr/w:gridSpan/@w:val") not in ([], ["1"]):
                raise ValueError(f"{section}证书备注：{name}的备注与其他列合并，不能覆盖已有评价标记。")
            if cell in seen and seen[cell] != kinds[0]:
                raise ValueError(f"{section}证书备注：{name}与其他证书类别共用备注，无法安全填写。")
            if cell not in seen:
                targets.append((kinds[0], name, cell))
                seen[cell] = kinds[0]
    return targets


def explosion_candidates(outline: Document, source: str):
    candidates = []
    issues = []
    blocks = list(outline.element.body)
    headings = [i for i, block in enumerate(blocks) if block.tag == qn("w:p") and
                re.fullmatch(r"(?:7)?(?:补充说明|其他说明)", d._norm(d._text(block)))]
    if len(headings) != 1:
        return [], ["防爆信息：本次大纲的补充说明/其他说明章节缺失或不唯一，请人工核对填写。"]
    pending = ""
    pending_source = ""
    for index in range(headings[0] + 1, len(blocks)):
        block = blocks[index]
        if block.tag == qn("w:sectPr"):
            continue
        if block.tag == qn("w:p"):
            paragraph = d.Paragraph(block, outline)
            if (re.match(r"^\s*(?:8[.．、]|附录|附件)", paragraph.text)
                    or d._xp(block, "./w:pPr/w:outlineLvl")
                    or paragraph.style.name.startswith("Heading")):
                break
            texts = [paragraph.text]
        else:
            texts = [d.Paragraph(p, outline).text for p in d._xp(block, ".//w:p")]
        for text in texts:
            where = f"{source} · 补充说明/其他说明 · 正文块{index + 1}"
            if pending and re.match(r"^\s*证\s*书\s*编\s*号\s*[：:]", text):
                text = pending + "\n" + text
                where = pending_source + f"～正文块{index + 1}"
            pending = text if "防爆" in text and not CERTIFICATE_LABEL.search(text) else ""
            pending_source = where
            for clause in re.split(r"[。；;]", text):
                if "防爆" not in clause:
                    continue
                matches = list(CERTIFICATE.finditer(clause))
                if not matches and CERTIFICATE_LABEL.search(clause):
                    issues.append("防爆信息中未找到明确的“证书编号：…”：" + clause.strip())
                if "防护" in clause or re.search(r"(?:无|未取得|尚未取得|不适用).{0,8}防爆|防爆.{0,8}(?:未取得|不适用)", clause):
                    issues.append("防爆信息含其他证书类别或否定说明，不能自动采用：" + clause.strip())
                for match in matches:
                    if re.match(r"\s*[,，]\s*[A-Za-z0-9._/-]*\d", clause[match.end():]):
                        issues.append("防爆证书编号后有并列编号，请人工核对：" + clause.strip())
                    for value in match.group(1).split("、"):
                        if not re.search(r"\d", value) or re.fullmatch(r"[_xX×＊*]+", value):
                            issues.append("防爆证书编号疑似占位内容，请人工核对：" + value)
                            continue
                        candidates.append(ObservationCertificateCandidate(value=value, source=where, text=text))
    unique = {candidate.value for candidate in candidates}
    if len(unique) != 1:
        issues.append("防爆证书编号未找到或存在多个不同候选，请人工填写或选择后核对。")
    return candidates, list(dict.fromkeys(issues))


def refresh_certificates(observe, outline, values: ObservationCertificates, source="当前大纲", *,
                         items=None, section="观察") -> ObservationCertificates:
    result = values.model_copy(deep=True)
    result.issues = []
    targets = certificate_targets(observe, section=section) if observe is not None else []
    if items is not None:
        targets.extend((key, "第五部分一览表 / " + name, cell)
                       for key, name, cell in certificate_targets(items, section="项目一览表"))
    for key, label in LABELS.items():
        field = getattr(result, key)
        field.targets = list(dict.fromkeys(name for kind, name, _cell in targets if kind == key))
        field.originals = list(dict.fromkeys(d._text(cell) for kind, _name, cell in targets if kind == key))
        field.candidates = []
        field.source_issues = []
        field.issues = []
        if key == "explosion":
            if outline is not None:
                field.candidates, field.source_issues = explosion_candidates(outline, source)
            else:
                field.source_issues = ["防爆信息：尚未选择本次大纲，请人工填写或上传大纲。"]
            if not field.manual_override:
                field.value = field.candidates[0].value if field.candidates and not field.source_issues else ""
        if field.targets and not field.value.strip():
            location = "观察/项目一览表" if items is not None else section
            field.issues = [f"{location}备注/{label}尚未填写，输出保留素材原文，请核对。"]
    return result


def fill_certificates(observe, outline, task, *, section="观察"):
    values = ObservationCertificates.model_validate(task.get("observation_certificates", {}))
    current = refresh_certificates(observe, outline, values, section=section)
    issues = list(current.issues)
    for key in LABELS:
        issues.extend(getattr(current, key).issues)
    for key, name, cell in certificate_targets(observe, section=section):
        value = getattr(current, key).value.strip()
        if not value:
            continue
        paragraphs = cell.findall(qn("w:p"))
        slots = [(paragraph, d.Paragraph(paragraph, observe).text) for paragraph in paragraphs]
        labels = [(paragraph, raw, match) for paragraph, raw in slots
                  for match in CERTIFICATE_LABEL.finditer(raw)]
        if len(labels) > 1:
            raise ValueError(f"{section}证书备注：{name}包含多个证书编号位置，请核对素材。")
        if labels:
            paragraph, raw, _label = labels[0]
            match = CERTIFICATE.search(raw)
            if match:
                start, end = (len(raw[:offset].replace("\n", "").replace("\t", "")) for offset in match.span(1))
                d._replace_text_span(paragraph, start, end, value)
            else:
                nodes = d._xp(paragraph, ".//w:t")
                if not nodes:
                    raise ValueError(f"{section}证书备注：{name}的编号位置不可编辑。")
                d._write_slot_text(nodes[-1], (nodes[-1].text or "") + value)
        else:
            if not d._text(cell):
                d._put(cell, "证书编号：" + value)
            else:
                if not paragraphs:
                    raise ValueError(f"{section}证书备注：{name}缺少原生文字段落。")
                paragraph = deepcopy(paragraphs[-1])
                d._put(paragraph, "证书编号：" + value)
                cell.append(paragraph)
    return issues
