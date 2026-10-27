"""Keep numeric minus signs attached in output copies, not in stored source values."""

import re
from zipfile import ZipFile

from docx.oxml.ns import qn
from lxml import etree


NEGATIVE_NUMBER = re.compile(r"(?<![A-Za-z0-9_.+\-−])[-−](?:\d+(?:\.\d+)?|\.\d+)(?:[eE][+-]?\d+)?")
IDENTITY_LABEL = re.compile(
    r"(?:(?:仪器设备|设备|样机|样品|报告|证书|规格|公称)?(?:型号|规格|编号|序号)"
    r"(?:[/、](?:型号|规格|编号))*|(?:试验|测量|受样|委托|开始|结束)*日期|证书有效期)[：:]?")
NS = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}


def protect_negative_paragraph(paragraph):
    nodes = etree.XPath(".//w:t|.//w:br|.//w:tab", namespaces=NS)(paragraph)
    values = [node.text or "" if node.tag == qn("w:t") else "\n" for node in nodes]
    text = "".join(values)
    identity_ranges = []
    for field in re.finditer(r"[^；;\n，,]+", text):
        label = re.match(r"([^：:]+)[：:]", field.group())
        if label and IDENTITY_LABEL.fullmatch(re.sub(r"\s+", "", label.group(1))):
            identity_ranges.append(field.span())
    positions = {match.start() for match in NEGATIVE_NUMBER.finditer(text)
                 if not any(start <= match.start() < end for start, end in identity_ranges)}
    offset = 0
    for node, value in zip(nodes, values):
        if any(offset <= index < offset + len(value) for index in positions):
            node.text = "".join("\u2011" if offset + index in positions else char
                                for index, char in enumerate(value))
        offset += len(value)


def _identity_cells(root):
    excluded = set()
    for table in root.xpath(".//w:tbl", namespaces=NS):
        columns = set()
        for row_index, row in enumerate(table.findall(qn("w:tr"))):
            column = 0
            previous_label = ""
            for cell in row.findall(qn("w:tc")):
                text = "".join(cell.xpath(".//w:t/text()", namespaces=NS))
                spans = cell.xpath("./w:tcPr/w:gridSpan/@w:val", namespaces=NS)
                width = int(spans[0]) if spans else 1
                compact = re.sub(r"\s+", "", text)
                if row_index == 0 and IDENTITY_LABEL.fullmatch(compact):
                    columns.update(range(column, column + width))
                if columns.intersection(range(column, column + width)) or (
                        IDENTITY_LABEL.fullmatch(re.sub(r"\s+", "", previous_label))):
                    excluded.add(cell)
                previous_label = text
                column += width
    return excluded


def protect_negative_numbers(path):
    temporary = path.with_suffix(".numbers.tmp")
    try:
        with ZipFile(path) as source, ZipFile(temporary, "w") as target:
            target.comment = source.comment
            for entry in source.infolist():
                data = source.read(entry.filename)
                if entry.filename.startswith("word/") and entry.filename.endswith(".xml"):
                    root = etree.fromstring(data)
                    excluded = _identity_cells(root)
                    for paragraph in root.xpath(".//w:p", namespaces=NS):
                        if not any(parent in excluded for parent in paragraph.iterancestors(qn("w:tc"))):
                            protect_negative_paragraph(paragraph)
                    data = etree.tostring(root, encoding="UTF-8", xml_declaration=True, standalone=True)
                target.writestr(entry, data)
        temporary.replace(path)
    finally:
        if temporary.exists():
            temporary.unlink()
