"""V3 uses experimental observation finishing and V1's complete report checks."""

from copy import deepcopy
from pathlib import Path
from tempfile import TemporaryDirectory
import shutil
import time

from docx.oxml.ns import qn

import documents as d
import native_word
import observation_pagination as op
import optimized_pagination as v1
import separation_experimental_layout as layout
from separation_experimental_format import _header, verify_observe_lines
from separation_experimental_word import FragmentNativePass


class V3NativePass(FragmentNativePass):
    def command(self, phase, arguments, limit):
        if phase == "打开工作副本":
            if native.IS_WINDOWS:
                arguments = native.open_command(self.path.resolve())
            else:
                script = '''
on run argv
    tell application "Microsoft Word"
        if not (exists document (item 2 of argv)) then
            open file name (item 1 of argv) confirm conversions false read only false add to recent files false
        end if
    end tell
end run
'''
                arguments = ["/usr/bin/osascript", "-e", script, str(self.path.resolve()), self.path.name]
            limit = 30
        return super().command(phase, arguments, limit)


def _section(doc):
    first = op.report_tables(doc)[0]
    blocks = list(doc.element.body)
    geometry = next(node for block in blocks[blocks.index(first):]
                    for node in ([block] if block.tag == qn("w:sectPr")
                                 else d._xp(block, "./w:pPr/w:sectPr")))
    return next(section for section in doc.sections if section._sectPr == geometry)


def observation_fragment(doc):
    body = doc.element.body
    start = d._unique_anchor(body, "样机的基本信息")
    stop = d._unique_anchor(body, "三、试验项目记录")
    fragment = d._source_fragment(doc, list(body)[body.index(start):body.index(stop)])
    geometry = deepcopy(_section(doc)._sectPr)
    for node in d._xp(fragment.element.body, "./w:p/w:pPr/w:sectPr"):
        node.getparent().remove(node)
    old = fragment.element.body.find(qn("w:sectPr"))
    fragment.element.body.replace(old, geometry)
    for node in list(geometry):
        if node.tag == qn("w:headerReference") and node.get(qn("w:type")) != "default":
            geometry.remove(node)
    return fragment


def _format_header(doc):
    section = _section(doc)
    width = (section.page_width - section.left_margin - section.right_margin) / 12700
    center = (section.page_width / 2 - section.left_margin) / 12700
    header = section.header
    if _header(doc, header._element, width, center) is None:
        raise ValueError("完整报告观察章节缺少附件1页眉，不能跳过布局核验。")
    if section.different_first_page_header_footer:
        if _header(doc, section.first_page_header._element, width, center) is None:
            raise ValueError("完整报告观察章节的首页页眉缺失。")
    if doc.settings.odd_and_even_pages_header_footer:
        if _header(doc, section.even_page_header._element, width, center) is None:
            raise ValueError("完整报告观察章节的偶数页页眉缺失。")


def _verify_pdf(before, doc, pdf_path):
    import pymupdf

    fragment = observation_fragment(doc)
    layout._verify_observation_content(before, fragment, finalized_marks=True)
    layout._validate_project_layout(layout.observation_tables(fragment, final=True))
    with pymupdf.open(pdf_path) as pdf, pymupdf.open() as observed:
        for page in pdf:
            if any(rectangle.y1 < page.rect.height / 4 for rectangle in page.search_for("型式评价记录")):
                observed.insert_pdf(pdf, from_page=page.number, to_page=page.number)
        if not observed.page_count:
            raise ValueError("完整报告最终PDF缺少观察页，停止交付。")
        verify_observe_lines(fragment, observed)
        layout._verify_observation_page_content(fragment, observed)
        expected = layout._characters(fragment)
        actual = layout.Counter(layout.base.normalized_text("".join(page.get_text() for page in observed)))
        if expected - actual:
            raise ValueError("完整报告最终PDF遗漏观察正文文字，停止交付。")


def finalize_v3_pages(path, progress=None):
    started = time.monotonic()
    with native_word._LOCK, TemporaryDirectory(prefix="optimized-v3-", dir=Path(path).parent) as temporary:
        working = Path(temporary) / "working.docx"
        shutil.copyfile(path, working)
        doc = d._load(working)
        _format_header(doc)
        before = observation_fragment(doc)
        doc.save(working)
        if progress:
            progress("V3：在完整文件中按试验拆分逻辑重排观察页")
        observation = layout.paginate(
            working, "observe", progress, report_observe=True, native_pass_factory=V3NativePass)
        if progress:
            progress("V3：沿V1核验结果、页底及Excel签名，重新核验观察页和最终PDF")
        status = v1._finalize_optimized_pages(
            working, progress, finalized_observe=True,
            pdf_verifier=lambda doc, pdf: _verify_pdf(before, doc, pdf), native_pass_factory=V3NativePass)
        status["strategy"] = "optimized_v3"
        status["observation"] = observation
        observation["final_review"].update(pdf_verified=True, no_header_only_pages=True)
        status["passes"] += observation["passes"]
        for phase, seconds in observation["stage_seconds"].items():
            name = "V3观察/" + phase
            status["stage_seconds"][name] = round(seconds, 3)
        status["elapsed_seconds"] = round(time.monotonic() - started, 3)
        shutil.copyfile(working, path)
        return status
