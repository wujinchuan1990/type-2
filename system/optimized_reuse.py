"""V2: derive blank forms from verified report structure, never from its page numbers."""

from copy import deepcopy
from pathlib import Path
import shutil
import time
import uuid

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import documents as d
import observation_pagination as op
import optimized_pagination as fast
import result_pagination as rp


RESULT_PREFIX = "v2_result_"


def observation_geometry(doc):
    tables = op.prepare(doc)
    if not tables:
        return ()
    blocks = list(doc.element.body)
    section = next(section for block in blocks[blocks.index(tables[0]):]
                   for section in ([block] if block.tag == qn("w:sectPr")
                                   else d._xp(block, ".//w:sectPr")))
    return tuple((name, tuple(sorted(node.attrib.items())) if node is not None else ())
                 for name in ("pgSz", "pgMar", "cols", "docGrid", "vAlign", "textDirection")
                 for node in [section.find(qn("w:" + name))])


def observation_layout(doc, observations, positions):
    """A measured layout, not a promise based on similar-looking source documents."""
    queries = op.queries(observations)
    return {"geometry": observation_geometry(doc), "text": queries,
            "positions": {name + suffix: positions[name + suffix]
                          for name, _ in queries for suffix in ("_s", "_e")}}


def observation_layout_matches(doc, observations, positions, plan):
    if observation_geometry(doc) != plan["geometry"]:
        return False, "观察页型不同"
    if op.queries(observations) != plan["text"]:
        return False, "观察内容、段落或续表结构不同"
    expected = plan["positions"]
    if not expected:
        return False, "观察布局证据为空"
    first = next(iter(expected))
    page_shift = positions[first]["page"] - expected[first]["page"]
    y_shift = positions[first]["y"] - expected[first]["y"]
    # Record's chapter start can move by at most two points, not a line or a page.
    # Every measured paragraph endpoint must undergo exactly the same translation.
    if abs(y_shift) > 2:
        return False, "观察章节起点变化超过2磅"
    for name, old in expected.items():
        current = positions[name]
        if (current["page"] - old["page"] != page_shift
                or abs(current["x"] - old["x"]) > .2
                or abs(current["y"] - old["y"] - y_shift) > .2):
            return False, "观察段落的实际Word位置发生变化"
    return True, ""


class ReuseIncompatible(ValueError):
    """A known structural incompatibility permits independent record generation."""


def tag_result_sources(path):
    doc = d._load(path)
    for table in d._tables(doc):
        if not rp._table_name(table):
            continue
        if table.tblPr.find(qn("w:tblCaption")) is None:
            caption = OxmlElement("w:tblCaption")
            caption.set(qn("w:val"), RESULT_PREFIX + uuid.uuid4().hex)
            table.tblPr.append(caption)
    doc.save(path)


def observation_fragment(report, source):
    body = report.element.body
    start = d._unique_anchor(body, "样机的基本信息")
    stop = d._unique_anchor(body, "三、试验项目记录")
    blocks = list(body)[body.index(start):body.index(stop)]
    if not blocks or any(d._xp(block, ".//w:sectPr") for block in blocks):
        raise ReuseIncompatible("报告观察区域含额外分节，不能安全复用到配套观察页。")
    tables = op.prepare(report)
    if not tables or any(not rp._metadata_caption(table).startswith(op.FLOW_PREFIX) for table in tables):
        raise ReuseIncompatible("观察表缺少同源续表标识，不能区分原始独立表与自动分页。")
    fragment = d._source_fragment(report, blocks)
    # The companion retains the observe material's geometry/header, not the report's
    # surrounding sections. Relationships/styles for all body content come from report.
    section = deepcopy(source.sections[-1]._sectPr)
    for ref in d._xp(section, "./w:headerReference | ./w:footerReference"):
        section.remove(ref)
    fragment.element.body.replace(fragment.element.body[-1], section)
    rp._remove_marks(fragment, op.PREFIX)
    # Retain finished names/boundaries provisionally. The companion must prove the
    # same layout with its own Word coordinates before any of them are trusted.
    return fragment


def unpack_metadata(doc):
    try:
        d._unpack_metadata_groups(doc)
    except ValueError as exc:
        raise ReuseIncompatible(str(exc)) from exc
    rp._remove_marks(doc, rp.ROW_PREFIX)


def release_result_breaks(doc):
    """Join only siblings cloned from the same original table, never distinct tests."""
    tables = [table for table in d._tables(doc) if rp._table_name(table)]
    if any(not rp._metadata_caption(table).startswith(RESULT_PREFIX) for table in tables):
        raise ReuseIncompatible("结果表缺少原表来源标识，不能安全释放报告的固定分页。")
    for table in tables:
        if table.getparent() is None:
            continue
        while True:
            separator = table.getnext()
            following = separator.getnext() if separator is not None else None
            if (separator is None or separator.tag != qn("w:p")
                    or following is None or following.tag != qn("w:tbl")
                    or rp._metadata_caption(following) != rp._metadata_caption(table)):
                break
            if (d._text(separator).strip() or d._xp(
                    separator, ".//w:sectPr | .//w:bookmarkStart | .//w:drawing | .//w:pict | "
                    ".//w:object | .//w:fldChar | .//w:fldSimple | .//w:sdt")):
                raise ReuseIncompatible("结果自动续表之间有有效内容，不能安全合并。")
            if table.tblGrid.xml != following.tblGrid.xml:
                raise ReuseIncompatible("结果续表列宽发生变化，不能安全合并。")
            for row in d._rows(following)[1:]:
                table.append(row)
            following.getparent().remove(following)
            separator.getparent().remove(separator)
    # Blank values can fit an intact subproject that was oversized in the report.
    rp._remove_marks(doc, rp.OVERFLOW_PREFIX)
    rp.mark_result_tables(doc)


def generate_pair(root, task, outline_path, report_path, record_path, record_paths=(), *, progress=None):
    from negative_numbers import protect_negative_numbers
    from record_documents import compose_original_record

    report_path, record_path = Path(report_path), Path(record_path)
    baseline = report_path.with_name("v2-source-" + uuid.uuid4().hex + ".docx")
    reusable = report_path.with_name("v2-layout-" + uuid.uuid4().hex + ".docx")
    shutil.copyfile(report_path, baseline)
    timings = {}
    diagnostics = {"mode": "shared_structure", "fallback_reason": "",
                   "page_numbers_reused": False, "automatic_breaks_released": True}

    def stage(text):
        if progress:
            progress(text)

    try:
        started = time.monotonic()
        stage("V2型评报告：完整优化重排及核验")
        protect_negative_numbers(report_path)
        rp.prepare_metadata_groups(report_path)
        tag_result_sources(report_path)
        report_status = fast.finalize_reusable_report(
            report_path, reusable, progress=lambda text: stage("V2型评报告：" + text))
        report_status["strategy"] = "optimized_v2"
        observation_plan = report_status.pop("_observation_layout", None)
        timings["型评报告重排"] = time.monotonic() - started
        verified = d._load(reusable)
        diagnostics["observation_tables_available"] = len(op.prepare(verified))
        diagnostics["result_tables_available"] = sum(bool(rp._table_name(t)) for t in d._tables(verified))
        started = time.monotonic()
        stage("V2配套记录：保留观察候选布局，清空填写项并释放结果旧分页")
        try:
            record_result = compose_original_record(
                root, task, outline_path, reusable, record_path, record_paths, reuse_layout=True)
        except ReuseIncompatible as exc:
            observation_plan = None
            diagnostics.update(mode="independent_fallback", fallback_reason=str(exc),
                               automatic_breaks_released=False)
            stage("V2安全回退：" + str(exc))
            record_result = compose_original_record(
                root, task, outline_path, baseline, record_path, record_paths)
            record_result["warnings"].append("优化自动重排V2配套记录已回退独立重排：" + str(exc))
        protect_negative_numbers(record_path)
        rp.prepare_metadata_groups(record_path)
        timings["配套记录合成及页底约束"] = time.monotonic() - started
        started = time.monotonic()
        stage("V2配套记录：独立Word真实定位、调整及最终核验")
        record_status = fast.finalize_derived_record(
            record_path, observation_plan, progress=lambda text: stage("V2配套记录：" + text))
        diagnostics["observation_layout"] = record_status.pop("observation_reuse")
        record_status.update(strategy="optimized_v2", reuse=diagnostics)
        timings["配套记录重排"] = time.monotonic() - started
        return record_result, {"report": report_status, "record": record_status}, timings
    finally:
        baseline.unlink(missing_ok=True)
        reusable.unlink(missing_ok=True)
