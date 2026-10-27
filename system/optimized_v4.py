"""Shared experimental layout with full-report semantics and publication guards."""

from pathlib import Path
import shutil
from tempfile import TemporaryDirectory
import time

import documents as d
import native_word
import observation_pagination as op
import optimized_pagination as v1
import optimized_v3 as v3
import separation_experimental_layout as layout


def result_snapshot(doc):
    rows = []
    for table in d._tables(doc):
        physical_rows = d._rows(table)
        if not physical_rows or not d._result_header_row(physical_rows[0]):
            continue
        for index, (_physical, effective) in enumerate(d._effective_rows(table)):
            if index == 0 or d._result_header_row(physical_rows[index]):
                continue
            rows.append(tuple("".join(op._text(cell).replace("\u2011", "-").split()) for cell in effective))
    return tuple(rows)


def verify_results(expected, doc):
    if result_snapshot(doc) != expected:
        raise ValueError("V4排版改变了结果行顺序、技术要求、样机编号或实测值，未发布新版本。")


def finalize_v4_pages(path, progress=None):
    with native_word._LOCK:
        return _finalize_v4_pages(
            path, progress, native_pass_factory=v3.V3NativePass, renderer=native_word.render,
            prompt_guard_factory=v1._word_prompt_guard)


def _finalize_v4_pages(path, progress=None, *, native_pass_factory, renderer, prompt_guard_factory):
    """Run the same full V4 policy with an explicitly supplied, caller-owned native runtime."""
    if not all(callable(part) for part in (native_pass_factory, renderer, prompt_guard_factory)):
        raise ValueError("V4原生会话、最终渲染及提示处理器必须完整提供，未开始重排。")
    started = time.monotonic()
    with TemporaryDirectory(prefix="optimized-v4-", dir=Path(path).parent) as temporary:
        working = Path(temporary) / "working.docx"
        shutil.copyfile(path, working)
        doc = d._load(working)
        expected = result_snapshot(doc)
        v3._format_header(doc)
        before_observation = v3.observation_fragment(doc)
        doc.save(working)
        if progress:
            progress("V4：共用试验分页器整理完整报告结果，保留样机组和实测值")
        results = layout.paginate(
            working, "result", progress, report_results=True, native_pass_factory=native_pass_factory)
        verify_results(expected, d._load(working))
        if progress:
            progress("V4：共用试验观察页整理、续表及容量回填，不执行拆分数据改写")
        observation = layout.paginate(
            working, "observe", progress, report_observe=True, native_pass_factory=native_pass_factory)
        verify_results(expected, d._load(working))

        def verify_pdf(document, pdf):
            verify_results(expected, document)
            v3._verify_pdf(before_observation, document, pdf)

        if progress:
            progress("V4：完整报告结果、页底及签名组全量检查")
        status = v1._finalize_optimized_pages(
            working, progress, finalized_observe=True, pdf_verifier=verify_pdf,
            native_pass_factory=native_pass_factory, prompt_guard_factory=prompt_guard_factory)
        if progress:
            progress("V4：重新打开最终可编辑文件，复核Word/PDF和数据完整性")
        final_pdf = Path(temporary) / "verified.pdf"
        rendered = renderer(working, final_pdf, progress=progress, verify_editing_pages=True)
        verify_pdf(d._load(working), final_pdf)
        if rendered["pages"] != status["pages"]:
            raise ValueError("V4最终保存后的页数发生变化，未发布新版本。")
        status.update(strategy="optimized_v4", observation=observation, results=results,
                      layout_policy="shared-experimental-full-report-v4",
                      result_rows_checked=len(expected), result_data_preserved=True,
                      post_save_verified=True, elapsed_seconds=round(time.monotonic() - started, 3))
        observation["final_review"].update(pdf_verified=True, no_header_only_pages=True)
        status["passes"] += results["passes"] + observation["passes"]
        status["pdf_exports"] += 2
        for prefix, part in (("V4结果", results), ("V4观察", observation)):
            for phase, seconds in part["stage_seconds"].items():
                status["stage_seconds"][prefix + "/" + phase] = round(seconds, 3)
        shutil.copyfile(working, path)
        return status
