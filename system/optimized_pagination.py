"""Experimental selective pagination; standard generation remains independent."""

from contextlib import contextmanager
from pathlib import Path
import shutil
import subprocess
from tempfile import TemporaryDirectory
import threading
import time
import uuid

import documents as d
import native_word as native
import observation_pagination as op
import result_pagination as rp


PAGE_SCRIPT = r'''
on run argv
    set expectedName to POSIX file (item 1 of argv) as text
    set resultText to ""
    tell application "Microsoft Word"
        set d to document (item 2 of argv)
        if full name of d is not expectedName then error "Document path mismatch."
        repeat with bookmarkName in items 3 thru -1 of argv
            set r to text object of bookmark (contents of bookmarkName) of d
            set physicalPage to get range information r information type active end page number
            set resultText to resultText & (contents of bookmarkName) & tab & (physicalPage as text) & linefeed
        end repeat
    end tell
    return resultText
end run
'''

POSITION_SCRIPT = PAGE_SCRIPT.replace(
    'set resultText to resultText & (contents of bookmarkName) & tab & (physicalPage as text) & linefeed',
    '''set horizontalPosition to get range information r information type horizontal position relative to page
            set verticalPosition to get range information r information type vertical position relative to page
            set resultText to resultText & (contents of bookmarkName) & tab & (physicalPage as text) & tab & (horizontalPosition as text) & tab & (verticalPosition as text) & linefeed''')

OBSERVATION_SCRIPT = r'''
on run argv
    set expectedName to POSIX file (item 1 of argv) as text
    set resultText to ""
    tell application "Microsoft Word"
        set d to document (item 2 of argv)
        if full name of d is not expectedName then error "Document path mismatch."
        repeat with i from 3 to count argv by 3
            set markName to item i of argv
            set expectedText to item (i + 1) of argv
            set expectedSize to item (i + 2) of argv as integer
            set s to start of content of text object of bookmark (markName & "_s") of d
            set e to start of content of text object of bookmark (markName & "_e") of d
            set r to create range d start s end e
            if e - s is not expectedSize then error "Observation character mapping mismatch: " & markName
            if expectedSize > 0 then
                if content of r is not expectedText then error "Observation character mapping mismatch: " & markName
            end if
            set r to create range d start s end s
            set currentPage to get range information r information type active end page number
            set segments to (currentPage as text) & ":0"
            if e > s then
                set r to create range d start (e - 1) end (e - 1)
                set lastPage to get range information r information type active end page number
                set cursorPosition to s
                repeat while currentPage < lastPage
                    set lo to cursorPosition + 1
                    set hi to e - 1
                    repeat while lo < hi
                        set mid to (lo + hi) div 2
                        set r to create range d start mid end mid
                        set pg to get range information r information type active end page number
                        if pg > currentPage then
                            set hi to mid
                        else
                            set lo to mid + 1
                        end if
                    end repeat
                    set r to create range d start lo end lo
                    set pg to get range information r information type active end page number
                    if pg <= currentPage then error "Observation page boundary did not advance."
                    set segments to segments & "," & (pg as text) & ":" & ((lo - s) as text)
                    set currentPage to pg
                    set cursorPosition to lo
                end repeat
            end if
            set resultText to resultText & markName & tab & segments & linefeed
        end repeat
    end tell
    return resultText
end run
'''


def _worker_script(script):
    """识别PAGE/POSITION/OBSERVATION脚本（含FragmentNativePass改写版），映射为worker命令。"""
    if script == native.READY_SCRIPT:
        return "ready", False
    if script == native.CLOSE_SCRIPT:
        return "close", False
    strict = "bookmarkPosition" in script
    if "horizontalPosition" in script:
        return "position", strict
    if "expectedSize" in script:
        return "observation", strict
    if "physicalPage" in script:
        return "page", strict
    raise ValueError("未知的Word分页脚本，无法在Windows上执行。")


class NativePass:
    """Keep one exact-path copy open across boundary, detailed and final queries."""

    def __init__(self, source, deadline, timings):
        self.source = Path(source)
        self.deadline = deadline
        self.timings = timings
        self.path = native._WORKSPACE / ("word-" + uuid.uuid4().hex + ".docx")
        self.pdf = self.path.with_suffix(".pdf")
        self.bookmarks = {}
        self.queries = 0
        self.capture_observation_layout = False
        self.equalize_result_rows = False

    def command(self, phase, arguments, limit):
        remaining = self.deadline - time.monotonic()
        if remaining <= 0:
            raise ValueError(f"优化自动重排在“{phase}”达到5分钟总时限，未发布新版本。")
        started = time.monotonic()
        try:
            result = subprocess.run(arguments, capture_output=True, text=True, encoding="utf-8",
                                    errors="replace", timeout=min(limit, remaining))
        except subprocess.TimeoutExpired as exc:
            raise ValueError(f"优化自动重排在“{phase}”响应超时（本次等待上限{min(limit, remaining):.1f}秒），"
                             "工作副本保留供检查，未发布新版本。") from exc
        finally:
            self.timings[phase] = self.timings.get(phase, 0) + time.monotonic() - started
        if result.returncode:
            raise ValueError(f"优化自动重排在“{phase}”失败：" + result.stderr.strip())
        return result.stdout

    def script(self, phase, script, *arguments, limit=60):
        if native.IS_WINDOWS:
            name, strict = _worker_script(script)
            prefix = ("--strict",) if strict else ()
            command_arguments = native.worker_command(
                name, *prefix, str(self.path.resolve()), self.path.name, *arguments)
        else:
            command_arguments = ["/usr/bin/osascript", "-e", script, str(self.path.resolve()),
                                 self.path.name, *arguments]
        return self.command(phase, command_arguments, limit)

    def __enter__(self):
        waiting = time.monotonic()
        native._LOCK.acquire()
        self.deadline += time.monotonic() - waiting
        entered = False
        try:
            native._WORKSPACE.mkdir(parents=True, exist_ok=True)
            shutil.copyfile(self.source, self.path)
            if self.equalize_result_rows:
                started = time.monotonic()
                rp.equalize_shared_row_heights(
                    self.path, timeout_seconds=max(.1, self.deadline - time.monotonic()))
                self.timings["均分共享技术要求行高"] = (
                    self.timings.get("均分共享技术要求行高", 0) + time.monotonic() - started)
            self.command("打开工作副本", native.open_command(self.path.resolve()), 10)
            ready_deadline = min(self.deadline, time.monotonic() + 20)
            while time.monotonic() < ready_deadline:
                ready = self.script("等待文档", native.READY_SCRIPT,
                                    limit=max(.1, ready_deadline - time.monotonic()))
                if ready.strip() == str(self.path.resolve()):
                    break
                time.sleep(.5)
            else:
                raise ValueError("优化自动重排未等到本次Word工作副本，未发布新版本。")
            self.command("分页及保存", native._word_command(native.SCRIPT, str(self.path.resolve()),
                                     str(self.pdf.resolve()), self.path.name), 30)
            entered = True
            return self
        finally:
            if not entered:
                native._LOCK.release()

    def pages(self, names, phase):
        missing = list(dict.fromkeys(name for name in names if name not in self.bookmarks))
        for offset in range(0, len(missing), 100):
            batch = missing[offset:offset + 100]
            coordinates = self.capture_observation_layout and all(name.startswith(op.PREFIX) for name in batch)
            text = self.script(phase, POSITION_SCRIPT if coordinates else PAGE_SCRIPT, *batch)
            received = {}
            for line in text.strip().splitlines():
                parts = line.split("\t")
                if (len(parts) != (4 if coordinates else 2) or parts[0] not in batch or parts[0] in received
                        or not parts[1].isdigit() or int(parts[1]) < 1):
                    raise ValueError("优化自动重排收到无效页码定位，未使用不完整结果。")
                received[parts[0]] = {"page": int(parts[1])}
                if coordinates:
                    import math

                    try:
                        x, y = map(float, parts[2:])
                    except ValueError as exc:
                        raise ValueError("Word观察位置不是有效数值。") from exc
                    if not all(math.isfinite(value) and value >= 0 for value in (x, y)):
                        raise ValueError("Word观察位置不可用，不能复用布局。")
                    received[parts[0]].update(x=x, y=y)
            if set(received) != set(batch):
                raise ValueError("优化自动重排的页码定位不完整，未发布新版本。")
            self.queries += len(batch)
            self.bookmarks.update(received)
        return {name: self.bookmarks[name] for name in names}

    def export_pdf(self):
        self.command("最终PDF校验", native._word_command(native.EXPORT_SCRIPT,
                                  str(self.path.resolve()), str(self.pdf.resolve()), self.path.name), 45)
        if not self.pdf.is_file() or not self.pdf.stat().st_size:
            raise ValueError("优化自动重排未生成最终核对PDF，未发布新版本。")
        import pymupdf

        with pymupdf.open(self.pdf) as pdf:
            if not pdf.page_count:
                raise ValueError("优化自动重排的最终PDF没有页面。")
            return pdf.page_count

    def observation_locations(self, marks):
        pages = self.pages([name + suffix for name, _text in marks for suffix in ("_s", "_e")],
                           "观察段落边界筛查")
        locations = {name: [(pages[name + "_s"]["page"], 0)] for name, _text in marks
                     if pages[name + "_s"]["page"] == pages[name + "_e"]["page"]}
        pending = [(name, text) for name, text in marks if name not in locations]
        for offset in range(0, len(pending), 40):
            batch = pending[offset:offset + 40]
            arguments = [value for name, text in batch for value in (name, text, str(op._units(text)))]
            response = self.script("观察字符分页定位", OBSERVATION_SCRIPT, *arguments)
            expected = {name for name, _text in batch}
            for line in response.strip().splitlines():
                parts = line.split("\t")
                if len(parts) != 2 or parts[0] not in expected or parts[0] in locations:
                    raise ValueError("Word观察字符分页定位无效或重复，未发布新版本。")
                pairs = [part.split(":") for part in parts[1].split(",")]
                if any(len(pair) != 2 or not all(value.isdigit() for value in pair) for pair in pairs):
                    raise ValueError("Word观察字符分页定位无效，未发布新版本。")
                locations[parts[0]] = [(int(page), int(position)) for page, position in pairs]
            if not expected.issubset(locations):
                raise ValueError("Word观察字符分页定位不完整，未发布新版本。")
            self.queries += len(batch)
        return locations

    def __exit__(self, kind, value, traceback):
        try:
            if kind is None:
                self.script("关闭工作副本", native.CLOSE_SCRIPT, limit=10)
                shutil.copyfile(self.path, self.source)
                self.path.unlink()
                if self.pdf.exists():
                    self.pdf.unlink()
        finally:
            native._LOCK.release()


def _names(measured):
    return [name for _table, rows, contents in measured for pair in [*rows, *contents] for name in pair]


@contextmanager
def _word_prompt_guard():
    stop, errors = threading.Event(), []
    watcher = threading.Thread(target=rp._dismiss_prompt, args=(stop, errors), daemon=True)
    watcher.start()
    try:
        yield
    finally:
        stop.set()
        watcher.join(timeout=6)
    if errors:
        raise ValueError("；".join(errors))


def finalize_optimized_pages(path, progress=None):
    if progress:
        progress("等待本机Word队列")
    with native._LOCK:
        return _finalize_optimized_pages(path, progress)


def finalize_reusable_report(path, reuse_output, progress=None):
    """V2 alone retains provenance in a private, fully verified derivation copy."""
    if Path(path).resolve() == Path(reuse_output).resolve():
        raise ValueError("重排报告与内部复用副本不能使用同一路径。")
    with native._LOCK:
        return _finalize_optimized_pages(path, progress, reuse_output=reuse_output)


def finalize_derived_record(path, observation_plan, progress=None):
    with native._LOCK:
        return _finalize_optimized_pages(path, progress, observation_plan=observation_plan, derived=True)


def _finalize_optimized_pages(path, progress=None, *, reuse_output=None, observation_plan=None, derived=False,
                              finalized_observe=False, pdf_verifier=None, native_pass_factory=None,
                              prompt_guard_factory=None):
    started = time.monotonic()
    deadline = started + 300
    timings, warnings, rounds = {}, [], []
    names_presented = finalized_observe
    native_pass = native_pass_factory or NativePass
    prompt_guard = _word_prompt_guard if prompt_guard_factory is None else prompt_guard_factory
    if not callable(prompt_guard):
        raise ValueError("分页提示处理器必须是上下文工厂，未开始重排。")
    if finalized_observe:
        from separation_experimental_layout import observation_tables as v3_tables
        from separation_experimental_word import FragmentNativePass

        native_pass = native_pass_factory or FragmentNativePass
    select_observations = (lambda doc: v3_tables(doc, final=True, report=True)) if finalized_observe else op.prepare
    observation_reuse = {"reused": False, "fallback_reason": "", "position_tolerance_pt": .2,
                         "maximum_translation_pt": 2}
    with TemporaryDirectory(prefix="optimized-pagination-", dir=Path(path).parent) as temporary:
        working = Path(temporary) / "working.docx"
        shutil.copyfile(path, working)
        for iteration in range(20):
            if progress:
                progress(f"第{iteration + 1}轮：筛查跨页表")
            if time.monotonic() >= deadline:
                raise ValueError("优化自动重排达到5分钟总时限，未发布新版本。")
            preparation = time.monotonic()
            doc = d._load(working)
            observation_tables = select_observations(doc)
            if derived and iteration == 0 and observation_plan is None:
                op._join_flow(observation_tables)
                observation_tables = op.prepare(doc)
                observation_reuse["fallback_reason"] = "没有可核验的报告完成布局"
            observations = op.measure(doc, observation_tables)
            rp._remove_marks(doc, rp.ROW_PREFIX)
            warnings.extend(rp.repair_continuation_cells(doc))
            measured = rp._measure_marks(doc)
            metadata = rp._measure_metadata(doc)
            signatures = rp._measure_excel_signatures(doc)
            doc.save(working)
            d._normalize_report_fonts(working)
            timings["结构准备"] = timings.get("结构准备", 0) + time.monotonic() - preparation
            changed = False
            with prompt_guard():
                pass_instance = native_pass(working, deadline, timings)
                pass_instance.equalize_result_rows = (
                    reuse_output is None and observation_plan is None and not derived and not finalized_observe)
                with pass_instance as session:
                    session.capture_observation_layout = bool(observation_plan) or (
                        reuse_output is not None and names_presented)
                    boundary_names = [name for _table, rows, _contents in measured
                                      for name in (rows[0][0], rows[-1][1])]
                    boundaries = session.pages(boundary_names, "表格边界筛查")
                    pending = [item for item in measured
                               if boundaries[item[1][0][0]]["page"] != boundaries[item[1][-1][1]]["page"]]
                    doc = d._load(session.path)
                    if observations:
                        observation_tables = select_observations(doc)
                        locations = session.observation_locations(op.queries(observations))
                        if observation_plan is not None:
                            from optimized_reuse import observation_layout_matches

                            matches, reason = observation_layout_matches(
                                doc, observations, session.bookmarks, observation_plan)
                            observation_reuse.update(reused=matches, fallback_reason=reason)
                            observation_plan = None
                            if matches:
                                names_presented = True
                            else:
                                op._join_flow(observation_tables)
                                changed = True
                                if progress:
                                    progress("观察布局复用回退：" + reason + "；释放同源分页并重新修复")
                        if names_presented:
                            op.validate_name_layout(observation_tables, observations, locations)
                            if finalized_observe:
                                from separation_experimental_layout import _measured_table_pages, _validate_project_layout

                                if any(len(pages) != 1 for pages in
                                       _measured_table_pages(observation_tables, observations, locations)):
                                    raise ValueError("最终完整报告的观察表仍跨页，未发布新版本。")
                                _validate_project_layout(observation_tables)
                        elif not changed:
                            changed = op.split(doc, observation_tables, observations, locations,
                                               carry_nonfinal_remarks=reuse_output is None and not derived)
                    if pending and not changed:
                        details = session.pages(_names(pending), "跨页表详细定位")
                        changed = rp.split_measured_tables(doc, pending, details)
                    if not changed:
                        if progress:
                            progress(f"第{iteration + 1}轮：最终全量定位核验")
                        # No stale cross-pass cache: all cells and metadata are checked
                        # against this same open document before the only PDF export.
                        bookmarks = session.pages(_names(measured) +
                                                  [name for group in [*metadata, *signatures] for name in group],
                                                  "最终全量定位")
                        changed = rp.split_measured_tables(doc, measured, bookmarks)
                        if not changed:
                            rp._validate_metadata_pages(metadata, bookmarks)
                            rp._validate_excel_signature_pages(signatures, bookmarks)
                            if observations and not names_presented:
                                changed = op.finish_name_layout(observation_tables)
                                names_presented = True
                            if not changed:
                                if reuse_output is not None and observations and not session.capture_observation_layout:
                                    names = [name + suffix for name, _ in op.queries(observations)
                                             for suffix in ("_s", "_e")]
                                    session.capture_observation_layout = True
                                    for name in names:
                                        session.bookmarks.pop(name, None)
                                    session.pages(names, "报告观察布局留证")
                                pages = session.export_pdf()
                                if pdf_verifier is not None:
                                    pdf_verifier(doc, session.pdf)
                    rounds.append({"pass": iteration + 1, "tables": len(measured),
                                   "detailed_tables": len(pending), "positions_queried": session.queries,
                                   "changed": changed})
            if not changed:
                layout = None
                if reuse_output is not None:
                    doc.save(reuse_output)
                    if observations:
                        from optimized_reuse import observation_layout

                        layout = observation_layout(doc, observations, session.bookmarks)
                rp._remove_marks(doc, rp.ROW_PREFIX)
                rp._remove_marks(doc, rp.TABLE_PREFIX)
                rp._remove_marks(doc, rp.OVERFLOW_PREFIX)
                op.cleanup(doc)
                for caption in d._xp(doc.element, ".//w:tblPr/w:tblCaption"):
                    if caption.get(d.qn("w:val"), "").startswith("v2_result_"):
                        caption.getparent().remove(caption)
            doc.save(working)
            d._normalize_report_fonts(working)
            if not changed:
                shutil.copyfile(working, path)
                status = {"pages": pages, "passes": iteration + 1, "strategy": "optimized",
                        "metadata_groups_checked": len(metadata), "rounds": rounds, "pdf_exports": 1,
                        "excel_signature_groups_checked": len(signatures),
                        "elapsed_seconds": round(time.monotonic() - started, 3),
                        "stage_seconds": {key: round(value, 3) for key, value in timings.items()},
                        "warnings": list(dict.fromkeys(warnings))}
                if reuse_output is not None:
                    status["_observation_layout"] = layout
                if derived:
                    status["observation_reuse"] = observation_reuse
                return status
    raise ValueError("优化自动重排超过20轮仍未收敛，未发布新版本。")
