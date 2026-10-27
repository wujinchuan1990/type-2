import os
import shutil
import subprocess
import sys
import threading
import time
import uuid
from pathlib import Path
from typing import Callable, List, Optional

_LOCK = threading.RLock()
_WORKSPACE = Path(__file__).resolve().parent / "data" / "word-work"

IS_WINDOWS = sys.platform == "win32"
_WORKER = Path(__file__).resolve().with_name("word_worker.py")


def worker_command(name, *arguments):
    return [sys.executable, str(_WORKER), name, *arguments]


_SCRIPT_COMMANDS = {}


def _word_command(script, *arguments):
    """把一条AppleScript调用替换为当前平台的等价命令。"""
    if not IS_WINDOWS:
        return ["/usr/bin/osascript", "-e", script, *arguments]
    import result_pagination as rp

    name = _SCRIPT_COMMANDS.get(script)
    if name is None and script == rp.EQUAL_HEIGHT_SCRIPT:
        name = "equal_height"
    if name is None:
        raise RuntimeError("未知的Word脚本，无法在Windows上执行。")
    return worker_command(name, *arguments)


def open_command(path):
    if not IS_WINDOWS:
        return ["/usr/bin/open", "-a", "Microsoft Word", str(path)]
    return worker_command("open", str(path))


def _same_path(left, right):
    if IS_WINDOWS:
        return os.path.normcase(os.path.normpath(str(left))) == os.path.normcase(os.path.normpath(str(right)))
    return str(left) == str(right)


READY_SCRIPT = r'''
on run argv
    set resultPath to ""
    tell application "Microsoft Word"
        if exists document (item 2 of argv) then
            set resultPath to full name of document (item 2 of argv)
        end if
    end tell
    if resultPath is "" then return ""
    return POSIX path of resultPath
end run
'''

SCRIPT = r'''
on run argv
    set docPath to item 1 of argv
    set pdfPath to item 2 of argv
    set expectedName to POSIX file docPath as text
    set outputName to POSIX file pdfPath as text
    tell application "Microsoft Word"
        set d to get document (item 3 of argv)
        if full name of d is not expectedName then
            error "Word opened an unexpected document; export stopped."
        end if
        repaginate d
        save d
    end tell
end run
'''

EXPORT_SCRIPT = r'''
on run argv
    set expectedName to POSIX file (item 1 of argv) as text
    set outputName to POSIX file (item 2 of argv) as text
    tell application "Microsoft Word"
        set d to get document (item 3 of argv)
        if full name of d is not expectedName then error "Document path mismatch."
        save as d file name outputName file format format PDF
    end tell
end run
'''

CLOSE_SCRIPT = r'''
on run argv
    set expectedName to POSIX file (item 1 of argv) as text
    tell application "Microsoft Word"
        if exists document (item 2 of argv) then
            set d to get document (item 2 of argv)
            if full name of d is expectedName then close d saving no
        end if
    end tell
end run
'''

EDITING_PAGES_SCRIPT = r'''
on run argv
    set expectedName to POSIX file (item 1 of argv) as text
    tell application "Microsoft Word"
        set d to get document (item 2 of argv)
        if full name of d is not expectedName then error "Document path mismatch."
        return compute statistics d statistic statistic pages
    end tell
end run
'''

BOOKMARK_SCRIPT = r'''
on run argv
    set expectedName to POSIX file (item 1 of argv) as text
    set resultText to ""
    tell application "Microsoft Word"
        set d to get document (item 2 of argv)
        if full name of d is not expectedName then error "Document path mismatch."
        repeat with bookmarkName in items 3 thru -1 of argv
            set r to text object of bookmark (contents of bookmarkName) of d
            set physicalPage to get range information r information type active end page number
            set printedPage to get range information r information type active end adjusted page number
            set listNumber to 0
            if (contents of bookmarkName) starts with "sepn_" then set listNumber to list value of list format of r
            set verticalPosition to 0
            if (contents of bookmarkName) starts with "sep_after" or (contents of bookmarkName) starts with "photo_" then set verticalPosition to get range information r information type vertical position relative to page
            set resultText to resultText & (contents of bookmarkName) & tab & (physicalPage as text) & tab & (printedPage as text) & tab & (listNumber as text) & tab & (verticalPosition as text) & linefeed
        end repeat
    end tell
    return resultText
end run
'''

_SCRIPT_COMMANDS.update({
    READY_SCRIPT: "ready",
    SCRIPT: "repaginate",
    EXPORT_SCRIPT: "export",
    CLOSE_SCRIPT: "close",
    EDITING_PAGES_SCRIPT: "editing_pages",
    BOOKMARK_SCRIPT: "bookmarks",
})


def render(
    docx: Path, pdf: Path,
    progress: Optional[Callable[[str], None]] = None,
    timeout_seconds: float = 90,
    bookmark_names: Optional[List[str]] = None,
    verify_editing_pages: bool = False,
    equalize_result_rows: bool = False,
    _shared_row_height_groups=None,
):
    """Only open this generation's uniquely named copy, never users' source files."""
    if progress is not None:
        progress("等待本机Word队列")
    _LOCK.acquire()
    deadline = time.monotonic() + timeout_seconds
    destination_docx, destination_pdf = docx, pdf

    def stage(name):
        if progress is not None:
            progress(name)

    def command(arguments, limit):
        remaining = deadline - time.monotonic()
        if remaining <= 0:
            raise ValueError("本次 Word 检查已达到总时限，原始 Word 仍可下载；不会自动重试。")
        return subprocess.run(arguments, capture_output=True, text=True, encoding="utf-8",
                              errors="replace", timeout=min(limit, remaining))

    try:
        stage("opening")
        _WORKSPACE.mkdir(parents=True, exist_ok=True)
        docx = _WORKSPACE / f"word-{uuid.uuid4().hex}.docx"
        pdf = docx.with_suffix(".pdf")
        shutil.copyfile(destination_docx, docx)
        height_groups = _shared_row_height_groups or []
        if equalize_result_rows:
            import result_pagination as rp

            stage("equalizing_shared_result_rows")
            rp.equalize_shared_row_heights(
                docx, timeout_seconds=max(.1, deadline - time.monotonic()), progress=progress)
        opened = command(open_command(docx.resolve()), 10)
        if opened.returncode:
            raise ValueError("无法交给 Microsoft Word 打开工作副本：" + opened.stderr.strip())
        stage("waiting_for_document")
        ready_deadline = min(deadline, time.monotonic() + 20)
        while time.monotonic() < ready_deadline:
            ready = command(
                _word_command(READY_SCRIPT, str(docx.resolve()), docx.name),
                max(0.1, ready_deadline - time.monotonic()),
            )
            if ready.returncode:
                raise ValueError("Word 文档状态查询失败：" + ready.stderr.strip())
            if ready.stdout.strip() and _same_path(ready.stdout.strip(), docx.resolve()):
                break
            time.sleep(0.5)
        else:
            raise ValueError("Word 未返回本次工作副本，原因尚未确定。原文件仍可下载，可改为手工核对。")
        stage("repaginating")
        result = command(
            _word_command(SCRIPT, str(docx.resolve()), str(pdf.resolve()), docx.name),
            30,
        )
        if result.returncode:
            raise ValueError("Word 分页/PDF 校验失败：" + result.stderr.strip())
        for offset in range(0, len(height_groups), 40):
            import result_pagination as rp

            batch = height_groups[offset:offset + 40]
            stage("equalizing_shared_result_rows")
            result = command(
                _word_command(rp.EQUAL_HEIGHT_SCRIPT, str(docx.resolve()), docx.name,
                              *[str(value) for group in batch for value in group]), 60)
            if result.returncode:
                raise ValueError("Word共享技术要求行高均分失败：" + result.stderr.strip())
            if result.stdout.strip() != str(len(batch)):
                raise ValueError("Word共享技术要求行高均分返回的组数不完整：" + result.stdout.strip())
        editing_pages = None
        if verify_editing_pages:
            stage("reading_editing_pages")
            result = command(
                _word_command(EDITING_PAGES_SCRIPT, str(docx.resolve()), docx.name),
                30,
            )
            if result.returncode:
                raise ValueError("Word 编辑视图页数查询失败：" + result.stderr.strip())
            if not result.stdout.strip().isdigit() or int(result.stdout.strip()) < 1:
                raise ValueError("Word 返回的编辑视图页数无效，停止导出。")
            editing_pages = int(result.stdout.strip())
        bookmarks = {}
        if bookmark_names:
            stage("reading_page_boundaries")
            for offset in range(0, len(bookmark_names), 100):
                result = command(
                    _word_command(BOOKMARK_SCRIPT, str(docx.resolve()), docx.name,
                                  *bookmark_names[offset:offset + 100]),
                    60,
                )
                if result.returncode:
                    raise ValueError("Word 实际页码定位失败：" + result.stderr.strip())
                for row in result.stdout.strip().splitlines():
                    name, physical, printed, number, vertical = row.split("\t")
                    bookmarks[name] = {"page": int(physical), "printed": int(printed), "number": int(number),
                                       "y": float(vertical.replace(",", "."))}
            if set(bookmarks) != set(bookmark_names):
                raise ValueError("Word 返回的页码定位不完整，不能按缓存分页拆分。")
        stage("exporting_pdf")
        result = command(
            _word_command(EXPORT_SCRIPT, str(docx.resolve()), str(pdf.resolve()), docx.name),
            45,
        )
        if result.returncode:
            raise ValueError("Word PDF 导出失败：" + result.stderr.strip())
        if not pdf.is_file() or pdf.stat().st_size == 0:
            raise ValueError("Word 未生成有效 PDF，不能标记为已校验。")
        stage("reading_pdf")
        import pymupdf

        with pymupdf.open(pdf) as rendered:
            if not rendered.page_count:
                raise ValueError("生成的 PDF 没有页面。")
            pages = rendered.page_count
        stage("closing_copy")
        result = command(
            _word_command(CLOSE_SCRIPT, str(docx.resolve()), docx.name),
            10,
        )
        if result.returncode:
            raise ValueError("Word 已导出，但无法关闭工作副本：" + result.stderr.strip())
        if verify_editing_pages and editing_pages != pages:
            raise ValueError(f"Word 编辑视图为{editing_pages}页，PDF为{pages}页，分页不一致，停止交付。")
        if height_groups:
            import documents as d

            document = d._load(docx)
            rp._remove_marks(document, rp.HEIGHT_PREFIX)
            document.save(docx)
        shutil.copyfile(docx, destination_docx)
        shutil.copyfile(pdf, destination_pdf)
        docx.unlink()
        pdf.unlink()
        return {"pages": pages, **({"bookmarks": bookmarks} if bookmark_names else {})}
    except subprocess.TimeoutExpired as exc:
        raise ValueError(
            "Word 检查响应超时；原文件未改写、仍可下载。不会自动重试，也未关闭其他文档。"
            f"若 Word 提示文件访问，请仅授权专用排版目录：{_WORKSPACE}。"
        ) from exc
    finally:
        _LOCK.release()
