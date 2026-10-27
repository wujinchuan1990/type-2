"""Windows Microsoft Word COM worker.

每一类命令对应 native_word.py / optimized_pagination.py 中的一条 AppleScript，
输出格式保持与 macOS osascript 版本完全一致（制表符分隔文本），上层解析与校验逻辑不变。

用法: python word_worker.py <命令> [--strict] <参数...>
"""

import os
import sys


def _main():
    try:
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    except Exception:
        pass
    if len(sys.argv) < 2:
        _fail("缺少Word命令名。")
    name = sys.argv[1]
    rest = sys.argv[2:]
    strict = False
    if rest and rest[0] == "--strict":
        strict = True
        rest = rest[1:]
    try:
        command = COMMANDS[name]
    except KeyError:
        _fail(f"未知的Word命令：{name}")
    try:
        command(*rest)
    except WorkerError as exc:
        _fail(str(exc))
    except Exception as exc:  # COM错误统一转为退出码1
        _fail(f"{type(exc).__name__}: {exc}")


def _fail(message):
    print(message, file=sys.stderr)
    raise SystemExit(1)


class WorkerError(Exception):
    """可预期的Word操作失败，消息会返回给调用方。"""


_APP = None


def get_word():
    global _APP
    if _APP is None:
        import pythoncom
        import win32com.client

        pythoncom.CoInitialize()
        try:
            _APP = win32com.client.Dispatch("Word.Application")
        except pythoncom.com_error as exc:
            raise WorkerError("无法连接Microsoft Word，请确认已安装：" + str(exc)) from exc
        try:
            _APP.DisplayAlerts = 0  # wdAlertsNone，避免域更新等弹窗阻塞自动化
            _APP.Visible = True
        except Exception:
            pass
    return _APP


def _normcase(path):
    return os.path.normcase(os.path.normpath(str(path)))


def find_document(app, doc_path):
    target = _normcase(doc_path)
    for doc in app.Documents:
        try:
            full_name = str(doc.FullName)
        except Exception:
            continue
        if _normcase(full_name) == target:
            return doc
    return None


def require_document(app, doc_path):
    doc = find_document(app, doc_path)
    if doc is None:
        raise WorkerError("Document path mismatch.")
    return doc


def bookmark_position_strict(doc, mark_name):
    if not doc.Bookmarks.Exists(mark_name):
        raise WorkerError(f"Bookmark absent: {mark_name}")
    bookmark = doc.Bookmarks(mark_name)
    if int(bookmark.StoryType) != 1:  # wdMainTextStory
        raise WorkerError(f"Bookmark outside main story: {mark_name}")
    start, end = int(bookmark.Start), int(bookmark.End)
    if start < 0 or start != end:
        raise WorkerError(f"Bookmark is not a collapsed measurement: {mark_name}")
    return start


def command_open(doc_path):
    app = get_word()
    if find_document(app, doc_path) is not None:
        return
    try:
        app.Documents.Open(FileName=os.path.abspath(doc_path), ConfirmConversions=False,
                           ReadOnly=False, AddToRecentFiles=False)
    except Exception as exc:
        raise WorkerError("无法交给Microsoft Word打开工作副本：" + str(exc)) from exc


def command_ready(doc_path, _name=None):
    app = get_word()
    doc = find_document(app, doc_path)
    if doc is None:
        print("")
        return
    print(str(doc.FullName))


def command_repaginate(doc_path, _pdf_path=None, _name=None):
    doc = require_document(get_word(), doc_path)
    doc.Repaginate()
    doc.Save()


def command_equal_height(doc_path, _name=None, *values):
    if len(values) % 3:
        raise WorkerError("行高均分参数必须是(首书签,末书签,行数)三元组。")
    doc = require_document(get_word(), doc_path)
    for index in range(0, len(values), 3):
        first_name, last_name = values[index], values[index + 1]
        expected_rows = int(values[index + 2])
        a = int(doc.Bookmarks(first_name).Range.Start)
        b = int(doc.Bookmarks(last_name).Range.Start)
        first_cell = doc.Range(a, a).Cells(1)
        last_cell = doc.Range(b, b).Cells(1)
        if int(last_cell.RowIndex) - int(first_cell.RowIndex) + 1 != expected_rows:
            raise WorkerError("Shared requirement row mapping changed.")
        end = int(last_cell.Range.End) - 1  # 去掉单元格结束符，对应AppleScript的content末端
        group_range = doc.Range(a, end)
        selected_rows = group_range.Rows
        selected_rows.DistributeHeight()
        common_height = float(first_cell.Height)
        if common_height < 0 or common_height > 10000:
            raise WorkerError("Word returned invalid shared row height.")
        if common_height < 22.7:
            common_height = 22.7
        selected_rows.HeightRule = 1  # wdRowHeightAtLeast
        selected_rows.Height = common_height
        if abs(float(selected_rows.Height) - common_height) > 0.01:
            raise WorkerError("Shared requirement rows did not receive equal minimum heights.")
    doc.Repaginate()
    doc.Save()
    print(len(values) // 3)


def command_editing_pages(doc_path, _name=None):
    doc = require_document(get_word(), doc_path)
    print(int(doc.ComputeStatistics(2)))  # wdStatisticPages


def _bookmark_range(doc, name, strict):
    if strict:
        try:
            start = bookmark_position_strict(doc, name)
        except WorkerError as exc:
            raise WorkerError(f"Bookmark {name}: {exc}") from exc
        return doc.Range(start, start)
    return doc.Bookmarks(name).Range


def command_page(doc_path, _name=None, *names, strict=False):
    doc = require_document(get_word(), doc_path)
    for name in names:
        text_range = _bookmark_range(doc, name, strict)
        print(f"{name}\t{int(text_range.Information(3))}")  # wdActiveEndPageNumber


def command_position(doc_path, _name=None, *names, strict=False):
    doc = require_document(get_word(), doc_path)
    for name in names:
        text_range = _bookmark_range(doc, name, strict)
        print(f"{name}\t{int(text_range.Information(3))}\t"
              f"{float(text_range.Information(5))}\t{float(text_range.Information(6))}")


def command_observation(doc_path, _name=None, *values, strict=False):
    if len(values) % 3:
        raise WorkerError("观察定位参数必须是(标记,期望文本,期望字符数)三元组。")
    doc = require_document(get_word(), doc_path)
    for index in range(0, len(values), 3):
        mark, expected_text = values[index], values[index + 1]
        expected_size = int(values[index + 2])
        try:
            if strict:
                s = bookmark_position_strict(doc, mark + "_s")
                e = bookmark_position_strict(doc, mark + "_e")
            else:
                s = int(doc.Bookmarks(mark + "_s").Range.Start)
                e = int(doc.Bookmarks(mark + "_e").Range.Start)
        except WorkerError as exc:
            raise WorkerError(f"Observation bookmarks {mark}_s / {mark}_e: {exc}") from exc
        if e - s != expected_size:
            raise WorkerError(f"Observation character mapping mismatch: {mark}")
        if expected_size > 0 and str(doc.Range(s, e).Text) != expected_text:
            raise WorkerError(f"Observation character mapping mismatch: {mark}")
        current_page = int(doc.Range(s, s).Information(3))
        segments = [f"{current_page}:0"]
        if e > s:
            last_page = int(doc.Range(e - 1, e - 1).Information(3))
            cursor = s
            while current_page < last_page:
                low, high = cursor + 1, e - 1
                while low < high:
                    middle = (low + high) // 2
                    page = int(doc.Range(middle, middle).Information(3))
                    if page > current_page:
                        high = middle
                    else:
                        low = middle + 1
                page = int(doc.Range(low, low).Information(3))
                if page <= current_page:
                    raise WorkerError("Observation page boundary did not advance.")
                segments.append(f"{page}:{low - s}")
                current_page = page
                cursor = low
        print(f"{mark}\t{','.join(segments)}")


def command_bookmarks(doc_path, _name=None, *names):
    doc = require_document(get_word(), doc_path)
    for name in names:
        text_range = doc.Bookmarks(name).Range
        physical = int(text_range.Information(3))
        printed = int(text_range.Information(1))  # wdActiveEndAdjustedPageNumber
        number = int(text_range.ListFormat.ListValue) if name.startswith("sepn_") else 0
        vertical = (float(text_range.Information(6))
                    if name.startswith("sep_after") or name.startswith("photo_") else 0.0)
        print(f"{name}\t{physical}\t{printed}\t{number}\t{vertical}")


def command_export(doc_path, pdf_path, _name=None):
    doc = require_document(get_word(), doc_path)
    try:
        doc.SaveAs2(FileName=os.path.abspath(pdf_path), FileFormat=17)  # wdFormatPDF
    except Exception as exc:
        raise WorkerError("Word PDF导出失败：" + str(exc)) from exc


def command_close(doc_path, _name=None):
    doc = find_document(get_word(), doc_path)
    if doc is not None:
        doc.Close(SaveChanges=0)  # wdDoNotSaveChanges


COMMANDS = {
    "open": command_open,
    "ready": command_ready,
    "repaginate": command_repaginate,
    "equal_height": command_equal_height,
    "editing_pages": command_editing_pages,
    "page": lambda *args: command_page(*args, strict="--strict" in sys.argv),
    "position": lambda *args: command_position(*args, strict="--strict" in sys.argv),
    "observation": lambda *args: command_observation(*args, strict="--strict" in sys.argv),
    "bookmarks": command_bookmarks,
    "export": command_export,
    "close": command_close,
}


if __name__ == "__main__":
    _main()
