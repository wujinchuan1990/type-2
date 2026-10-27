"""Optional bounded JSON transport for file downloads through JSON-only proxies."""

import base64
from dataclasses import dataclass
import hashlib
import mimetypes
import os
from pathlib import Path
from typing import Literal, Optional

from fastapi import HTTPException, Query
from fastapi.responses import FileResponse, JSONResponse

DOWNLOAD_CHUNK_BYTES = 1024 * 1024


@dataclass(frozen=True)
class DownloadOptions:
    encoding: Literal["binary", "base64"] = "binary"
    offset: int = 0
    version: Optional[str] = None


def download_options(
    download_encoding: Literal["binary", "base64"] = Query("binary"),
    download_offset: int = Query(0, ge=0),
    download_version: Optional[str] = Query(None, pattern=r"^[0-9a-f]{64}$"),
) -> DownloadOptions:
    return DownloadOptions(download_encoding, download_offset, download_version)


def file_version(stat):
    return hashlib.sha256(
        f"{stat.st_dev}:{stat.st_ino}:{stat.st_size}:{stat.st_mtime_ns}".encode("ascii")
    ).hexdigest()


def download_response(path, options: DownloadOptions, *, filename=None, media_type=None):
    path = Path(path)
    if options.encoding == "binary":
        if options.offset or options.version:
            raise HTTPException(422, "文件分块参数只能用于JSON下载。")
        return FileResponse(path, filename=filename, media_type=media_type)
    if options.offset and not options.version:
        raise HTTPException(422, "继续下载必须提供首块返回的文件版本。")
    try:
        source = path.open("rb")
    except FileNotFoundError as exc:
        raise HTTPException(404, "下载文件不存在。") from exc
    with source:
        before = os.fstat(source.fileno())
        version = file_version(before)
        if options.version is not None and options.version != version:
            raise HTTPException(409, "下载文件已变化，请重新下载。")
        if options.offset > before.st_size:
            raise HTTPException(416, "下载位置超出文件长度。")
        source.seek(options.offset)
        data = source.read(DOWNLOAD_CHUNK_BYTES)
        if file_version(os.fstat(source.fileno())) != version or (
                len(data) != min(DOWNLOAD_CHUNK_BYTES, before.st_size - options.offset)):
            raise HTTPException(409, "下载期间文件已变化，请重新下载。")
    name = filename or path.name
    return JSONResponse({
        "filename": name,
        "media_type": media_type or mimetypes.guess_type(name)[0] or "application/octet-stream",
        "version": version,
        "total_bytes": before.st_size,
        "offset": options.offset,
        "next_offset": options.offset + len(data),
        "eof": options.offset + len(data) == before.st_size,
        "sha256": hashlib.sha256(data).hexdigest(),
        "data_base64": base64.b64encode(data).decode("ascii"),
    }, headers={"Cache-Control": "no-store"})
