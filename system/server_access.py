"""Explicit access policy for the original workbench behind a protected HTTPS gateway."""

import hmac
import logging
import os
import re
import stat
from dataclasses import dataclass, field
from pathlib import Path
from typing import Optional
from urllib.parse import urlsplit

from starlette.requests import Request
from starlette.responses import JSONResponse

LOGGER = logging.getLogger("report_workbench")


@dataclass(frozen=True)
class ServerAccess:
    public_origin: str
    proxy_key: bytes = field(repr=False)

    def __post_init__(self):
        origin = urlsplit(self.public_origin)
        hostname = origin.hostname or ""
        if (origin.scheme != "https" or origin.username is not None or origin.password is not None
                or origin.path or origin.query or origin.fragment
                or not re.fullmatch(r"[a-z0-9]+(?:[.-][a-z0-9-]+)*", hostname)
                or any(not label or len(label) > 63 or label.startswith("-") or label.endswith("-")
                       for label in hostname.split("."))
                or origin.port == 443 or origin.port == 0
                or self.public_origin != "https://" + hostname + (
                    ":" + str(origin.port) if origin.port is not None else "")
                or "*" in origin.netloc
                or not isinstance(self.proxy_key, bytes)
                or re.fullmatch(rb"[a-f0-9]{64}", self.proxy_key) is None):
            raise ValueError("服务器须配置准确的HTTPS源地址和独立的64位十六进制网关密钥。")

    @property
    def hostname(self) -> str:
        return urlsplit(self.public_origin).netloc.split(":")[0]

    def reject(self, request: Request) -> Optional[JSONResponse]:
        keys = request.headers.getlist("x-workbench-proxy-key")
        if (len(keys) != 1 or len(keys[0]) != 64
                or not hmac.compare_digest(keys[0].encode("utf-8"), self.proxy_key)):
            return self._rejected(401, "proxy_authority", "服务器访问未获站点网关核验。")
        hosts = request.headers.getlist("host")
        if len(hosts) != 1 or hosts[0].lower() != urlsplit(self.public_origin).netloc:
            return self._rejected(400, "host", "请求域名与已配置的服务器域名不符。")
        origins = request.headers.getlist("origin")
        if origins and (len(origins) != 1 or origins[0] != self.public_origin):
            return self._rejected(403, "origin", "仅允许已配置的HTTPS站点同源访问。")
        sites = request.headers.getlist("sec-fetch-site")
        if request.url.path.startswith("/api/") and sites and (
                len(sites) != 1 or sites[0] not in ("same-origin", "none")):
            return self._rejected(403, "fetch_site", "不接受其他站点发起的业务请求。")
        return None

    @staticmethod
    def _rejected(status, code, message):
        LOGGER.warning("Protected workbench request rejected (%s).", code)
        return JSONResponse({"detail": message}, status_code=status,
                            headers={"Cache-Control": "no-store", "X-Content-Type-Options": "nosniff"})


def read_proxy_key(path: Path) -> bytes:
    if not path.is_absolute() or path.is_symlink():
        raise ValueError("网关密钥须使用独立的绝对文件路径，不能使用符号链接。")
    descriptor = os.open(path, os.O_RDONLY | os.O_NOFOLLOW)
    with os.fdopen(descriptor, "rb") as stream:
        details = os.fstat(stream.fileno())
        if (not stat.S_ISREG(details.st_mode) or details.st_uid != os.getuid()
                or stat.S_IMODE(details.st_mode) != 0o600):
            raise ValueError("网关密钥文件须由运行用户拥有且权限为0600。")
        value = stream.read(66)
    if re.fullmatch(rb"[a-f0-9]{64}\n?", value) is None:
        raise ValueError("网关密钥须为64位十六进制随机值，可带一个结尾换行。")
    return value.rstrip(b"\n")
