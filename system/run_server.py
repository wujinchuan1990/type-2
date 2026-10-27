"""Run the unchanged workbench behind an authenticated HTTPS reverse proxy."""

import argparse
import json
import os
from pathlib import Path
import platform
import shutil
import stat

import uvicorn

from server_access import ServerAccess, read_proxy_key

PROJECT = Path(__file__).resolve().parent.parent


def configuration():
    names = ("REPORT_SERVER_ORIGIN", "REPORT_SERVER_SECRET_FILE",
             "REPORT_SERVER_DATA_DIR", "REPORT_SERVER_TEMPLATE_ROOT")
    if any(not os.environ.get(name) for name in names):
        raise ValueError("服务器必须显式配置：" + "、".join(names))
    data = Path(os.environ["REPORT_SERVER_DATA_DIR"])
    templates = Path(os.environ["REPORT_SERVER_TEMPLATE_ROOT"])
    secret = Path(os.environ["REPORT_SERVER_SECRET_FILE"])
    for path in (data, templates, secret):
        if not path.is_absolute() or path.is_symlink():
            raise ValueError("服务器数据、模板和密钥必须使用明确的绝对路径，不能使用符号链接。")
    data = data.resolve(strict=True)
    templates = templates.resolve(strict=True)
    secret = secret.resolve(strict=True)
    if (data == PROJECT or PROJECT in data.parents or secret == PROJECT or PROJECT in secret.parents
            or data == templates or data in templates.parents or templates in data.parents
            or data in secret.parents or templates in secret.parents):
        raise ValueError("服务器数据和密钥须独立于源代码、模板及彼此的目录，不能使用本机业务库。")
    details = data.stat()
    if (not data.is_dir() or details.st_uid != os.getuid()
            or stat.S_IMODE(details.st_mode) != 0o700 or not os.access(data, os.R_OK | os.W_OK | os.X_OK)):
        raise ValueError("服务器数据目录须由运行用户拥有、可读写且权限为0700。")
    if not templates.is_dir():
        raise ValueError("服务器模板根目录必须已存在；不会自动复制本机素材或字体。")
    access = ServerAccess(os.environ["REPORT_SERVER_ORIGIN"], read_proxy_key(secret))
    return access, data, templates


def create_server_app():
    from app import create_app

    access, data, templates = configuration()
    return create_app(data_dir=data, template_root=templates, server_access=access)


def preflight():
    access, _, _ = configuration()
    from word_fonts import font

    fonts = {}
    for name in ("Arial", "宋体"):
        try:
            font(name)
        except (OSError, ValueError) as error:
            fonts[name] = {"available": False, "error": str(error)}
        else:
            fonts[name] = {"available": True}
    return {
        "scope": "original_workbench_server_configuration_only",
        "configuration_valid": True,
        "public_origin": access.public_origin,
        "bind": "127.0.0.1",
        "external_authenticated_https_gateway_required": True,
        "fonts": fonts,
        "font_deployment_license_verified": False,
        "native_word": {
            "macos": platform.system() == "Darwin",
            "osascript_available": shutil.which("osascript") is not None,
            "word_application_present": Path("/Applications/Microsoft Word.app").is_dir(),
            "execution_verified": False,
            "unattended_operation_verified": False,
        },
        "business_data_migrated": False,
        "public_deployment_verified": False,
        "production_ready": False,
    }


def main():
    parser = argparse.ArgumentParser(description="原型式评价工作台的受保护服务器入口")
    parser.add_argument("--port", type=int, default=8876)
    parser.add_argument("--check", action="store_true", help="只核验显式配置，不启动服务或打开业务库")
    args = parser.parse_args()
    if not 1024 <= args.port <= 65535:
        parser.error("服务器端口必须在1024至65535之间。")
    try:
        if args.check:
            print(json.dumps(preflight(), ensure_ascii=False))
            return
        configuration()
    except (OSError, ValueError) as error:
        parser.error(str(error))
    uvicorn.run("run_server:create_server_app", factory=True, host="127.0.0.1", port=args.port,
                workers=1, proxy_headers=False, server_header=False)


if __name__ == "__main__":
    main()
