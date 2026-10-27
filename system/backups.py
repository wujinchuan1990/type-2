"""Consistent local backup packages and isolated, non-destructive restore checks."""

from contextlib import closing, contextmanager
from datetime import datetime, timezone
import hashlib
import json
from pathlib import Path, PurePosixPath
import re
import shutil
import sqlite3
import stat
import threading
import uuid
from zipfile import ZIP_DEFLATED, ZipFile, BadZipFile

from fastapi import APIRouter, Depends, File, UploadFile

from download_transport import DownloadOptions, download_options, download_response

MAX_BYTES = 2 * 1024 ** 3
MAX_MANIFEST_BYTES = 8 * 1024 ** 2


class BackupGate:
    def __init__(self):
        self.condition = threading.Condition()
        self.active = 0
        self.exclusive = False

    def enter(self):
        with self.condition:
            self.condition.wait_for(lambda: not self.exclusive)
            self.active += 1

    def leave(self):
        with self.condition:
            self.active -= 1
            self.condition.notify_all()

    @contextmanager
    def snapshot(self):
        with self.condition:
            self.condition.wait_for(lambda: not self.exclusive)
            self.exclusive = True
            self.condition.wait_for(lambda: self.active == 0)
        try:
            yield
        finally:
            with self.condition:
                self.exclusive = False
                self.condition.notify_all()


def digest(path):
    value = hashlib.sha256()
    with path.open("rb") as source:
        for chunk in iter(lambda: source.read(1024 * 1024), b""):
            value.update(chunk)
    return value.hexdigest()


def make_backup(store, root, folder):
    identity = uuid.uuid4().hex
    working = folder / identity
    working.mkdir(parents=True)
    complete = False
    try:
        result = _write_backup(store, root, working, identity)
        complete = True
        return result
    finally:
        if not complete:
            shutil.rmtree(working)


def _write_backup(store, root, working, identity):
    for path in (store.directory, store.path, root / "system", root / "system/static",
                 root / "report model", root / "cover of original record"):
        if path.is_symlink():
            raise ValueError("备份源含符号链接，未跟随外部路径：" + str(path))
    for kind in ("generation_jobs", "photo_previews", "exports"):
        for item in store.list(kind):
            if item.get("status") in ("queued", "running") or item.get("layout", {}).get("status") == "running":
                raise ValueError("仍有生成或Word预览正在运行，请完成后再备份。")
    for path in (store.directory / "separations").glob("*/job.json"):
        if json.loads(path.read_text())["status"] in ("analyzing", "exporting"):
            raise ValueError("报告拆分正在运行，请完成后再备份。")
    database = working / "workbench.sqlite3"
    with closing(sqlite3.connect(str(store.path))) as source, closing(sqlite3.connect(str(database))) as target:
        source.backup(target)
        if target.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
            raise ValueError("备份数据库完整性检查失败。")
    entries = {"system/data/workbench.sqlite3": database}
    for path in store.directory.rglob("*"):
        relative = path.relative_to(store.directory)
        if relative.parts[0] in ("backups", "word-work") or path.name.startswith("workbench.sqlite3"):
            continue
        if path.is_symlink():
            raise ValueError("备份目录含符号链接，未跟随外部路径：" + str(relative))
        if path.is_file():
            entries["system/data/" + relative.as_posix()] = path
    for name in ("report model", "cover of original record"):
        for path in (root / name).rglob("*"):
            if path.is_symlink():
                raise ValueError("素材含符号链接，未跟随外部路径：" + str(path.relative_to(root)))
            if path.is_file() and not path.name.startswith(("~$", ".")):
                entries[path.relative_to(root).as_posix()] = path
    for path in (root / "system").glob("*.py"):
        entries["system/" + path.name] = path
    for path in (root / "system/static").glob("*"):
        if path.is_file():
            entries["system/static/" + path.name] = path
    for name in ("requirements.txt", "README.md", "start.command"):
        path = root / "system" / name
        if path.is_file():
            entries["system/" + name] = path
    for name, path in entries.items():
        if path.is_symlink():
            raise ValueError("备份源含符号链接，未跟随外部路径：" + name)
    if sum(path.stat().st_size for path in entries.values()) > MAX_BYTES:
        raise ValueError("备份内容超过2GB，请先人工归档历史文件；未删除任何文件。")
    manifest = {"format": 1, "created": datetime.now(timezone.utc).isoformat(),
                "files": {name: digest(path) for name, path in entries.items()}}
    encoded = json.dumps(manifest, ensure_ascii=False).encode("utf-8")
    if len(encoded) > MAX_MANIFEST_BYTES:
        raise ValueError("备份清单过大，请先人工归档历史文件；未删除任何文件。")
    _check_references(database, manifest["files"])
    temporary = working / "backup.tmp"
    with ZipFile(temporary, "w", ZIP_DEFLATED) as archive:
        for name, path in entries.items():
            archive.write(path, name)
            if digest(path) != manifest["files"][name]:
                raise ValueError("文件在备份期间发生变化，请重试：" + name)
        archive.writestr("backup-manifest.json", encoded)
    for name, path in entries.items():
        if digest(path) != manifest["files"][name]:
            raise ValueError("文件在备份期间发生变化，请重试：" + name)
    output = working / "backup.zip"
    temporary.replace(output)
    database.unlink()
    return {"id": identity, "created": manifest["created"], "files": len(entries),
            "bytes": output.stat().st_size, "url": f"/api/backups/{identity}/download"}


def _check_references(database, files):
    with closing(sqlite3.connect(database.as_uri() + "?mode=ro", uri=True)) as db:
        objects = {(kind, identity): json.loads(payload) for kind, identity, payload in db.execute(
            "SELECT kind,id,payload FROM objects WHERE kind IN ('tasks','uploads','photos')")}
    for (kind, identity), item in objects.items():
        if kind in ("uploads", "photos"):
            extension = item.get("extension", ".docx" if kind == "uploads" else ".png")
            if kind == "photos" and extension not in (".png", ".jpg"):
                raise ValueError("备份中照片存储格式无效：" + identity)
            name = f"system/data/{kind}/{identity}{extension}"
            if name not in files or (item.get("sha256") and files[name] != item["sha256"]):
                raise ValueError("备份中上传资料或照片缺失/内容与数据库不符：" + identity)
        if kind == "tasks":
            uploads = [*item.get("record_ids", []), *([item["outline_id"]] if item.get("outline_id") else [])]
            for reference_kind, identities in (("uploads", uploads), ("photos", item.get("photos", {}).values())):
                if any((reference_kind, key) not in objects for key in identities):
                    raise ValueError("备份任务引用的上传资料或照片缺失：" + item.get("name", identity))


def verify_restore(package, target):
    if target.exists() or target.is_symlink():
        raise ValueError("恢复校验目录必须是新的独立目录，未覆盖已有内容。")
    complete = False
    try:
        result = _verify_restore(package, target)
        complete = True
        return result
    except sqlite3.DatabaseError as exc:
        raise ValueError("恢复副本数据库无效或业务表损坏。") from exc
    finally:
        if not complete and target.is_dir():
            shutil.rmtree(target)


def _verify_restore(package, target):
    try:
        with ZipFile(package) as archive:
            infos = archive.infolist()
            if len(infos) > 100000 or sum(info.file_size for info in infos) > MAX_BYTES:
                raise ValueError("备份展开后超过允许大小。")
            names = [info.filename for info in infos]
            if len(names) != len(set(names)):
                raise ValueError("备份含重复路径。")
            for info in infos:
                path = PurePosixPath(info.filename)
                if (path.is_absolute() or ".." in path.parts or "\\" in info.filename
                        or ":" in info.filename or stat.S_ISLNK(info.external_attr >> 16)
                        or info.is_dir() or path.as_posix() != info.filename):
                    raise ValueError("备份含不安全路径，未恢复。")
            if archive.getinfo("backup-manifest.json").file_size > MAX_MANIFEST_BYTES:
                raise ValueError("备份清单超过允许大小。")
            manifest = json.loads(archive.read("backup-manifest.json"))
            if not isinstance(manifest, dict) or manifest.get("format") != 1 or not isinstance(manifest.get("files"), dict):
                raise ValueError("不支持的备份格式。")
            if set(names) != set(manifest["files"]) | {"backup-manifest.json"}:
                raise ValueError("备份文件与清单不一致。")
            if "system/data/workbench.sqlite3" not in manifest["files"]:
                raise ValueError("备份缺少数据库。")
            for name, expected in manifest["files"].items():
                if not re.fullmatch(r"[0-9a-f]{64}", str(expected)):
                    raise ValueError("备份校验值无效。")
                destination = target / name
                destination.parent.mkdir(parents=True, exist_ok=True)
                with archive.open(name) as source, destination.open("xb") as output:
                    value = hashlib.sha256()
                    for chunk in iter(lambda: source.read(1024 * 1024), b""):
                        value.update(chunk)
                        output.write(chunk)
                if value.hexdigest() != expected:
                    raise ValueError("备份文件校验失败：" + name)
    except (BadZipFile, KeyError, UnicodeDecodeError, json.JSONDecodeError) as exc:
        raise ValueError("备份包损坏或缺少有效清单。") from exc
    database = target / "system/data/workbench.sqlite3"
    with closing(sqlite3.connect(database.as_uri() + "?mode=ro", uri=True)) as db:
        if db.execute("PRAGMA integrity_check").fetchone()[0] != "ok":
            raise ValueError("恢复副本数据库完整性检查失败。")
        count = db.execute("SELECT COUNT(*) FROM objects WHERE kind='tasks'").fetchone()[0]
    _check_references(database, manifest["files"])
    return {"status": "verified", "tasks": count, "files": len(manifest["files"]),
            "message": "已恢复到独立校验目录，文件校验、数据库及任务上传/照片引用检查通过，当前系统未覆盖。"}


def routes(store, root, gate):
    router = APIRouter(prefix="/api/backups")
    folder = store.directory / "backups"
    folder.mkdir(exist_ok=True)

    def location(identity):
        if not re.fullmatch(r"[a-f0-9]{32}", identity):
            raise ValueError("无效备份标识。")
        return folder / identity

    @router.post("")
    def create():
        with gate.snapshot():
            result = make_backup(store, root, folder)
            (location(result["id"]) / "info.json").write_text(json.dumps(result, ensure_ascii=False))
        return result

    @router.get("")
    def listing():
        return sorted([json.loads(p.read_text()) for p in folder.glob("*/info.json")],
                      key=lambda item: item["created"], reverse=True)

    @router.get("/{identity}/download")
    def download(identity: str, options: DownloadOptions = Depends(download_options)):
        path = location(identity) / "backup.zip"
        if not path.is_file():
            raise KeyError(identity)
        return download_response(path, options, filename="型式评价系统备份-" + identity[:8] + ".zip")

    @router.post("/verify")
    def verify(file: UploadFile = File(...)):
        identity = "verify-" + uuid.uuid4().hex
        working = folder / identity
        working.mkdir()
        package = working / "uploaded.zip"
        size = 0
        complete = False
        try:
            with package.open("xb") as output:
                for chunk in iter(lambda: file.file.read(1024 * 1024), b""):
                    size += len(chunk)
                    if size > MAX_BYTES:
                        raise ValueError("备份包超过2GB。")
                    output.write(chunk)
            result = verify_restore(package, working / "restored")
            complete = True
            return {**result, "verification_id": identity}
        finally:
            package.unlink(missing_ok=True)
            if not complete:
                shutil.rmtree(working)

    return router
