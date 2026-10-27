import json
import re
import sqlite3
import uuid
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path

from contracts import DEFAULT_EQUIPMENT_STORAGE_LOCATION, EQUIPMENT_TRIAL_GROUP_BY_IDENTITY
from equipment_typography import equipment_identity, normalize_equipment_typography
from report_types import normalize_report_payload


DEFAULT_PRINCIPLE = {
    "id": "impeller-water-meter",
    "name": "叶轮式水表原理",
    "text": "当水流通过水表时，驱动叶轮旋转，因水流驱动叶轮处喷口的截面积为常数，"
            "水流的流速与叶轮的转速成正比，故叶轮的转速与流量成正比。"
            "通过叶轮轴上的联动部件与计数机构相连接，使计数机构累积叶轮的转数，"
            "从而记下水表的水量。",
}


class ConflictError(ValueError):
    pass


class Store:
    def __init__(self, directory: Path):
        self.directory = directory
        directory.mkdir(parents=True, exist_ok=True)
        for name in ("uploads", "exports", "photos"):
            (directory / name).mkdir(exist_ok=True)
        self.path = directory / "workbench.sqlite3"
        with self.connection() as db:
            db.execute("BEGIN IMMEDIATE")
            db.execute(
                "CREATE TABLE IF NOT EXISTS objects "
                "(kind TEXT NOT NULL,id TEXT NOT NULL,payload TEXT NOT NULL,"
                "revision INTEGER NOT NULL,updated_at TEXT NOT NULL,PRIMARY KEY(kind,id))"
            )
            db.execute(
                "CREATE TABLE IF NOT EXISTS history "
                "(kind TEXT NOT NULL,id TEXT NOT NULL,revision INTEGER NOT NULL,"
                "payload TEXT NOT NULL,updated_at TEXT NOT NULL)"
            )
            encoded = json.dumps(dict(DEFAULT_PRINCIPLE, revision=1), ensure_ascii=False)
            now = datetime.now(timezone.utc).isoformat()
            inserted = db.execute(
                "INSERT OR IGNORE INTO objects VALUES (?,?,?,?,?)",
                ("principles", DEFAULT_PRINCIPLE["id"], encoded, 1, now),
            )
            if inserted.rowcount:
                db.execute("INSERT INTO history VALUES (?,?,?,?,?)",
                           ("principles", DEFAULT_PRINCIPLE["id"], 1, encoded, now))
            for row in db.execute(
                "SELECT id,payload,revision FROM objects WHERE kind='equipment'"
            ).fetchall():
                equipment = json.loads(row["payload"])
                cleaned = dict(equipment)
                cleaned.pop("version", None)
                cleaned.pop("revision", None)
                cleaned.setdefault("storage_location", DEFAULT_EQUIPMENT_STORAGE_LOCATION)
                identity = equipment_identity(*(str(cleaned.get(key, "")) for key in ("name", "model", "number")))
                if "trial_group" not in cleaned and identity in EQUIPMENT_TRIAL_GROUP_BY_IDENTITY:
                    cleaned["trial_group"] = EQUIPMENT_TRIAL_GROUP_BY_IDENTITY[identity]
                cleaned = normalize_equipment_typography(cleaned)
                if cleaned == equipment and row["revision"] == 0:
                    continue
                encoded = json.dumps(cleaned, ensure_ascii=False)
                db.execute(
                    "UPDATE objects SET payload=?,revision=0 WHERE kind='equipment' AND id=?",
                    (encoded, row["id"]),
                )
            db.execute("DELETE FROM history WHERE kind='equipment'")

    @contextmanager
    def connection(self):
        with sqlite3.connect(str(self.path), timeout=20) as db:
            db.row_factory = sqlite3.Row
            yield db

    def get(self, kind, identity):
        with self.connection() as db:
            row = db.execute(
                "SELECT payload FROM objects WHERE kind=? AND id=?", (kind, identity)
            ).fetchone()
        if not row:
            raise KeyError(identity)
        return normalize_report_payload(json.loads(row["payload"]))

    def list(self, kind):
        with self.connection() as db:
            rows = db.execute(
                "SELECT payload,updated_at FROM objects WHERE kind=? ORDER BY updated_at DESC",
                (kind,),
            ).fetchall()
        return [dict(normalize_report_payload(json.loads(row["payload"])), updated_at=row["updated_at"]) for row in rows]

    def first_saved_at(self, kind, identity):
        with self.connection() as db:
            row = db.execute(
                "SELECT updated_at FROM history WHERE kind=? AND id=? AND revision=1",
                (kind, identity),
            ).fetchone()
        return row["updated_at"] if row else ""

    def save(self, kind, payload, revision_key="revision"):
        data = normalize_report_payload(dict(payload))
        identity = data.get("id") or uuid.uuid4().hex
        with self.connection() as db:
            db.execute("BEGIN IMMEDIATE")
            if kind == "tasks":
                if db.execute("SELECT 1 FROM objects WHERE kind='deleted_tasks' AND id=?", (identity,)).fetchone():
                    raise ConflictError("该任务已被删除，不能从旧窗口重新保存。请新建任务。")
                references = [(identity, "record") for identity in data.get("record_ids", [])]
                if data.get("outline_id"):
                    references.append((data["outline_id"], "outline"))
                for upload_id, expected_kind in references:
                    upload = db.execute(
                        "SELECT payload FROM objects WHERE kind='uploads' AND id=?", (upload_id,)
                    ).fetchone()
                    if not upload or json.loads(upload["payload"])["kind"] != expected_kind:
                        raise ConflictError("任务引用的上传资料已删除或类型不符，请刷新并重新选择：" + upload_id)
                for photo_id in set(data.get("photos", {}).values()):
                    if not db.execute("SELECT 1 FROM objects WHERE kind='photos' AND id=?", (photo_id,)).fetchone():
                        raise ConflictError("任务引用的照片不存在，请重新上传：" + photo_id)
            previous = db.execute(
                "SELECT revision FROM objects WHERE kind=? AND id=?", (kind, identity)
            ).fetchone()
            version = previous["revision"] if previous else 0
            versioned = kind != "equipment"
            if versioned and previous and data.get(revision_key) != version:
                raise ConflictError("内容已在其他窗口修改，请重新打开后再保存。")
            data.update(id=identity)
            if versioned:
                data[revision_key] = version + 1
            else:
                data.pop("version", None)
                data.pop("revision", None)
            stored_revision = version + 1 if versioned else 0
            now = datetime.now(timezone.utc).isoformat()
            encoded = json.dumps(data, ensure_ascii=False)
            db.execute(
                "INSERT OR REPLACE INTO objects VALUES (?,?,?,?,?)",
                (kind, identity, encoded, stored_revision, now),
            )
            if versioned:
                db.execute(
                    "INSERT INTO history VALUES (?,?,?,?,?)",
                    (kind, identity, stored_revision, encoded, now),
                )
        return data

    def delete_task(self, identity, revision):
        with self.connection() as db:
            db.execute("BEGIN IMMEDIATE")
            previous = db.execute(
                "SELECT revision FROM objects WHERE kind='tasks' AND id=?", (identity,)
            ).fetchone()
            if not previous:
                raise KeyError(identity)
            if previous["revision"] != revision:
                raise ConflictError("任务已在其他窗口更新，请重新打开任务列表、核对版本后再删除。")
            now = datetime.now(timezone.utc).isoformat()
            marker = json.dumps({"id": identity, "deleted_at": now})
            db.execute("INSERT INTO objects VALUES (?,?,?,?,?)", ("deleted_tasks", identity, marker, 1, now))
            db.execute("DELETE FROM objects WHERE kind='tasks' AND id=?", (identity,))
            db.execute("DELETE FROM history WHERE kind='tasks' AND id=?", (identity,))
        return {"id": identity, "deleted": True}

    def delete_upload(self, identity):
        if not re.fullmatch(r"[0-9a-f]{32}", identity):
            raise ValueError("无效的上传资料标识。")
        extension = self.get("uploads", identity).get("extension", ".docx")
        if extension not in (".docx", ".xlsx"):
            raise ValueError("不支持的上传文件格式，未执行删除。")
        path = self.directory / "uploads" / (identity + extension)
        removed = False
        content = b""
        try:
            with self.connection() as db:
                db.execute("BEGIN IMMEDIATE")
                if not db.execute(
                    "SELECT 1 FROM objects WHERE kind='uploads' AND id=?", (identity,)
                ).fetchone():
                    raise KeyError(identity)
                references = []
                for row in db.execute("SELECT payload FROM objects WHERE kind='generation_jobs'"):
                    job = json.loads(row["payload"])
                    if job["status"] in ("queued", "running"):
                        task = job["request"]["task"]
                        if identity == task.get("outline_id") or identity in task.get("record_ids", []):
                            references.append("正在生成的任务：" + task.get("name", "未命名任务"))
                for row in db.execute("SELECT id,payload FROM objects WHERE kind='tasks'"):
                    task = json.loads(row["payload"])
                    if identity == task.get("outline_id") or identity in task.get("record_ids", []):
                        references.append("%s（任务ID：%s）" % (task.get("name", "未命名任务"), row["id"]))
                if references:
                    raise ConflictError("资料仍被以下已保存任务使用，请先打开任务、解除关联并保存："
                                        + "；".join(references))
                if path.is_symlink() or not path.is_file():
                    raise ValueError("上传副本丢失或不是普通文件，未执行删除，请核对本机资料。")
                # Keep a bounded upload backup until the database commit succeeds.
                content = path.read_bytes()
                db.execute("DELETE FROM objects WHERE kind='uploads' AND id=?", (identity,))
                db.execute("DELETE FROM history WHERE kind='uploads' AND id=?", (identity,))
                path.unlink()
                removed = True
        except BaseException:
            if removed:
                with path.open("xb") as target:
                    target.write(content)
            raise
        return {"id": identity, "deleted": True}
