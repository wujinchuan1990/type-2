"""Remove unused image relationships and resources from an output copy."""

from copy import copy
import os
from pathlib import Path
import posixpath
import shutil
import tempfile
from typing import Dict
from urllib.parse import unquote
from zipfile import BadZipFile, ZipFile

from lxml import etree


def _target(rels_name: str, target: str) -> str:
    base = posixpath.dirname(posixpath.dirname(rels_name))
    target = unquote(target.split("#", 1)[0])
    return posixpath.normpath(
        target.lstrip("/") if target.startswith("/") else posixpath.join(base, target))


def _xml(data: bytes):
    root = etree.fromstring(data, etree.XMLParser(resolve_entities=False, no_network=True))
    if root.getroottree().docinfo.doctype:
        raise ValueError("图片清理不支持含DTD的文档部件。")
    return root


def remove_unused_images(path: Path) -> Dict[str, int]:
    """Keep any image still referenced by any retained package part."""
    path = Path(path)
    before_bytes = path.stat().st_size
    temporary = None
    try:
        with ZipFile(path) as source:
            names = source.namelist()
            if len(names) != len(set(names)):
                raise ValueError("DOCX存在重名部件，不能安全清理图片。")
            media = {name for name in names
                     if name.startswith("word/media/") and not name.endswith("/")}
            candidates, protected, relationships = set(), set(), {}
            for name in names:
                if not name.endswith(".rels"):
                    continue
                rels = _xml(source.read(name))
                relationships[name] = rels
                internal = [(rel, _target(name, rel.get("Target", ""))) for rel in rels
                            if rel.get("TargetMode") != "External"]
                if not any(target in media for _, target in internal):
                    continue
                owner = posixpath.join(posixpath.dirname(posixpath.dirname(name)),
                                      posixpath.basename(name)[:-5])
                used = None
                if owner in names and owner.endswith((".xml", ".vml")):
                    root = _xml(source.read(owner))
                    used = {value for element in root.iter() for value in element.attrib.values()}
                for rel, target in internal:
                    if target not in media:
                        continue
                    if (not rel.get("Type", "").endswith("/image")
                            or used is None or rel.get("Id") in used):
                        protected.add(target)
                    else:
                        candidates.add(target)
            removed = candidates - protected
            if not removed:
                return {"removed_images": 0, "before_bytes": before_bytes, "after_bytes": before_bytes}
            if any(name.startswith("_xmlsignatures/") for name in names):
                raise ValueError("DOCX含数字签名，清理图片会使签名失效，停止交付。")

            changes = {}
            for name, rels in relationships.items():
                stale = [rel for rel in rels if rel.get("TargetMode") != "External"
                         and _target(name, rel.get("Target", "")) in removed]
                for rel in stale:
                    rels.remove(rel)
                if stale:
                    changes[name] = etree.tostring(
                        rels, xml_declaration=True, encoding="UTF-8", standalone=True)
            types = _xml(source.read("[Content_Types].xml"))
            stale_types = [node for node in types
                           if unquote(node.get("PartName", "")).lstrip("/") in removed]
            for node in stale_types:
                types.remove(node)
            if stale_types:
                changes["[Content_Types].xml"] = etree.tostring(
                    types, xml_declaration=True, encoding="UTF-8", standalone=True)

            with tempfile.NamedTemporaryFile(
                    prefix=path.name + ".", suffix=".images.tmp", dir=path.parent, delete=False) as handle:
                temporary = Path(handle.name)
            with ZipFile(temporary, "w") as target:
                target.comment = source.comment
                for entry in source.infolist():
                    if entry.filename not in removed:
                        target.writestr(copy(entry), changes.get(entry.filename, source.read(entry.filename)))
            with ZipFile(temporary) as result:
                if result.testzip() is not None or set(result.namelist()) != set(names) - removed:
                    raise ValueError("图片清理后的DOCX完整性校验失败。")
                for name in set(names) - removed - set(changes):
                    if result.read(name) != source.read(name):
                        raise ValueError("图片清理意外改变文档部件：" + name)
        after_bytes = temporary.stat().st_size
        shutil.copymode(path, temporary)
        os.replace(temporary, path)
        return {"removed_images": len(removed), "before_bytes": before_bytes, "after_bytes": after_bytes}
    except (BadZipFile, etree.XMLSyntaxError, KeyError) as exc:
        raise ValueError("无法安全清理DOCX未使用图片：" + str(exc)) from exc
    finally:
        if temporary is not None and temporary.exists():
            temporary.unlink()
