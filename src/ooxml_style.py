from __future__ import annotations

import copy
import hashlib
import json
import posixpath
import re
import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from .design import get_design_value


P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"

REL_PREFIX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

ET.register_namespace("p", P_NS)
ET.register_namespace("r", R_NS)
ET.register_namespace("", PKG_REL_NS)
ET.register_namespace("", CONTENT_TYPES_NS)


CONTENT_TYPES = {
    "ppt/presProps.xml": "application/vnd.openxmlformats-officedocument.presentationml.presProps+xml",
    "ppt/viewProps.xml": "application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml",
    "ppt/tableStyles.xml": "application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml",
    "ppt/theme/theme1.xml": "application/vnd.openxmlformats-officedocument.theme+xml",
    "ppt/theme/theme2.xml": "application/vnd.openxmlformats-officedocument.theme+xml",
    "ppt/slideMasters/slideMaster1.xml": "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
    "ppt/slideMasters/slideMaster2.xml": "application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml",
    "ppt/slideLayouts/slideLayout1.xml": "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
    "ppt/slideLayouts/slideLayout2.xml": "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
    "ppt/slideLayouts/slideLayout3.xml": "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
    "ppt/slideLayouts/slideLayout4.xml": "application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml",
}


def xml_tag(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def style_asset_root(design: dict[str, Any] | None) -> Path | None:
    skill_dir = get_design_value(design, "_skill_dir", None)
    relative = get_design_value(
        design,
        "assets.ooxml_style_parts",
        "assets/concept-practice-style/ooxml-style-parts",
    )
    if not skill_dir or not relative:
        return None
    root = Path(skill_dir) / relative
    return root if root.is_dir() else None


def presentation_metadata_path(design: dict[str, Any] | None) -> Path | None:
    skill_dir = get_design_value(design, "_skill_dir", None)
    relative = get_design_value(
        design,
        "assets.presentation_metadata",
        "assets/concept-practice-style/presentation-metadata.xml",
    )
    if not skill_dir or not relative:
        return None
    path = Path(skill_dir) / relative
    return path if path.is_file() else None


def load_style_manifest(style_root: Path) -> dict[str, Any]:
    manifest_path = style_root / "style-parts-manifest.json"
    return json.loads(manifest_path.read_text(encoding="utf-8"))


def verify_style_manifest(style_root: Path) -> None:
    manifest = load_style_manifest(style_root)
    for entry in manifest.get("parts", []):
        relative = entry.get("part")
        if not relative:
            raise ValueError("style-parts-manifest.json contains an entry without a part field.")
        path = style_root / relative
        if not path.is_file():
            raise FileNotFoundError(f"Style OOXML part listed in manifest is missing: {relative}")
        expected_size = entry.get("size_bytes")
        expected_hash = entry.get("sha256")
        actual_size = path.stat().st_size
        actual_hash = sha256_file(path)
        if expected_size != actual_size:
            raise ValueError(f"Style OOXML part size mismatch for {relative}: {actual_size} != {expected_size}")
        if expected_hash != actual_hash:
            raise ValueError(f"Style OOXML part SHA-256 mismatch for {relative}: {actual_hash} != {expected_hash}")


def copy_manifest_style_parts(style_root: Path, package_root: Path) -> int:
    copied = 0
    manifest = load_style_manifest(style_root)
    for entry in manifest.get("parts", []):
        relative = entry["part"]
        src = style_root / relative
        dst = package_root / relative
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(src, dst)
        copied += 1
    return copied


def next_relationship_id(root: ET.Element) -> str:
    max_id = 0
    for rel in root.findall(xml_tag(PKG_REL_NS, "Relationship")):
        rid = rel.get("Id", "")
        if rid.startswith("rId") and rid[3:].isdigit():
            max_id = max(max_id, int(rid[3:]))
    return f"rId{max_id + 1}"


def ensure_relationship(root: ET.Element, rel_type: str, target: str) -> str:
    for rel in root.findall(xml_tag(PKG_REL_NS, "Relationship")):
        if rel.get("Type") == rel_type and rel.get("Target") == target:
            return rel.get("Id", "")
    rid = next_relationship_id(root)
    ET.SubElement(
        root,
        xml_tag(PKG_REL_NS, "Relationship"),
        Id=rid,
        Type=rel_type,
        Target=target,
    )
    return rid


def ensure_content_type_overrides(content_types_path: Path) -> int:
    tree = ET.parse(content_types_path)
    root = tree.getroot()
    existing = {
        override.get("PartName")
        for override in root.findall(xml_tag(CONTENT_TYPES_NS, "Override"))
    }
    added = 0
    for part, content_type in CONTENT_TYPES.items():
        part_name = f"/{part}"
        if part_name in existing:
            continue
        ET.SubElement(
            root,
            xml_tag(CONTENT_TYPES_NS, "Override"),
            PartName=part_name,
            ContentType=content_type,
        )
        added += 1
    tree.write(content_types_path, encoding="UTF-8", xml_declaration=True)
    return added


def replace_default_text_style(presentation_path: Path, metadata_path: Path | None) -> bool:
    if metadata_path is None:
        return False
    metadata_root = ET.parse(metadata_path).getroot()
    default_text_style = metadata_root.find(xml_tag(P_NS, "defaultTextStyle"))
    if default_text_style is None:
        return False

    tree = ET.parse(presentation_path)
    root = tree.getroot()
    existing = root.find(xml_tag(P_NS, "defaultTextStyle"))
    if existing is not None:
        root.remove(existing)

    insert_after_names = {"notesSz", "sldSz", "sldIdLst", "sldMasterIdLst"}
    insert_at = len(root)
    for idx, child in enumerate(list(root)):
        local = child.tag.split("}", 1)[-1]
        if local not in insert_after_names:
            insert_at = idx
            break
    root.insert(insert_at, copy.deepcopy(default_text_style))
    tree.write(presentation_path, encoding="UTF-8", xml_declaration=True)
    return True


def ensure_presentation_master_relationships(package_root: Path) -> dict[str, str]:
    rel_path = package_root / "ppt/_rels/presentation.xml.rels"
    tree = ET.parse(rel_path)
    root = tree.getroot()
    master1_rid = ensure_relationship(
        root,
        f"{REL_PREFIX}/slideMaster",
        "slideMasters/slideMaster1.xml",
    )
    master2_rid = ensure_relationship(
        root,
        f"{REL_PREFIX}/slideMaster",
        "slideMasters/slideMaster2.xml",
    )
    ensure_relationship(root, f"{REL_PREFIX}/theme", "theme/theme1.xml")
    tree.write(rel_path, encoding="UTF-8", xml_declaration=True)
    return {"slideMaster1": master1_rid, "slideMaster2": master2_rid}


def ensure_master_id_list(package_root: Path, master_relationship_ids: dict[str, str]) -> int:
    path = package_root / "ppt/presentation.xml"
    tree = ET.parse(path)
    root = tree.getroot()
    master_list = root.find(xml_tag(P_NS, "sldMasterIdLst"))
    if master_list is None:
        master_list = ET.Element(xml_tag(P_NS, "sldMasterIdLst"))
        root.insert(0, master_list)

    existing_rids = {
        master.get(xml_tag(R_NS, "id"))
        for master in master_list.findall(xml_tag(P_NS, "sldMasterId"))
    }
    existing_ids = [
        int(master.get("id", "0"))
        for master in master_list.findall(xml_tag(P_NS, "sldMasterId"))
        if master.get("id", "").isdigit()
    ]
    next_id = max(existing_ids or [2147483652]) + 1
    added = 0
    for rid in master_relationship_ids.values():
        if rid in existing_rids:
            continue
        ET.SubElement(
            master_list,
            xml_tag(P_NS, "sldMasterId"),
            id=str(next_id),
            **{xml_tag(R_NS, "id"): rid},
        )
        next_id += 1
        added += 1

    tree.write(path, encoding="UTF-8", xml_declaration=True)
    return added


def slide_number_from_relationship_path(path: Path) -> int | None:
    match = re.search(r"slide(\d+)\.xml\.rels$", path.name)
    return int(match.group(1)) if match else None


def redirect_slide_layout_relationships(
    package_root: Path,
    problem_target: str,
    *,
    title_slide_numbers: set[int] | None = None,
    title_target: str | None = None,
) -> int:
    rel_dir = package_root / "ppt/slides/_rels"
    redirected = 0
    title_slide_numbers = title_slide_numbers or set()
    title_target = title_target or problem_target
    for rel_path in sorted(rel_dir.glob("slide*.xml.rels")):
        slide_number = slide_number_from_relationship_path(rel_path)
        target = title_target if slide_number in title_slide_numbers else problem_target
        tree = ET.parse(rel_path)
        root = tree.getroot()
        changed = False
        for rel in root.findall(xml_tag(PKG_REL_NS, "Relationship")):
            if rel.get("Type") != f"{REL_PREFIX}/slideLayout":
                continue
            if rel.get("Target") != target:
                rel.set("Target", target)
                redirected += 1
                changed = True
        if changed:
            tree.write(rel_path, encoding="UTF-8", xml_declaration=True)
    return redirected


def zip_path_exists(entries: set[str], rels_path: str, target: str) -> bool:
    if target.startswith("/"):
        return target.lstrip("/") in entries
    if "/_rels/" in rels_path:
        owner_dir, rels_file = rels_path.split("/_rels/", 1)
        owner_file = rels_file.removesuffix(".rels")
        base_dir = posixpath.dirname(posixpath.join(owner_dir, owner_file))
    elif rels_path == "_rels/.rels":
        base_dir = ""
    else:
        base_dir = posixpath.dirname(rels_path)
    return posixpath.normpath(posixpath.join(base_dir, target)) in entries


def find_missing_internal_relationship_targets(package_root: Path) -> list[dict[str, str]]:
    entries = {
        path.relative_to(package_root).as_posix()
        for path in package_root.rglob("*")
        if path.is_file()
    }
    missing: list[dict[str, str]] = []
    for rel_path in package_root.rglob("*.rels"):
        rel_name = rel_path.relative_to(package_root).as_posix()
        root = ET.parse(rel_path).getroot()
        for rel in root.findall(xml_tag(PKG_REL_NS, "Relationship")):
            if rel.get("TargetMode") == "External":
                continue
            target = rel.get("Target")
            if target and not zip_path_exists(entries, rel_name, target):
                missing.append(
                    {
                        "relationship_part": rel_name,
                        "relationship_id": rel.get("Id", ""),
                        "target": target,
                    }
                )
    return missing


def rebuild_zip(package_root: Path, pptx_path: Path) -> None:
    rebuilt = pptx_path.with_suffix(".tmp.pptx")
    with zipfile.ZipFile(rebuilt, "w", zipfile.ZIP_DEFLATED) as dst:
        for path in package_root.rglob("*"):
            if path.is_file():
                dst.write(path, path.relative_to(package_root).as_posix())
    shutil.move(str(rebuilt), pptx_path)


def apply_ooxml_style_parts(
    pptx_path: Path,
    design: dict[str, Any] | None,
    *,
    title_slide_numbers: list[int] | None = None,
) -> dict[str, Any]:
    if not get_design_value(design, "ooxml_style.enabled", True):
        return {"applied": False, "reason": "disabled"}

    style_root = style_asset_root(design)
    if style_root is None:
        return {"applied": False, "reason": "style_root_missing"}

    slide_layout_target = get_design_value(
        design,
        "ooxml_style.slide_layout_target",
        "../slideLayouts/slideLayout2.xml",
    )
    problem_slide_layout_target = get_design_value(
        design,
        "ooxml_style.problem_slide_layout_target",
        slide_layout_target,
    )
    title_slide_layout_target = get_design_value(
        design,
        "ooxml_style.title_slide_layout_target",
        slide_layout_target,
    )

    verify_style_manifest(style_root)
    with tempfile.TemporaryDirectory() as tmp:
        package_root = Path(tmp)
        with zipfile.ZipFile(pptx_path) as src:
            bad_entry = src.testzip()
            if bad_entry is not None:
                raise ValueError(f"Invalid PPTX zip entry before style injection: {bad_entry}")
            src.extractall(package_root)

        copied_count = copy_manifest_style_parts(style_root, package_root)
        default_text_style_applied = replace_default_text_style(
            package_root / "ppt/presentation.xml",
            presentation_metadata_path(design),
        )
        master_relationship_ids = ensure_presentation_master_relationships(package_root)
        master_id_count = ensure_master_id_list(package_root, master_relationship_ids)
        content_type_count = ensure_content_type_overrides(package_root / "[Content_Types].xml")
        redirected_slide_rels = redirect_slide_layout_relationships(
            package_root,
            problem_slide_layout_target,
            title_slide_numbers=set(title_slide_numbers or []),
            title_target=title_slide_layout_target,
        )
        missing_targets = find_missing_internal_relationship_targets(package_root)
        if missing_targets:
            raise ValueError(f"Missing internal PPTX relationship targets after style injection: {missing_targets}")

        rebuild_zip(package_root, pptx_path)

    return {
        "applied": True,
        "style_root": str(style_root),
        "copied_part_count": copied_count,
        "default_text_style_applied": default_text_style_applied,
        "presentation_master_relationships": master_relationship_ids,
        "presentation_master_ids_added": master_id_count,
        "content_type_overrides_added": content_type_count,
        "redirected_slide_layout_relationships": redirected_slide_rels,
        "slide_layout_target": slide_layout_target,
        "problem_slide_layout_target": problem_slide_layout_target,
        "title_slide_layout_target": title_slide_layout_target,
        "title_slide_numbers": sorted(title_slide_numbers or []),
    }
