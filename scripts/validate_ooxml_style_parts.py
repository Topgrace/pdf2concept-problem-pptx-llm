#!/usr/bin/env python3
"""Validate that generated PPTX files contain the built-in OOXML style parts."""

from __future__ import annotations

import argparse
import hashlib
import json
import posixpath
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_SKILL_DIR = SCRIPT_DIR.parent

PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
REL_PREFIX = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
P_NS = "http://schemas.openxmlformats.org/presentationml/2006/main"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"


def xml_tag(namespace: str, name: str) -> str:
    return f"{{{namespace}}}{name}"


def sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def style_manifest(skill_dir: Path) -> dict:
    manifest_path = skill_dir / "assets/concept-practice-style/ooxml-style-parts/style-parts-manifest.json"
    return json.loads(manifest_path.read_text(encoding="utf-8"))


def style_map(skill_dir: Path) -> dict:
    style_path = skill_dir / "assets/design/style-map.json"
    return json.loads(style_path.read_text(encoding="utf-8"))


def style_value(data: dict, path: str, default: str) -> str:
    current = data
    for part in path.split("."):
        if not isinstance(current, dict) or part not in current:
            return default
        current = current[part]
    return current if isinstance(current, str) else default


def relationship_owner_base(rels_path: str) -> str:
    if rels_path == "_rels/.rels":
        return ""
    if "/_rels/" in rels_path:
        owner_dir, rels_file = rels_path.split("/_rels/", 1)
        owner_file = rels_file.removesuffix(".rels")
        return posixpath.dirname(posixpath.join(owner_dir, owner_file))
    return posixpath.dirname(rels_path)


def resolve_relationship_target(rels_path: str, target: str) -> str:
    if target.startswith("/"):
        return target.lstrip("/")
    return posixpath.normpath(posixpath.join(relationship_owner_base(rels_path), target))


def slide_number(path: str) -> int:
    match = re.search(r"slide(\d+)\.xml\.rels$", path)
    return int(match.group(1)) if match else 0


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def slide_text(deck: zipfile.ZipFile, number: int) -> str:
    slide_path = f"ppt/slides/slide{number}.xml"
    if slide_path not in deck.namelist():
        return ""
    root = ET.fromstring(deck.read(slide_path))
    return clean("".join(node.text or "" for node in root.findall(f".//{{{A_NS}}}t")))


def is_title_slide(deck: zipfile.ZipFile, number: int) -> bool:
    return "개념핵심" in slide_text(deck, number)


def inspect_deck(
    pptx_path: Path,
    skill_dir: Path = DEFAULT_SKILL_DIR,
    *,
    expected_slide_layout_target: str | None = None,
    expected_title_slide_layout_target: str | None = None,
) -> dict:
    manifest = style_manifest(skill_dir)
    design = style_map(skill_dir)
    default_layout_target = style_value(design, "ooxml_style.slide_layout_target", "../slideLayouts/slideLayout2.xml")
    expected_problem_layout_target = expected_slide_layout_target or style_value(
        design,
        "ooxml_style.problem_slide_layout_target",
        default_layout_target,
    )
    expected_title_layout_target = expected_title_slide_layout_target or style_value(
        design,
        "ooxml_style.title_slide_layout_target",
        expected_problem_layout_target,
    )
    with zipfile.ZipFile(pptx_path) as deck:
        bad_entry = deck.testzip()
        if bad_entry is not None:
            raise ValueError(f"Invalid PPTX zip entry: {bad_entry}")

        entries = set(deck.namelist())
        failures = []
        style_parts = []
        for entry in manifest.get("parts", []):
            part = entry.get("part")
            if not part:
                failures.append({"rule": "manifest_entry_has_part", "entry": entry})
                continue
            if part not in entries:
                style_parts.append({"part": part, "exists": False, "hash_matches": False})
                failures.append({"rule": "style_part_exists", "part": part})
                continue
            data = deck.read(part)
            hash_matches = sha256_bytes(data) == entry.get("sha256")
            size_matches = len(data) == entry.get("size_bytes")
            style_parts.append(
                {
                    "part": part,
                    "exists": True,
                    "size_matches": size_matches,
                    "hash_matches": hash_matches,
                }
            )
            if not size_matches or not hash_matches:
                failures.append(
                    {
                        "rule": "style_part_matches_manifest",
                        "part": part,
                        "size_matches": size_matches,
                        "hash_matches": hash_matches,
                    }
                )

        content_types_root = ET.fromstring(deck.read("[Content_Types].xml"))
        content_type_overrides = {
            override.get("PartName")
            for override in content_types_root.findall(xml_tag(CONTENT_TYPES_NS, "Override"))
        }
        missing_content_type_overrides = [
            f"/{entry['part']}"
            for entry in manifest.get("parts", [])
            if entry.get("part", "").endswith(".xml") and f"/{entry['part']}" not in content_type_overrides
        ]
        if missing_content_type_overrides:
            failures.append(
                {
                    "rule": "content_type_override_exists",
                    "missing": missing_content_type_overrides,
                }
            )

        slide_layout_relationships = []
        for rel_path in sorted(
            [
                name
                for name in entries
                if re.fullmatch(r"ppt/slides/_rels/slide\d+\.xml\.rels", name)
            ],
            key=slide_number,
        ):
            root = ET.fromstring(deck.read(rel_path))
            current_slide_number = slide_number(rel_path)
            title_slide = is_title_slide(deck, current_slide_number)
            expected_target = expected_title_layout_target if title_slide else expected_problem_layout_target
            for rel in root.findall(xml_tag(PKG_REL_NS, "Relationship")):
                if rel.get("Type") == f"{REL_PREFIX}/slideLayout":
                    slide_layout_relationships.append(
                        {
                            "relationship_part": rel_path,
                            "relationship_id": rel.get("Id"),
                            "slide": current_slide_number,
                            "slide_type": "title" if title_slide else "problem",
                            "target": rel.get("Target"),
                            "expected": expected_target,
                        }
                    )
                    if rel.get("Target") != expected_target:
                        failures.append(
                            {
                                "rule": "slide_layout_target_matches_style_contract",
                                "relationship_part": rel_path,
                                "slide": current_slide_number,
                                "slide_type": "title" if title_slide else "problem",
                                "target": rel.get("Target"),
                                "expected": expected_target,
                            }
                        )

        presentation_relationships = []
        presentation_root = ET.fromstring(deck.read("ppt/_rels/presentation.xml.rels"))
        for rel in presentation_root.findall(xml_tag(PKG_REL_NS, "Relationship")):
            if rel.get("Type") == f"{REL_PREFIX}/slideMaster":
                presentation_relationships.append(
                    {
                        "relationship_id": rel.get("Id"),
                        "target": rel.get("Target"),
                    }
                )
        master_targets = {rel["target"] for rel in presentation_relationships}
        for required in ("slideMasters/slideMaster1.xml", "slideMasters/slideMaster2.xml"):
            if required not in master_targets:
                failures.append(
                    {
                        "rule": "presentation_master_relationship_exists",
                        "target": required,
                    }
                )

        missing_internal_relationship_targets = []
        for rel_path in sorted(name for name in entries if name.endswith(".rels")):
            root = ET.fromstring(deck.read(rel_path))
            for rel in root.findall(xml_tag(PKG_REL_NS, "Relationship")):
                if rel.get("TargetMode") == "External":
                    continue
                target = rel.get("Target")
                if not target:
                    continue
                resolved = resolve_relationship_target(rel_path, target)
                if resolved not in entries:
                    missing_internal_relationship_targets.append(
                        {
                            "relationship_part": rel_path,
                            "relationship_id": rel.get("Id"),
                            "target": target,
                            "resolved_target": resolved,
                        }
                    )
        if missing_internal_relationship_targets:
            failures.append(
                {
                    "rule": "internal_relationship_targets_exist",
                    "missing": missing_internal_relationship_targets,
                }
            )

    return {
        "path": str(pptx_path),
        "style_part_count": len(style_parts),
        "style_parts": style_parts,
        "slide_layout_relationships": slide_layout_relationships,
        "presentation_master_relationships": presentation_relationships,
        "missing_content_type_overrides": missing_content_type_overrides,
        "missing_internal_relationship_targets": missing_internal_relationship_targets,
        "failures": failures,
        "passes": not failures,
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Validate embedded OOXML theme/master/layout style parts in a PPTX.")
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--skill-dir", type=Path, default=DEFAULT_SKILL_DIR)
    parser.add_argument("--expected-slide-layout-target")
    parser.add_argument("--expected-title-slide-layout-target")
    parser.add_argument("--report", type=Path)
    args = parser.parse_args(argv)

    report = inspect_deck(
        args.pptx,
        args.skill_dir,
        expected_slide_layout_target=args.expected_slide_layout_target,
        expected_title_slide_layout_target=args.expected_title_slide_layout_target,
    )
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0 if report["passes"] else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
