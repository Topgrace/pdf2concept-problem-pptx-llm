#!/usr/bin/env python3
"""Self-check the pdf2concept-problem-pptx skill package before sharing."""

from __future__ import annotations

import argparse
import hashlib
import json
import py_compile
import re
import sys
from pathlib import Path
from typing import Any


SCRIPT_DIR = Path(__file__).resolve().parent
DEFAULT_SKILL_DIR = SCRIPT_DIR.parent

REQUIRED_FILES = [
    "SKILL.md",
    "assets/design/style-map.json",
    "assets/concept-practice-style/assets-manifest.json",
    "assets/concept-practice-style/ooxml-style-parts/style-parts-manifest.json",
    "references/conversion-checklist.md",
    "references/design-contract.md",
    "scripts/build_concept_practice_deck_from_llm_ir.py",
    "scripts/export_llm_extraction_packet.py",
    "scripts/self_check_skill.py",
    "scripts/validate_editable_deck.py",
    "scripts/validate_generation_report.py",
    "scripts/validate_ooxml_style_parts.py",
    "scripts/validate_reference_pattern.py",
    "src/constants.py",
    "src/design.py",
    "src/ir.py",
    "src/llm_ir.py",
    "src/math_renderer.py",
    "src/ooxml_style.py",
    "src/pdf_utils.py",
    "src/render.py",
    "src/report.py",
    "src/visual_refine.py",
]

FORBIDDEN_PATHS = [
    "reference-inputs",
    "references/unit1-10-19-style-map.md",
    "scripts/build_concept_practice_deck.py",
    "scripts/generate_concept_practice_deck.py",
    "scripts/generate_unit1_10_19.py",
    "src/extract.py",
]

FORBIDDEN_TEXT_PATTERNS = [
    "reference-inputs",
    "기대결과",
    r"[A-Za-z]:\\",
]


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def add_failure(failures: list[dict[str, Any]], code: str, message: str, **extra: Any) -> None:
    failures.append({"code": code, "message": message, **extra})


def text_files(skill_dir: Path) -> list[Path]:
    suffixes = {".md", ".py", ".json", ".yaml", ".yml", ".xml", ".rels"}
    return [
        path
        for path in skill_dir.rglob("*")
        if path.is_file() and path.suffix.lower() in suffixes and "__pycache__" not in path.parts
    ]


def check_required_files(skill_dir: Path, failures: list[dict[str, Any]]) -> None:
    for relative in REQUIRED_FILES:
        path = skill_dir / relative
        if not path.is_file():
            add_failure(
                failures,
                "missing_required_file",
                "A required skill file is missing.",
                path=relative,
            )


def check_forbidden_paths(skill_dir: Path, failures: list[dict[str, Any]]) -> None:
    for relative in FORBIDDEN_PATHS:
        path = skill_dir / relative
        if path.exists():
            add_failure(
                failures,
                "forbidden_path_present",
                "A removed or runtime-forbidden path is still present in the skill package.",
                path=relative,
            )
    pptx_files = [path.relative_to(skill_dir).as_posix() for path in skill_dir.rglob("*.pptx") if path.is_file()]
    if pptx_files:
        add_failure(
            failures,
            "raw_pptx_in_skill",
            "Raw PPTX files must not be bundled in the independent skill package.",
            files=pptx_files,
        )


def check_forbidden_text(skill_dir: Path, failures: list[dict[str, Any]]) -> None:
    matches: list[dict[str, Any]] = []
    for path in text_files(skill_dir):
        if path.name == "self_check_skill.py":
            continue
        relative = path.relative_to(skill_dir).as_posix()
        try:
            content = path.read_text(encoding="utf-8")
        except UnicodeDecodeError:
            continue
        for pattern in FORBIDDEN_TEXT_PATTERNS:
            if re.search(pattern, content):
                matches.append({"file": relative, "pattern": pattern})
    if matches:
        add_failure(
            failures,
            "forbidden_text_reference",
            "Skill text contains a forbidden local dependency reference.",
            matches=matches,
        )


def check_json_files(skill_dir: Path, failures: list[dict[str, Any]]) -> None:
    for path in skill_dir.rglob("*.json"):
        if "__pycache__" in path.parts:
            continue
        try:
            json.loads(path.read_text(encoding="utf-8"))
        except Exception as exc:
            add_failure(
                failures,
                "invalid_json",
                "A JSON file could not be parsed.",
                path=path.relative_to(skill_dir).as_posix(),
                error=str(exc),
            )


def check_python_compile(skill_dir: Path, failures: list[dict[str, Any]]) -> None:
    for folder in ("src", "scripts"):
        root = skill_dir / folder
        if not root.exists():
            continue
        for path in root.rglob("*.py"):
            try:
                py_compile.compile(str(path), doraise=True)
            except py_compile.PyCompileError as exc:
                add_failure(
                    failures,
                    "python_compile_failed",
                    "A Python source file failed to compile.",
                    path=path.relative_to(skill_dir).as_posix(),
                    error=str(exc),
                )


def check_assets_manifest(skill_dir: Path, failures: list[dict[str, Any]]) -> None:
    manifest_path = skill_dir / "assets/concept-practice-style/assets-manifest.json"
    if not manifest_path.exists():
        return
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    asset_root = manifest_path.parent
    entries = list(manifest.get("assets", [])) + list(manifest.get("metadata_assets", []))
    for entry in entries:
        relative = entry.get("file")
        if not relative:
            add_failure(failures, "asset_manifest_entry_missing_file", "An asset manifest entry has no file field.")
            continue
        path = asset_root / relative
        if not path.is_file():
            add_failure(
                failures,
                "manifest_asset_missing",
                "A file listed in assets-manifest.json is missing.",
                path=str(relative),
            )
            continue
        actual_size = path.stat().st_size
        actual_hash = sha256_file(path)
        if entry.get("size_bytes") != actual_size:
            add_failure(
                failures,
                "manifest_asset_size_mismatch",
                "A manifest asset file size does not match.",
                path=str(relative),
                expected=entry.get("size_bytes"),
                actual=actual_size,
            )
        if entry.get("sha256") != actual_hash:
            add_failure(
                failures,
                "manifest_asset_hash_mismatch",
                "A manifest asset SHA-256 does not match.",
                path=str(relative),
                expected=entry.get("sha256"),
                actual=actual_hash,
            )


def check_style_parts_manifest(skill_dir: Path, failures: list[dict[str, Any]]) -> None:
    manifest_path = skill_dir / "assets/concept-practice-style/ooxml-style-parts/style-parts-manifest.json"
    if not manifest_path.exists():
        return
    manifest = json.loads(manifest_path.read_text(encoding="utf-8"))
    asset_root = manifest_path.parent
    for entry in manifest.get("parts", []):
        relative = entry.get("part")
        if not relative:
            add_failure(failures, "style_part_manifest_entry_missing_part", "A style part manifest entry has no part field.")
            continue
        path = asset_root / relative
        if not path.is_file():
            add_failure(
                failures,
                "manifest_style_part_missing",
                "A file listed in style-parts-manifest.json is missing.",
                path=str(relative),
            )
            continue
        actual_size = path.stat().st_size
        actual_hash = sha256_file(path)
        if entry.get("size_bytes") != actual_size:
            add_failure(
                failures,
                "manifest_style_part_size_mismatch",
                "A style part file size does not match.",
                path=str(relative),
                expected=entry.get("size_bytes"),
                actual=actual_size,
            )
        if entry.get("sha256") != actual_hash:
            add_failure(
                failures,
                "manifest_style_part_hash_mismatch",
                "A style part SHA-256 does not match.",
                path=str(relative),
                expected=entry.get("sha256"),
                actual=actual_hash,
            )


def check_design_contract(skill_dir: Path, failures: list[dict[str, Any]]) -> None:
    path = skill_dir / "assets/design/style-map.json"
    if not path.exists():
        return
    data = json.loads(path.read_text(encoding="utf-8"))
    if data.get("schema") != "pdf2concept-problem-pptx.design.v1":
        add_failure(
            failures,
            "unexpected_design_schema",
            "The design map schema is missing or unexpected.",
            expected="pdf2concept-problem-pptx.design.v1",
            actual=data.get("schema"),
        )
    for key in ("canvas", "fonts", "assets", "ooxml_style", "colors", "problem_slide", "animation", "extraction"):
        if key not in data:
            add_failure(
                failures,
                "design_map_missing_section",
                "The design map is missing a required section.",
                section=key,
            )


def run_self_check(skill_dir: Path) -> dict[str, Any]:
    skill_dir = skill_dir.resolve()
    failures: list[dict[str, Any]] = []
    check_required_files(skill_dir, failures)
    check_forbidden_paths(skill_dir, failures)
    check_forbidden_text(skill_dir, failures)
    check_json_files(skill_dir, failures)
    check_python_compile(skill_dir, failures)
    check_assets_manifest(skill_dir, failures)
    check_style_parts_manifest(skill_dir, failures)
    check_design_contract(skill_dir, failures)
    return {
        "passes": not failures,
        "skill_dir": str(skill_dir),
        "failure_count": len(failures),
        "failures": failures,
        "summary": {
            "required_file_count": len(REQUIRED_FILES),
            "python_file_count": len(list((skill_dir / "src").glob("*.py"))) + len(list((skill_dir / "scripts").glob("*.py"))),
            "text_file_count": len(text_files(skill_dir)),
        },
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Self-check the pdf2concept-problem-pptx skill package.")
    parser.add_argument("--skill-dir", type=Path, default=DEFAULT_SKILL_DIR)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args(argv)

    report = run_self_check(args.skill_dir)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0 if report["passes"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
