from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pptx.dml.color import RGBColor


def load_design(skill_dir: Path) -> dict[str, Any]:
    """Load the skill-level design contract.

    `skill_dir` may be either the skill root or the JSON file itself.
    Missing design files return an empty mapping so older tests can still run.
    """

    design_path = skill_dir
    if skill_dir.is_dir():
        design_path = skill_dir / "assets" / "design" / "style-map.json"
    if not design_path.exists():
        return {}
    design = json.loads(design_path.read_text(encoding="utf-8"))
    if skill_dir.is_dir():
        design["_skill_dir"] = str(skill_dir)
    return design


def get_design_value(design: dict[str, Any] | None, dotted_path: str, default: Any) -> Any:
    node: Any = design or {}
    for part in dotted_path.split("."):
        if not isinstance(node, dict) or part not in node:
            return default
        node = node[part]
    return node


def rgb_from_hex(value: str | None, default: RGBColor) -> RGBColor:
    if not value:
        return default
    normalized = value.strip().lstrip("#")
    if len(normalized) != 6:
        return default
    try:
        return RGBColor(
            int(normalized[0:2], 16),
            int(normalized[2:4], 16),
            int(normalized[4:6], 16),
        )
    except ValueError:
        return default
