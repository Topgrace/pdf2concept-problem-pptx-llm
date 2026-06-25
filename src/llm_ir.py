from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from .ir import DisplaySegment, LayoutShape, PracticeBlock, PracticeItem, SourceBox


SCHEMA = "pdf2concept-problem-pptx.llm-ir.v1"
BLANK_TOKEN_RE = r"\^\[\s*\]|\[\s*\]|__|□"


def text_square_blank_count(text: str) -> int:
    return len(re.findall(BLANK_TOKEN_RE, text))


def source_box_from_dict(data: dict[str, Any] | None) -> SourceBox | None:
    if not data:
        return None
    return SourceBox(
        x0=float(data["x0"]),
        y0=float(data["y0"]),
        x1=float(data["x1"]),
        y1=float(data["y1"]),
    )


def display_segments_from_dicts(values: list[dict[str, Any]] | None) -> list[DisplaySegment] | None:
    if not values:
        return None
    segments = []
    for value in values:
        gap_after = value.get("gap_after_in")
        segments.append(
            DisplaySegment(
                kind=str(value.get("kind", "math")),
                text=str(value.get("text", "")),
                line_index=int(value.get("line_index", 0)),
                gap_after_in=float(gap_after) if gap_after is not None else None,
                width_in=float(value["width_in"]) if value.get("width_in") is not None else None,
                shape=str(value.get("shape", "")) or None,
            )
        )
    return segments


def int_list_from_value(value: Any) -> list[int]:
    if not isinstance(value, list):
        return []
    result = []
    for item in value:
        try:
            result.append(int(item))
        except (TypeError, ValueError):
            continue
    return result


def float_from_value(value: Any, default: float) -> float:
    if value in (None, ""):
        return default
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def optional_float_from_value(value: Any) -> float | None:
    if value in (None, ""):
        return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def layout_shapes_from_dicts(values: list[dict[str, Any]] | None) -> list[LayoutShape] | None:
    if not values:
        return None
    shapes = []
    for value in values:
        if not isinstance(value, dict):
            continue
        try:
            line_start = int(value.get("line_start", 0))
        except (TypeError, ValueError):
            line_start = 0
        try:
            line_end = int(value.get("line_end", line_start))
        except (TypeError, ValueError):
            line_end = line_start
        shapes.append(
            LayoutShape(
                kind=str(value.get("kind", "")).strip(),
                shape=str(value.get("shape", "")).strip(),
                name=str(value.get("name", "")).strip(),
                line_start=line_start,
                line_end=line_end,
                tick_lines=int_list_from_value(value.get("tick_lines")),
                x_anchor=str(value.get("x_anchor", "before_display_lines")).strip() or "before_display_lines",
                x_offset_in=float_from_value(value.get("x_offset_in"), 0.0),
                x1_in=optional_float_from_value(value.get("x1_in")),
                y1_in=optional_float_from_value(value.get("y1_in")),
                control1_x_in=optional_float_from_value(value.get("control1_x_in")),
                control1_y_in=optional_float_from_value(value.get("control1_y_in")),
                control2_x_in=optional_float_from_value(value.get("control2_x_in")),
                control2_y_in=optional_float_from_value(value.get("control2_y_in")),
                x2_in=optional_float_from_value(value.get("x2_in")),
                y2_in=optional_float_from_value(value.get("y2_in")),
                y_start_offset_in=float_from_value(value.get("y_start_offset_in"), 0.0),
                y_end_offset_in=float_from_value(value.get("y_end_offset_in"), 0.0),
                tick_y_offset_in=float_from_value(value.get("tick_y_offset_in"), 0.0),
                width_in=float_from_value(value.get("width_in"), 0.16),
                stroke_pt=float_from_value(value.get("stroke_pt"), 1.0),
                stroke_color=str(value.get("stroke_color", "#111111")).strip() or "#111111",
                stroke_dash=str(value.get("stroke_dash", "solid")).strip() or "solid",
                arrowhead=str(value.get("arrowhead", "none")).strip() or "none",
            )
        )
    return shapes


def item_from_dict(data: dict[str, Any], fallback_index: int) -> PracticeItem:
    raw_text = str(data.get("raw_text") or "").strip()
    if not raw_text:
        number = data.get("number")
        expression = str(data.get("expression_text") or "").strip()
        if number is None or not expression:
            raise ValueError("Each LLM IR item must include raw_text or both number and expression_text.")
        raw_text = f"({int(number)}) {expression}"

    declared_blank_count = data.get("blank_count")
    actual_blank_count = text_square_blank_count(raw_text)
    if declared_blank_count is not None and int(declared_blank_count) != actual_blank_count:
        raise ValueError(
            f"LLM IR item blank_count mismatch for {raw_text!r}: "
            f"declared {declared_blank_count}, visible tokens {actual_blank_count}."
        )

    source_blanks = [
        source_box_from_dict(value)
        for value in data.get("source_blanks", [])
        if isinstance(value, dict)
    ]
    source_blanks = [box for box in source_blanks if box is not None]
    if source_blanks and len(source_blanks) != actual_blank_count:
        raise ValueError(
            f"LLM IR item source_blanks mismatch for {raw_text!r}: "
            f"{len(source_blanks)} source boxes, {actual_blank_count} visible blank tokens."
        )

    item = PracticeItem.from_text(
        raw_text,
        row_index=int(data.get("row_index", fallback_index)),
        column_index=int(data.get("column_index", 0)),
        source_box=source_box_from_dict(data.get("source_box")),
        source_blanks=source_blanks,
        source_lines=[str(line) for line in data.get("source_lines", [])] or None,
        display_lines=[str(line) for line in data.get("display_lines", [])] or None,
        display_segments=display_segments_from_dicts(data.get("display_segments")),
        layout_shapes=layout_shapes_from_dicts(data.get("layout_shapes")),
    )
    if data.get("expression_text"):
        item.expression_text = str(data["expression_text"]).strip()
    return item


def block_from_dict(data: dict[str, Any]) -> PracticeBlock:
    items = data.get("items", [])
    if not isinstance(items, list):
        raise ValueError("Each LLM IR block must contain an items list.")
    item_models = [
        item_from_dict(item, idx)
        for idx, item in enumerate(items)
        if isinstance(item, dict)
    ]
    return PracticeBlock(
        page=int(data["page"]),
        concept_no=str(data.get("concept_no", "")).strip(),
        concept_title=str(data.get("concept_title", "")).strip(),
        practice_no=str(data.get("practice_no", "")).strip(),
        prompt=str(data.get("prompt", "")).strip(),
        item_models=item_models,
        layout_type=str(data.get("layout_type", "vertical_list")).strip() or "vertical_list",
        source_lines=[str(line) for line in data.get("source_lines", [])],
    )


def load_llm_ir(path: Path) -> tuple[list[PracticeBlock], dict[str, str], str]:
    data = json.loads(path.read_text(encoding="utf-8"))
    schema = data.get("schema")
    if schema != SCHEMA:
        raise ValueError(f"Unsupported LLM IR schema: {schema!r}; expected {SCHEMA!r}.")
    if data.get("extraction_mode") != "llm_only":
        raise ValueError("LLM IR must set extraction_mode to 'llm_only'.")

    blocks_data = data.get("blocks")
    if not isinstance(blocks_data, list):
        raise ValueError("LLM IR must contain a blocks list.")

    blocks = [block_from_dict(block) for block in blocks_data if isinstance(block, dict)]
    blocks.sort(key=lambda block: (block.page, block.practice_no))
    excluded = {
        str(page): str(reason)
        for page, reason in dict(data.get("excluded_pages", {})).items()
    }
    requested_pages = str(data.get("requested_pages", "")).strip()
    if not requested_pages:
        raise ValueError("LLM IR must include requested_pages.")
    return blocks, excluded, requested_pages
