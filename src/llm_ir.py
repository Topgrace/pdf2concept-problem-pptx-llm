from __future__ import annotations

import json
import re
from pathlib import Path
from typing import Any

from .ir import DisplaySegment, PracticeBlock, PracticeItem, SourceBox


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
                shape=str(value.get("shape", "")) or None,
            )
        )
    return segments


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
