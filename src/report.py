from __future__ import annotations

from datetime import datetime, timezone
from pathlib import Path
import re

from .ir import PracticeBlock


def text_square_blank_count(text: str) -> int:
    return len(re.findall(r"\^\[\s*\]|\[\s*\]|__|□", text))


def ordered_unique(values: list[int]) -> list[int]:
    seen: set[int] = set()
    result: list[int] = []
    for value in values:
        if value not in seen:
            seen.add(value)
            result.append(value)
    return result


def build_quality_summary(blocks: list[PracticeBlock], deck_info: dict) -> dict:
    warnings: list[dict] = []
    normalizations: list[dict] = []
    source_item_count = 0
    square_blank_count = 0
    parenthesis_blank_count = 0
    fraction_item_count = 0
    exponent_item_count = 0
    multiline_item_count = 0

    for block in blocks:
        block_id = f"p{block.page}-practice{block.practice_no}"
        if block.layout_type == "unknown":
            warnings.append(
                {
                    "code": "unknown_layout_type",
                    "block": block_id,
                    "message": "The extractor could not classify the source layout; inspect the generated slide visually.",
                }
            )
        if not block.prompt.strip():
            warnings.append(
                {
                    "code": "empty_prompt",
                    "block": block_id,
                    "message": "The concept-practice block has no prompt text.",
                }
            )
        if not block.item_models:
            warnings.append(
                {
                    "code": "empty_item_inventory",
                    "block": block_id,
                    "message": "The concept-practice block has no extracted subproblem items.",
                }
            )

        numbers = [item.number for item in block.item_models if item.number is not None]
        if numbers and len(numbers) != len(set(numbers)):
            warnings.append(
                {
                    "code": "duplicate_item_numbers",
                    "block": block_id,
                    "item_numbers": numbers,
                    "message": "Duplicate subproblem numbers were extracted from one source block.",
                }
            )
        if numbers and numbers != sorted(numbers):
            entry = {
                "code": "row_major_render_order_applied",
                "block": block_id,
                "item_numbers": numbers,
                "message": "Source text order differs from numeric order; rendered order is normalized by row and column.",
            }
            if block.layout_type == "two_column_grid":
                normalizations.append(entry)
            else:
                warnings.append(entry)

        for item in block.item_models:
            source_item_count += 1
            square_blank_count += item.blank_count
            parenthesis_blank_count += int(item.has_parenthesis_blank)
            fraction_item_count += int(item.fraction_count > 0)
            exponent_item_count += int(item.has_exponent)
            multiline_item_count += int(len([line for line in item.source_lines if line.strip()]) > 1)
            source_blank_count = len(item.source_blanks)
            if source_blank_count:
                token_blank_count = text_square_blank_count(item.raw_text)
                if item.blank_count != source_blank_count or token_blank_count != source_blank_count:
                    warnings.append(
                        {
                            "code": "source_blank_count_mismatch",
                            "block": block_id,
                            "item_number": item.number,
                            "blank_count": item.blank_count,
                            "source_blank_count": source_blank_count,
                            "text_blank_count": token_blank_count,
                            "raw_text": item.raw_text,
                            "message": "A source item has a different PDF blank-rect count, IR blank_count, or visible blank-token count.",
                        }
                    )
            if item.number is None:
                warnings.append(
                    {
                        "code": "missing_item_number",
                        "block": block_id,
                        "raw_text": item.raw_text,
                        "message": "A subproblem item was extracted without a numeric label.",
                    }
                )
            if not item.source_lines:
                warnings.append(
                    {
                        "code": "missing_item_source_lines",
                        "block": block_id,
                        "item_number": item.number,
                        "message": "A subproblem item has no preserved source line trace.",
                    }
                )

    traced_item_count = sum(len(slide.get("item_inventory", [])) for slide in deck_info.get("slide_trace", []))
    if traced_item_count != source_item_count:
        warnings.append(
            {
                "code": "item_trace_count_mismatch",
                "source_item_count": source_item_count,
                "traced_item_count": traced_item_count,
                "message": "The rendered slide trace does not cover the same number of source items.",
            }
        )

    problem_slide_count = len(deck_info.get("slide_trace", []))
    title_slide_count = deck_info.get("slide_count", 0) - problem_slide_count
    paginated_block_count = len(
        {
            (slide.get("page"), slide.get("practice"))
            for slide in deck_info.get("slide_trace", [])
            if slide.get("chunk_count", 1) > 1
        }
    )
    row_split_slide_count = sum(
        1
        for slide in deck_info.get("slide_trace", [])
        if slide.get("layout_type") == "two_column_grid" and slide.get("chunk_count", 1) > 1
    )

    return {
        "source_block_count": len(blocks),
        "source_item_count": source_item_count,
        "problem_slide_count": problem_slide_count,
        "title_slide_count": title_slide_count,
        "paginated_block_count": paginated_block_count,
        "two_column_row_split_slide_count": row_split_slide_count,
        "source_square_blank_count": square_blank_count,
        "source_parenthesis_blank_item_count": parenthesis_blank_count,
        "source_fraction_item_count": fraction_item_count,
        "source_exponent_item_count": exponent_item_count,
        "source_multiline_item_count": multiline_item_count,
        "warning_count": len(warnings),
        "warnings": warnings,
        "normalization_count": len(normalizations),
        "normalizations": normalizations,
    }


def build_report(
    pdf_path: Path,
    pages: str,
    source_sha256: str,
    blocks: list[PracticeBlock],
    excluded: dict,
    deck_info: dict,
) -> dict:
    return {
        "generator": "build_concept_practice_deck_from_llm_ir.py",
        "pipeline": "rendered-page-images -> llm-authored-practice-block-ir -> editable-pptx-renderer -> ooxml-style-injection -> validation",
        "generation_mode": deck_info.get("generation_mode", "llm_only_ir_extraction"),
        "generated_at": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "source_pdf": {
            "path": str(pdf_path),
            "size": pdf_path.stat().st_size,
            "sha256": source_sha256,
        },
        "requested_pages": pages,
        "included_concept_practice_pages": ordered_unique([block.page for block in blocks]),
        "excluded_pages": excluded,
        "slide_count": deck_info["slide_count"],
        "ooxml_style": deck_info.get("ooxml_style", {}),
        "slide_trace": deck_info["slide_trace"],
        "block_inventory": [block.to_inventory() for block in blocks],
        "quality_summary": build_quality_summary(blocks, deck_info),
        "runtime_reference_pptx_used": False,
        "editable_reconstruction": True,
        "image_problem_crops_used": False,
        "limitations": [
            "IR quality depends on the LLM-authored llm-ir.json matching the rendered source page images.",
            "Reference-level visual parity requires adding measured layout rules to assets/design/style-map.json and renderer support for each source layout family.",
        ],
    }
