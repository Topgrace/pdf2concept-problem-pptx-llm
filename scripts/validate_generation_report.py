#!/usr/bin/env python3
"""Validate the generator report contract for a concept-practice PPTX build."""

from __future__ import annotations

import argparse
import json
import re
import sys
from pathlib import Path
from typing import Any


def parse_page_range(value: str) -> tuple[int, int]:
    cleaned = value.strip().replace("~", "-")
    match = re.match(r"^(\d+)(?:-(\d+))?$", cleaned)
    if not match:
        raise ValueError(f"Unsupported page range: {value}")
    start = int(match.group(1))
    end = int(match.group(2) or start)
    if start > end:
        raise ValueError(f"Page range start is greater than end: {value}")
    return start, end


def as_int_page(value: Any) -> int | None:
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def item_numbers(slide: dict[str, Any]) -> list[Any]:
    if "item_numbers" in slide:
        return slide["item_numbers"]
    return [item.get("number") for item in slide.get("item_inventory", [])]


def add_failure(failures: list[dict[str, Any]], code: str, message: str, **extra: Any) -> None:
    failures.append({"code": code, "message": message, **extra})


def validate_report(data: dict[str, Any], *, allow_warnings: bool = False) -> dict[str, Any]:
    failures: list[dict[str, Any]] = []

    required = [
        "generator",
        "generation_mode",
        "source_pdf",
        "requested_pages",
        "included_concept_practice_pages",
        "excluded_pages",
        "slide_count",
        "slide_trace",
        "block_inventory",
        "quality_summary",
        "runtime_reference_pptx_used",
        "editable_reconstruction",
        "image_problem_crops_used",
    ]
    for key in required:
        if key not in data:
            add_failure(failures, "missing_required_field", f"Missing required report field: {key}", field=key)

    if failures:
        return {"passes": False, "failures": failures}

    try:
        start_page, end_page = parse_page_range(str(data["requested_pages"]))
        requested_page_set = set(range(start_page, end_page + 1))
    except ValueError as exc:
        add_failure(failures, "invalid_requested_pages", str(exc), requested_pages=data.get("requested_pages"))
        requested_page_set = set()

    included_pages = [as_int_page(page) for page in data.get("included_concept_practice_pages", [])]
    included_pages = [page for page in included_pages if page is not None]
    excluded_pages = [as_int_page(page) for page in data.get("excluded_pages", {}).keys()]
    excluded_pages = [page for page in excluded_pages if page is not None]

    if len(included_pages) != len(set(included_pages)):
        add_failure(
            failures,
            "duplicate_included_pages",
            "included_concept_practice_pages must contain unique page numbers.",
            included_pages=included_pages,
        )
    if requested_page_set:
        outside_included = sorted(set(included_pages) - requested_page_set)
        outside_excluded = sorted(set(excluded_pages) - requested_page_set)
        missing_pages = sorted(requested_page_set - set(included_pages) - set(excluded_pages))
        overlap_pages = sorted(set(included_pages) & set(excluded_pages))
        if outside_included:
            add_failure(
                failures,
                "included_pages_outside_requested_range",
                "Included concept-practice pages must stay inside the requested range.",
                pages=outside_included,
            )
        if outside_excluded:
            add_failure(
                failures,
                "excluded_pages_outside_requested_range",
                "Excluded pages must stay inside the requested range.",
                pages=outside_excluded,
            )
        if missing_pages:
            add_failure(
                failures,
                "requested_pages_not_accounted_for",
                "Every requested page should be either included or explicitly excluded.",
                pages=missing_pages,
            )
        if overlap_pages:
            add_failure(
                failures,
                "pages_both_included_and_excluded",
                "A page cannot be both included and excluded.",
                pages=overlap_pages,
            )

    if data.get("runtime_reference_pptx_used") is not False:
        add_failure(
            failures,
            "runtime_reference_pptx_used",
            "The skill must not depend on a runtime reference PPTX for normal generation.",
        )
    if data.get("editable_reconstruction") is not True:
        add_failure(
            failures,
            "not_editable_reconstruction",
            "The report must mark the output as editable reconstruction.",
        )
    if data.get("image_problem_crops_used") is not False:
        add_failure(
            failures,
            "image_problem_crops_used",
            "Generated decks must not use full problem crops as main slide content.",
        )

    slide_count = data.get("slide_count")
    slide_trace = data.get("slide_trace", [])
    block_inventory = data.get("block_inventory", [])
    quality = data.get("quality_summary", {})

    if not isinstance(slide_count, int) or slide_count <= 0:
        add_failure(failures, "invalid_slide_count", "slide_count must be a positive integer.", slide_count=slide_count)
    if not isinstance(slide_trace, list):
        add_failure(failures, "invalid_slide_trace", "slide_trace must be a list.")
        slide_trace = []
    if not isinstance(block_inventory, list):
        add_failure(failures, "invalid_block_inventory", "block_inventory must be a list.")
        block_inventory = []
    if not isinstance(quality, dict):
        add_failure(failures, "invalid_quality_summary", "quality_summary must be an object.")
        quality = {}

    problem_slide_count = quality.get("problem_slide_count")
    title_slide_count = quality.get("title_slide_count")
    if isinstance(problem_slide_count, int) and problem_slide_count != len(slide_trace):
        add_failure(
            failures,
            "problem_slide_count_mismatch",
            "quality_summary.problem_slide_count must match slide_trace length.",
            problem_slide_count=problem_slide_count,
            slide_trace_count=len(slide_trace),
        )
    if isinstance(slide_count, int) and isinstance(problem_slide_count, int) and isinstance(title_slide_count, int):
        if problem_slide_count + title_slide_count != slide_count:
            add_failure(
                failures,
                "slide_count_breakdown_mismatch",
                "problem_slide_count + title_slide_count must equal slide_count.",
                slide_count=slide_count,
                problem_slide_count=problem_slide_count,
                title_slide_count=title_slide_count,
            )

    source_block_count = quality.get("source_block_count")
    if isinstance(source_block_count, int) and source_block_count != len(block_inventory):
        add_failure(
            failures,
            "source_block_count_mismatch",
            "quality_summary.source_block_count must match block_inventory length.",
            source_block_count=source_block_count,
            block_inventory_count=len(block_inventory),
        )

    block_item_count = sum(len(block.get("items", [])) for block in block_inventory if isinstance(block, dict))
    trace_item_count = sum(len(slide.get("item_inventory", [])) for slide in slide_trace if isinstance(slide, dict))
    source_item_count = quality.get("source_item_count")
    if isinstance(source_item_count, int) and source_item_count != block_item_count:
        add_failure(
            failures,
            "source_item_count_mismatch",
            "quality_summary.source_item_count must match block_inventory item count.",
            source_item_count=source_item_count,
            block_item_count=block_item_count,
        )
    if block_item_count != trace_item_count:
        add_failure(
            failures,
            "trace_item_count_mismatch",
            "slide_trace item_inventory count must match block_inventory item count.",
            block_item_count=block_item_count,
            trace_item_count=trace_item_count,
        )

    slide_numbers: list[int] = []
    required_trace_fields = [
        "slide_number",
        "slide_type",
        "page",
        "practice",
        "layout_type",
        "chunk_index",
        "chunk_count",
        "item_numbers",
        "visible_by_default_item_numbers",
        "click_reveal_item_numbers",
        "item_inventory",
    ]
    chunks_by_block: dict[tuple[Any, Any, Any, Any], list[int]] = {}
    chunk_counts_by_block: dict[tuple[Any, Any, Any, Any], int] = {}
    for index, slide in enumerate(slide_trace, start=1):
        if not isinstance(slide, dict):
            add_failure(failures, "invalid_slide_trace_entry", "Each slide_trace entry must be an object.", index=index)
            continue
        for field in required_trace_fields:
            if field not in slide:
                add_failure(
                    failures,
                    "missing_slide_trace_field",
                    "A slide_trace entry is missing a required field.",
                    index=index,
                    field=field,
                )
        slide_number = slide.get("slide_number")
        if isinstance(slide_number, int):
            slide_numbers.append(slide_number)
            if isinstance(slide_count, int) and not 1 <= slide_number <= slide_count:
                add_failure(
                    failures,
                    "slide_number_out_of_range",
                    "slide_trace slide_number must be inside the deck slide range.",
                    slide_number=slide_number,
                    slide_count=slide_count,
                )
        numbers = item_numbers(slide)
        inventory_numbers = [item.get("number") for item in slide.get("item_inventory", []) if isinstance(item, dict)]
        if numbers != inventory_numbers:
            add_failure(
                failures,
                "slide_item_number_mismatch",
                "slide_trace item_numbers must match its item_inventory item numbers.",
                slide_number=slide_number,
                item_numbers=numbers,
                inventory_numbers=inventory_numbers,
            )
        if numbers:
            expected_visible = [numbers[0]]
            expected_reveal = numbers[1:]
            if slide.get("visible_by_default_item_numbers") != expected_visible:
                add_failure(
                    failures,
                    "visible_by_default_mismatch",
                    "The first item in each problem slide should be visible by default.",
                    slide_number=slide_number,
                    expected=expected_visible,
                    actual=slide.get("visible_by_default_item_numbers"),
                )
            if slide.get("click_reveal_item_numbers") != expected_reveal:
                add_failure(
                    failures,
                    "click_reveal_order_mismatch",
                    "Click reveal item order must match the remaining slide items.",
                    slide_number=slide_number,
                    expected=expected_reveal,
                    actual=slide.get("click_reveal_item_numbers"),
                )
        key = (
            slide.get("page"),
            slide.get("practice"),
            slide.get("concept_no"),
            slide.get("concept_title"),
            slide.get("prompt", ""),
        )
        chunk_index = slide.get("chunk_index")
        chunk_count = slide.get("chunk_count")
        if isinstance(chunk_index, int) and isinstance(chunk_count, int):
            if not 1 <= chunk_index <= chunk_count:
                add_failure(
                    failures,
                    "invalid_chunk_index",
                    "chunk_index must be between 1 and chunk_count.",
                    slide_number=slide_number,
                    chunk_index=chunk_index,
                    chunk_count=chunk_count,
                )
            chunks_by_block.setdefault(key, []).append(chunk_index)
            chunk_counts_by_block[key] = chunk_count

    if len(slide_numbers) != len(set(slide_numbers)):
        add_failure(failures, "duplicate_slide_numbers", "slide_trace must not contain duplicate slide numbers.")
    if slide_numbers != sorted(slide_numbers):
        add_failure(
            failures,
            "slide_trace_not_sorted",
            "slide_trace entries should be sorted by actual slide number.",
            slide_numbers=slide_numbers,
        )
    for key, indexes in chunks_by_block.items():
        chunk_count = chunk_counts_by_block[key]
        expected = list(range(1, chunk_count + 1))
        actual = sorted(indexes)
        if actual != expected:
            add_failure(
                failures,
                "block_chunk_sequence_mismatch",
                "Each block's chunk indexes must cover 1..chunk_count exactly once.",
                block=list(key),
                expected=expected,
                actual=actual,
            )

    warning_count = quality.get("warning_count")
    warnings = quality.get("warnings", [])
    if isinstance(warning_count, int) and isinstance(warnings, list) and warning_count != len(warnings):
        add_failure(
            failures,
            "warning_count_mismatch",
            "quality_summary.warning_count must match warnings length.",
            warning_count=warning_count,
            warnings_length=len(warnings),
        )
    if not allow_warnings and isinstance(warning_count, int) and warning_count > 0:
        add_failure(
            failures,
            "quality_warnings_present",
            "quality_summary contains warnings that must be reviewed or explicitly allowed.",
            warning_count=warning_count,
            warnings=warnings,
        )

    normalization_count = quality.get("normalization_count")
    normalizations = quality.get("normalizations", [])
    if isinstance(normalization_count, int) and isinstance(normalizations, list) and normalization_count != len(normalizations):
        add_failure(
            failures,
            "normalization_count_mismatch",
            "quality_summary.normalization_count must match normalizations length.",
            normalization_count=normalization_count,
            normalizations_length=len(normalizations),
        )

    return {
        "passes": not failures,
        "failures": failures,
        "summary": {
            "requested_pages": data.get("requested_pages"),
            "included_pages": included_pages,
            "excluded_page_count": len(excluded_pages),
            "slide_count": slide_count,
            "problem_slide_count": problem_slide_count,
            "title_slide_count": title_slide_count,
            "source_block_count": source_block_count,
            "source_item_count": source_item_count,
            "warning_count": warning_count,
            "normalization_count": normalization_count,
        },
    }


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(description="Validate an LLM-IR concept-practice generation report JSON file.")
    parser.add_argument("input_report", type=Path)
    parser.add_argument("--allow-warnings", action="store_true", help="Do not fail on quality_summary warnings.")
    parser.add_argument("--report", type=Path, help="Optional path to write this validator's JSON report.")
    args = parser.parse_args(argv)

    data = json.loads(args.input_report.read_text(encoding="utf-8"))
    report = validate_report(data, allow_warnings=args.allow_warnings)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0 if report["passes"] else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
