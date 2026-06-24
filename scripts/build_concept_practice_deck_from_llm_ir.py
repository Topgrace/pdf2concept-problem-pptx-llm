#!/usr/bin/env python3
"""Build and validate a concept-practice PPTX from LLM-only IR JSON."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from types import SimpleNamespace

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
if str(SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(SKILL_DIR))

from src.design import load_design
from src.llm_ir import load_llm_ir
from src.pdf_utils import sha256_file
from src.render import build_presentation
from src.report import build_report
from src.visual_refine import (
    apply_design_adjustments,
    export_pptx_pngs,
    inspect_exponent_blank_geometry,
    make_comparison_images,
    next_adjustments,
    render_pdf_pages,
)
from validate_editable_deck import inspect_deck as inspect_editable_deck
from validate_generation_report import validate_report
from validate_ooxml_style_parts import inspect_deck as inspect_ooxml_style_deck
from validate_reference_pattern import inspect_deck as inspect_reference_pattern
from validate_reference_pattern import validate as validate_reference_pattern


def write_json(path: Path, data: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def decide_require_blank_shapes(generation_report: dict, explicit: bool | None) -> bool:
    if explicit is not None:
        return explicit
    quality = generation_report.get("quality_summary", {})
    return int(quality.get("source_square_blank_count", 0) or 0) > 0


def generate_deck_from_llm_ir(
    pdf_path: Path,
    ir_path: Path,
    output_path: Path,
    *,
    design_adjustments: dict | None = None,
) -> dict:
    blocks, excluded, requested_pages = load_llm_ir(ir_path)
    if not blocks:
        raise SystemExit("No concept-practice blocks were found in the LLM IR.")

    design = load_design(SKILL_DIR)
    if design_adjustments:
        design = apply_design_adjustments(design, design_adjustments)
    deck_info = build_presentation(blocks, output_path, design)
    deck_info["generation_mode"] = "llm_only_ir_extraction"
    report = build_report(
        pdf_path,
        requested_pages,
        sha256_file(pdf_path),
        blocks,
        excluded,
        deck_info,
    )
    report["generator"] = "build_concept_practice_deck_from_llm_ir.py"
    report["extraction_mode"] = "llm_only"
    report["llm_ir"] = {
        "path": str(ir_path),
        "schema": "pdf2concept-problem-pptx.llm-ir.v1",
    }
    return report


def run_visual_refinement_loop(
    pdf_path: Path,
    ir_path: Path,
    output_path: Path,
    report_dir: Path,
    *,
    max_iterations: int,
) -> tuple[dict, dict]:
    visual_dir = report_dir / "visual-refinement"
    visual_dir.mkdir(parents=True, exist_ok=True)
    base_design = load_design(SKILL_DIR)
    requested_blocks, _, requested_pages = load_llm_ir(ir_path)
    source_pages = sorted({block.page for block in requested_blocks})
    source_images = render_pdf_pages(pdf_path, source_pages, visual_dir / "source-pages")

    adjustments: dict[str, float] = {}
    iterations: list[dict] = []
    final_report: dict | None = None
    final_diagnostics: dict | None = None
    final_export_backend = None
    final_export_error = None
    final_used_adjustments: dict[str, float] = {}

    for iteration in range(1, max(1, max_iterations) + 1):
        generation_report = generate_deck_from_llm_ir(
            pdf_path,
            ir_path,
            output_path,
            design_adjustments=adjustments,
        )
        active_design = apply_design_adjustments(base_design, adjustments)
        iteration_dir = visual_dir / f"iteration-{iteration:02d}"
        slide_images, export_backend, export_error = export_pptx_pngs(output_path, iteration_dir / "slides")
        comparisons = make_comparison_images(source_images, slide_images, iteration_dir / "comparisons")
        diagnostics = inspect_exponent_blank_geometry(output_path, generation_report, active_design)
        passed = diagnostics.get("passes") is True
        next_values = {} if passed else next_adjustments(adjustments, diagnostics)
        entry = {
            "iteration": iteration,
            "passes": passed,
            "pptx": str(output_path),
            "source_pngs": source_images,
            "slide_pngs": slide_images,
            "comparison_pngs": comparisons,
            "export_backend": export_backend,
            "export_warning": export_error,
            "design_adjustments": adjustments,
            "diagnostics": diagnostics,
            "next_adjustments": next_values,
        }
        iterations.append(entry)
        final_report = generation_report
        final_diagnostics = diagnostics
        final_export_backend = export_backend
        final_export_error = export_error
        final_used_adjustments = dict(adjustments)
        if passed:
            break
        adjustments = next_values

    visual_report = {
        "passes": bool(final_diagnostics and final_diagnostics.get("passes") is True),
        "requested_pages": requested_pages,
        "max_iterations": max_iterations,
        "iteration_count": len(iterations),
        "final_design_adjustments": final_used_adjustments,
        "final_export_backend": final_export_backend,
        "final_export_warning": final_export_error,
        "iterations": iterations,
    }
    write_json(report_dir / "visual-refinement-report.json", visual_report)
    if final_report is None:
        raise RuntimeError("Visual refinement loop did not produce a generation report.")
    final_report["visual_refinement"] = {
        "enabled": True,
        "passes": visual_report["passes"],
        "report": str(report_dir / "visual-refinement-report.json"),
        "iteration_count": len(iterations),
        "final_design_adjustments": final_used_adjustments,
    }
    return final_report, visual_report


def run_build(
    pdf_path: Path,
    ir_path: Path,
    output_path: Path,
    report_dir: Path,
    *,
    require_blank_shapes: bool | None = None,
    allow_report_warnings: bool = False,
    visual_refine_loop: bool = False,
    max_iterations: int = 3,
) -> dict:
    if not pdf_path.exists():
        raise FileNotFoundError(f"Input PDF not found: {pdf_path}")
    if not ir_path.exists():
        raise FileNotFoundError(f"LLM IR JSON not found: {ir_path}")

    if visual_refine_loop:
        generation_report, visual_report = run_visual_refinement_loop(
            pdf_path,
            ir_path,
            output_path,
            report_dir,
            max_iterations=max_iterations,
        )
    else:
        generation_report = generate_deck_from_llm_ir(pdf_path, ir_path, output_path)
        visual_report = None
    write_json(report_dir / "generation-report.json", generation_report)

    generation_validation = validate_report(generation_report, allow_warnings=allow_report_warnings)
    write_json(report_dir / "generation-report-validation.json", generation_validation)

    editable_report = inspect_editable_deck(output_path)
    write_json(report_dir / "editable-report.json", editable_report)

    ooxml_style_report = inspect_ooxml_style_deck(output_path, SKILL_DIR)
    write_json(report_dir / "ooxml-style-report.json", ooxml_style_report)

    require_blanks = decide_require_blank_shapes(generation_report, require_blank_shapes)
    reference_args = SimpleNamespace(
        require_animations=True,
        require_group_animation=True,
        require_blank_shapes=require_blanks,
        forbid_visible_page_labels=True,
        forbid_off_slide_objects=True,
    )
    pattern_report = validate_reference_pattern(inspect_reference_pattern(output_path), reference_args)
    write_json(report_dir / "pattern-report.json", pattern_report)

    validations = {
        "generation_report": generation_validation.get("passes") is True,
        "editable_deck": editable_report.get("passes") is True,
        "ooxml_style": ooxml_style_report.get("passes") is True,
        "reference_pattern": pattern_report.get("passes") is True,
    }
    if visual_report is not None:
        validations["visual_refinement"] = visual_report.get("passes") is True
    build_report = {
        "passes": all(validations.values()),
        "output_pptx": str(output_path),
        "report_dir": str(report_dir),
        "require_blank_shapes": require_blanks,
        "validations": validations,
        "generation_summary": generation_validation.get("summary", {}),
        "editable_failures": editable_report.get("failures", []),
        "ooxml_style_failures": ooxml_style_report.get("failures", []),
        "pattern_failures": pattern_report.get("failures", []),
        "generation_failures": generation_validation.get("failures", []),
        "visual_refinement": visual_report or {"enabled": False},
    }
    write_json(report_dir / "build-summary.json", build_report)
    return build_report


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Build a concept-practice PPTX from LLM-only IR JSON and run standard validation."
    )
    parser.add_argument("--pdf", required=True, type=Path)
    parser.add_argument("--ir", required=True, type=Path, help="LLM-only IR JSON.")
    parser.add_argument("--output", required=True, type=Path)
    parser.add_argument("--report-dir", type=Path)
    parser.add_argument(
        "--require-blank-shapes",
        dest="require_blank_shapes",
        action="store_true",
        default=None,
        help="Force reference-pattern validation to require blank shapes.",
    )
    parser.add_argument(
        "--no-require-blank-shapes",
        dest="require_blank_shapes",
        action="store_false",
        help="Force reference-pattern validation not to require blank shapes.",
    )
    parser.add_argument(
        "--allow-report-warnings",
        action="store_true",
        help="Do not fail the generation report validation when quality_summary warnings are present.",
    )
    parser.add_argument(
        "--visual-refine-loop",
        action="store_true",
        help="Render source/PPTX PNGs and iterate exponent-blank visual diagnostics before final validation.",
    )
    parser.add_argument(
        "--max-iterations",
        type=int,
        default=3,
        help="Maximum visual refinement iterations when --visual-refine-loop is enabled.",
    )
    args = parser.parse_args(argv)

    report_dir = args.report_dir or args.output.parent / f"{args.output.stem}_reports"
    build_report = run_build(
        args.pdf,
        args.ir,
        args.output,
        report_dir,
        require_blank_shapes=args.require_blank_shapes,
        allow_report_warnings=args.allow_report_warnings,
        visual_refine_loop=args.visual_refine_loop,
        max_iterations=args.max_iterations,
    )
    print(json.dumps(build_report, ensure_ascii=False, indent=2))
    return 0 if build_report["passes"] else 1


if __name__ == "__main__":
    raise SystemExit(main())
