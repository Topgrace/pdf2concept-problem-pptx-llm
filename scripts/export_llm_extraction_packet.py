#!/usr/bin/env python3
"""Export source page images and schema for LLM-only IR extraction."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

import fitz

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
if str(SKILL_DIR) not in sys.path:
    sys.path.insert(0, str(SKILL_DIR))

from src.llm_ir import SCHEMA
from src.pdf_utils import parse_page_range


def render_page_images(pdf_path: Path, pages: list[int], output_dir: Path, dpi: int) -> list[dict]:
    image_dir = output_dir / "pages"
    image_dir.mkdir(parents=True, exist_ok=True)
    doc = fitz.open(pdf_path)
    matrix = fitz.Matrix(dpi / 72, dpi / 72)
    rendered: list[dict] = []
    for page_no in pages:
        page = doc[page_no - 1]
        image_path = image_dir / f"page-{page_no:03d}.png"
        pix = page.get_pixmap(matrix=matrix, alpha=False)
        pix.save(image_path)
        rendered.append(
            {
                "page": page_no,
                "image": str(image_path),
                "width_px": pix.width,
                "height_px": pix.height,
            }
        )
    return rendered


def build_llm_ir_template(pdf_path: Path, pages: str) -> dict:
    return {
        "schema": SCHEMA,
        "extraction_mode": "llm_only",
        "source_pdf": str(pdf_path),
        "requested_pages": pages,
        "blocks": [
            {
                "page": 0,
                "concept_no": "",
                "concept_title": "",
                "practice_no": "",
                "prompt": "",
                "layout_type": "vertical_list",
                "source_lines": [],
                "items": [
                    {
                        "number": 1,
                        "raw_text": "(1) a^[ ]×[ ]",
                        "blank_count": 2,
                        "row_index": 0,
                        "column_index": 0,
                        "source_lines": [],
                        "display_lines": [],
                        "display_segments": [
                            {
                                "kind": "math",
                                "text": "a^[ ]×[ ]",
                                "line_index": 0,
                                "gap_after_in": None,
                            },
                            {
                                "kind": "marker",
                                "shape": "right_arrow",
                                "text": "",
                                "line_index": 0,
                                "gap_after_in": 0.032,
                            }
                        ],
                    }
                ],
            }
        ],
        "excluded_pages": {
            "0": "reason this requested page has no target concept-practice block"
        },
    }


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(description="Export an LLM-only IR extraction packet for concept-practice PDF pages.")
    parser.add_argument("--pdf", required=True, type=Path)
    parser.add_argument("--pages", required=True, help="Page range, e.g. 28-59 or 28~59")
    parser.add_argument("--output-dir", required=True, type=Path)
    parser.add_argument("--dpi", type=int, default=144)
    args = parser.parse_args(argv)

    if not args.pdf.exists():
        raise SystemExit(f"Input PDF not found: {args.pdf}")

    start_page, end_page = parse_page_range(args.pages)
    pages = list(range(start_page, end_page + 1))
    args.output_dir.mkdir(parents=True, exist_ok=True)

    template = build_llm_ir_template(args.pdf, args.pages)
    template_path = args.output_dir / "llm-ir-template.json"
    template_path.write_text(json.dumps(template, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    report = {
        "purpose": "LLM-only extraction packet. The LLM must inspect the rendered page images and write llm-ir.json. Scripts do not create PracticeBlock IR from the PDF in this mode.",
        "source_pdf": str(args.pdf),
        "requested_pages": args.pages,
        "page_images": render_page_images(args.pdf, pages, args.output_dir, args.dpi),
        "llm_ir_template": str(template_path),
        "llm_ir_schema": SCHEMA,
        "llm_extraction_contract": {
            "extract_ir_from_page_images_only": True,
            "include_only_concept_practice_blocks": True,
            "mark_non_target_requested_pages_in_excluded_pages": True,
            "represent_each_visible_square_blank_as_one_visible_token": "[ ]",
            "represent_exponent_square_blank_as": "^[ ]",
            "blank_count_must_equal_visible_blank_token_count": True,
            "do_not_invent_square_blanks": True,
            "do_not_show_source_answers": True,
            "return_llm_ir_json_only": True,
        },
    }

    output_path = args.output_dir / "llm-extraction-packet.json"
    output_path.write_text(json.dumps(report, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    print(json.dumps({"packet": str(output_path), "pages": pages, "template": str(template_path)}, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
