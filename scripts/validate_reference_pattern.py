#!/usr/bin/env python3
"""Validate reference-style animation and blank-box patterns in a PPTX deck."""

from __future__ import annotations

import argparse
import json
import re
import sys
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET


NS = {
    "p": "http://schemas.openxmlformats.org/presentationml/2006/main",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

EMU_PER_INCH = 914400


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def slide_number(path: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", path)
    return int(match.group(1)) if match else 0


def emu_to_in(value: str | None) -> float | None:
    if value is None:
        return None
    return int(value) / EMU_PER_INCH


def direct_cnvpr(element: ET.Element, kind: str) -> ET.Element | None:
    if kind == "sp":
        return element.find("./p:nvSpPr/p:cNvPr", NS)
    if kind == "grpSp":
        return element.find("./p:nvGrpSpPr/p:cNvPr", NS)
    return None


def direct_xfrm(element: ET.Element, kind: str) -> ET.Element | None:
    if kind == "sp":
        return element.find("./p:spPr/a:xfrm", NS)
    if kind == "grpSp":
        return element.find("./p:grpSpPr/a:xfrm", NS)
    return None


def object_bounds(element: ET.Element, kind: str) -> tuple[float, float, float, float] | None:
    xfrm = direct_xfrm(element, kind)
    if xfrm is None:
        return None
    off = xfrm.find("./a:off", NS)
    ext = xfrm.find("./a:ext", NS)
    if off is None or ext is None:
        return None
    x = emu_to_in(off.get("x"))
    y = emu_to_in(off.get("y"))
    width = emu_to_in(ext.get("cx"))
    height = emu_to_in(ext.get("cy"))
    if x is None or y is None or width is None or height is None:
        return None
    return x, y, width, height


def intersects_slide(
    element: ET.Element,
    kind: str,
    slide_width: float,
    slide_height: float,
) -> bool:
    bounds = object_bounds(element, kind)
    if bounds is None:
        return True
    x, y, width, height = bounds
    return x + width > 0 and y + height > 0 and x < slide_width and y < slide_height


def within_slide_canvas(
    element: ET.Element,
    kind: str,
    slide_width: float,
    slide_height: float,
    tolerance: float = 0.01,
) -> bool:
    bounds = object_bounds(element, kind)
    if bounds is None:
        return True
    x, y, width, height = bounds
    return (
        x >= -tolerance
        and y >= -tolerance
        and x + width <= slide_width + tolerance
        and y + height <= slide_height + tolerance
    )


def shape_text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.findall(".//a:t", NS))


def shape_geometry(element: ET.Element) -> str | None:
    geom = element.find(".//a:prstGeom", NS)
    return geom.get("prst") if geom is not None else None


def shape_extents(element: ET.Element) -> tuple[float | None, float | None]:
    ext = element.find(".//a:xfrm/a:ext", NS)
    if ext is None:
        return None, None
    return emu_to_in(ext.get("cx")), emu_to_in(ext.get("cy"))


def is_reference_blank_shape(element: ET.Element) -> bool:
    if shape_geometry(element) != "roundRect":
        return False
    if clean(shape_text(element)):
        return False
    width, height = shape_extents(element)
    if width is None or height is None:
        return False
    return 0.35 <= width <= 1.15 and 0.25 <= height <= 0.75


def shape_fill(element: ET.Element) -> str | None:
    srgb = element.find("./p:spPr/a:solidFill/a:srgbClr", NS)
    if srgb is not None:
        return srgb.get("val")
    scheme = element.find("./p:spPr/a:solidFill/a:schemeClr", NS)
    if scheme is not None:
        return scheme.get("val")
    return None


def is_slide_local_background_or_panel(element: ET.Element, slide_width: float, slide_height: float) -> bool:
    if clean(shape_text(element)):
        return False
    bounds = object_bounds(element, "sp")
    if bounds is None:
        return False
    x, y, width, height = bounds
    fill = shape_fill(element)
    geometry = shape_geometry(element)
    if (
        geometry == "rect"
        and fill in {"CCDEED", "CCDEF0", "bg1"}
        and abs(x) <= 0.03
        and abs(y) <= 0.03
        and width >= slide_width - 0.05
        and height >= slide_height - 0.05
    ):
        return True
    return (
        geometry == "roundRect"
        and fill in {"FFFFFF", "bg1"}
        and x <= 0.40
        and 0.20 <= y <= 0.75
        and width >= 8.50
        and height >= 6.00
    )


def inspect_slide(root: ET.Element, number: int, slide_width: float, slide_height: float) -> dict:
    off_slide_objects = []
    outside_canvas_objects = []
    on_slide_elements: list[tuple[str, ET.Element]] = []
    for kind, path in (("sp", ".//p:sp"), ("grpSp", ".//p:grpSp")):
        for element in root.findall(path, NS):
            c_nv_pr = direct_cnvpr(element, kind)
            object_id = c_nv_pr.get("id") if c_nv_pr is not None else None
            name = c_nv_pr.get("name") if c_nv_pr is not None else None
            on_slide = intersects_slide(element, kind, slide_width, slide_height)
            within_canvas = within_slide_canvas(element, kind, slide_width, slide_height)
            bounds = object_bounds(element, kind)
            if on_slide:
                on_slide_elements.append((kind, element))
            else:
                off_slide_objects.append(
                    {
                        "id": object_id,
                        "kind": kind,
                        "name": name,
                        "text": clean(shape_text(element))[:80],
                    }
                )
            if not within_canvas:
                outside_canvas_objects.append(
                    {
                        "id": object_id,
                        "kind": kind,
                        "name": name,
                        "text": clean(shape_text(element))[:80],
                        "bounds": [round(value, 3) for value in bounds] if bounds else None,
                    }
                )

    text_nodes = [
        node.text or ""
        for kind, element in on_slide_elements
        for node in element.findall(".//a:t", NS)
    ]
    combined_text = clean(" ".join(text_nodes))
    is_title_slide = "개념핵심" in combined_text

    objects_by_id: dict[str, dict[str, str | None]] = {}
    for kind, path in (("sp", ".//p:sp"), ("grpSp", ".//p:grpSp")):
        for element in root.findall(path, NS):
            c_nv_pr = direct_cnvpr(element, kind)
            if c_nv_pr is not None and c_nv_pr.get("id"):
                on_slide = intersects_slide(element, kind, slide_width, slide_height)
                objects_by_id[c_nv_pr.get("id", "")] = {
                    "kind": kind,
                    "name": c_nv_pr.get("name"),
                    "text": clean(shape_text(element)),
                    "on_slide": on_slide,
                    "hidden": c_nv_pr.get("hidden"),
                }

    set_targets = []
    off_slide_set_targets = []
    missing_targets = []
    non_group_targets = []
    hidden_set_targets = []
    for node in root.findall(".//p:set", NS):
        target = node.find(".//p:spTgt", NS)
        target_id = target.get("spid") if target is not None else None
        target_info = objects_by_id.get(target_id or "")
        if target_id and target_info and not target_info["on_slide"]:
            off_slide_set_targets.append(
                {
                    "target_id": target_id,
                    "kind": target_info["kind"],
                    "text": target_info["text"][:80],
                }
            )
            continue
        set_targets.append(
            {
                "target_id": target_id,
                "kind": target_info["kind"] if target_info else None,
                "text": target_info["text"][:80] if target_info and target_info["text"] else "",
                "hidden": target_info["hidden"] if target_info else None,
            }
        )
        if target_id and target_id not in objects_by_id:
            missing_targets.append(target_id)
        elif target_id and target_info and target_info["kind"] != "grpSp":
            non_group_targets.append(
                {
                    "target_id": target_id,
                    "kind": target_info["kind"],
                    "text": target_info["text"][:80],
                }
            )
        elif target_id and target_info and target_info["hidden"]:
            hidden_set_targets.append(
                {
                    "target_id": target_id,
                    "kind": target_info["kind"],
                    "text": target_info["text"][:80],
                    "hidden": target_info["hidden"],
                }
            )

    blank_shapes = [
        {
            "id": direct_cnvpr(element, "sp").get("id") if direct_cnvpr(element, "sp") is not None else None,
            "name": direct_cnvpr(element, "sp").get("name") if direct_cnvpr(element, "sp") is not None else None,
        }
        for element in root.findall(".//p:sp", NS)
        if intersects_slide(element, "sp", slide_width, slide_height) and is_reference_blank_shape(element)
    ]

    local_background_overlays = []
    for element in root.findall(".//p:sp", NS):
        if not is_slide_local_background_or_panel(element, slide_width, slide_height):
            continue
        c_nv_pr = direct_cnvpr(element, "sp")
        bounds = object_bounds(element, "sp")
        local_background_overlays.append(
            {
                "id": c_nv_pr.get("id") if c_nv_pr is not None else None,
                "name": c_nv_pr.get("name") if c_nv_pr is not None else None,
                "geometry": shape_geometry(element),
                "fill": shape_fill(element),
                "bounds": [round(value, 3) for value in bounds] if bounds else None,
            }
        )

    literal_square_count = sum(text.count("□") for text in text_nodes)
    visible_page_labels = [
        text
        for text in [clean(value) for value in text_nodes]
        if re.fullmatch(r"\d+\s*쪽", text)
    ]
    return {
        "slide": number,
        "is_title_slide": is_title_slide,
        "literal_square_count": literal_square_count,
        "visible_page_labels": visible_page_labels,
        "set_count": len(set_targets),
        "set_targets": set_targets,
        "off_slide_set_target_count": len(off_slide_set_targets),
        "off_slide_set_targets": off_slide_set_targets,
        "missing_timing_targets": missing_targets,
        "non_group_set_targets": non_group_targets,
        "hidden_set_targets": hidden_set_targets,
        "blank_shape_count": 0 if is_title_slide else len(blank_shapes),
        "blank_shapes": [] if is_title_slide else blank_shapes[:20],
        "off_slide_object_count": len(off_slide_objects),
        "off_slide_objects": off_slide_objects[:20],
        "outside_canvas_object_count": len(outside_canvas_objects),
        "outside_canvas_objects": outside_canvas_objects[:20],
        "local_background_overlay_count": len(local_background_overlays),
        "local_background_overlays": local_background_overlays[:20],
        "sample_text": combined_text[:140],
    }


def inspect_deck(path: Path) -> dict:
    with zipfile.ZipFile(path) as deck:
        bad_entry = deck.testzip()
        if bad_entry is not None:
            raise ValueError(f"Invalid PPTX zip entry: {bad_entry}")
        slide_paths = sorted(
            [
                name
                for name in deck.namelist()
                if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
            ],
            key=slide_number,
        )
        presentation = ET.fromstring(deck.read("ppt/presentation.xml"))
        slide_size = presentation.find("./p:sldSz", NS)
        slide_width = emu_to_in(slide_size.get("cx")) if slide_size is not None else 13.333
        slide_height = emu_to_in(slide_size.get("cy")) if slide_size is not None else 7.5
        slides = [
            inspect_slide(
                ET.fromstring(deck.read(name)),
                slide_number(name),
                slide_width or 13.333,
                slide_height or 7.5,
            )
            for name in slide_paths
        ]
    return {
        "path": str(path),
        "slide_count": len(slides),
        "slide_size_inches": {
            "width": round(slide_width or 13.333, 3),
            "height": round(slide_height or 7.5, 3),
        },
        "slides": slides,
    }


def validate(report: dict, args: argparse.Namespace) -> dict:
    failures = []
    total_sets = sum(slide["set_count"] for slide in report["slides"])
    total_off_slide_sets = sum(slide["off_slide_set_target_count"] for slide in report["slides"])
    total_problem_blank_shapes = sum(slide["blank_shape_count"] for slide in report["slides"])
    total_literal_squares = sum(slide["literal_square_count"] for slide in report["slides"])
    total_off_slide_objects = sum(slide["off_slide_object_count"] for slide in report["slides"])
    total_outside_canvas_objects = sum(slide["outside_canvas_object_count"] for slide in report["slides"])
    total_visible_page_labels = sum(len(slide["visible_page_labels"]) for slide in report["slides"])
    total_local_background_overlays = sum(slide["local_background_overlay_count"] for slide in report["slides"])

    if total_literal_squares:
        failures.append(
            {
                "rule": "no_literal_square_blanks",
                "message": "Typed □ characters were found. Use reference-style blank shapes instead.",
                "slides": [
                    {
                        "slide": slide["slide"],
                        "literal_square_count": slide["literal_square_count"],
                        "sample_text": slide["sample_text"],
                    }
                    for slide in report["slides"]
                    if slide["literal_square_count"]
                ],
            }
        )

    missing_target_slides = [
        {
            "slide": slide["slide"],
            "missing_timing_targets": slide["missing_timing_targets"],
        }
        for slide in report["slides"]
        if slide["missing_timing_targets"]
    ]
    if missing_target_slides:
        failures.append(
            {
                "rule": "timing_targets_exist",
                "message": "Animation timing targets must resolve to existing slide object ids.",
                "slides": missing_target_slides,
            }
        )

    if args.require_animations and total_sets == 0:
        failures.append(
            {
                "rule": "require_animations",
                "message": "No p:set reveal animations were found.",
            }
        )

    if args.require_group_animation:
        non_group_target_slides = [
            {
                "slide": slide["slide"],
                "non_group_set_targets": slide["non_group_set_targets"],
            }
            for slide in report["slides"]
            if slide["non_group_set_targets"]
        ]
        if non_group_target_slides:
            failures.append(
                {
                    "rule": "set_targets_are_groups",
                    "message": "Reference-style p:set reveals must target grouped objects (p:grpSp), not loose text boxes.",
                    "slides": non_group_target_slides,
                }
            )

    hidden_target_slides = [
        {
            "slide": slide["slide"],
            "hidden_set_targets": slide["hidden_set_targets"],
        }
        for slide in report["slides"]
        if slide["hidden_set_targets"]
    ]
    if hidden_target_slides:
        failures.append(
            {
                "rule": "set_targets_not_hidden",
                "message": "Animated p:set targets must not have cNvPr hidden=1; PowerPoint may skip hidden targets during slideshow clicks.",
                "slides": hidden_target_slides,
            }
        )

    if args.require_blank_shapes and total_problem_blank_shapes == 0:
        failures.append(
            {
                "rule": "require_blank_shapes",
                "message": "No reference-style blank rounded-rectangle shapes were found on non-title slides.",
            }
        )

    if args.forbid_visible_page_labels and total_visible_page_labels:
        failures.append(
            {
                "rule": "forbid_visible_page_labels",
                "message": "Visible source page labels such as '10쪽' are not allowed on the slide canvas.",
                "slides": [
                    {
                        "slide": slide["slide"],
                        "visible_page_labels": slide["visible_page_labels"],
                    }
                    for slide in report["slides"]
                    if slide["visible_page_labels"]
                ],
            }
        )

    if args.forbid_off_slide_objects and total_outside_canvas_objects:
        failures.append(
            {
                "rule": "forbid_off_slide_objects",
                "message": "Generated decks must not contain unrelated objects fully or partially outside the slide canvas.",
                "slides": [
                    {
                        "slide": slide["slide"],
                        "outside_canvas_objects": slide["outside_canvas_objects"],
                    }
                    for slide in report["slides"]
                    if slide["outside_canvas_object_count"]
                ],
            }
        )

    if total_local_background_overlays:
        failures.append(
            {
                "rule": "forbid_slide_local_background_overlays",
                "message": "Use the bundled master/layout background and panel; do not create slide-local full-canvas backgrounds or large content panels.",
                "slides": [
                    {
                        "slide": slide["slide"],
                        "local_background_overlays": slide["local_background_overlays"],
                    }
                    for slide in report["slides"]
                    if slide["local_background_overlay_count"]
                ],
            }
        )

    report["summary"] = {
        "total_set_animations": total_sets,
        "total_off_slide_set_animations_ignored": total_off_slide_sets,
        "total_problem_blank_shapes": total_problem_blank_shapes,
        "total_literal_square_characters": total_literal_squares,
        "total_off_slide_objects_ignored": total_off_slide_objects,
        "total_outside_canvas_objects": total_outside_canvas_objects,
        "total_visible_page_labels": total_visible_page_labels,
        "total_local_background_overlays": total_local_background_overlays,
    }
    report["failures"] = failures
    report["passes"] = not failures
    return report


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Fail when PPTX animation targets or blank boxes drift from the measured reference pattern."
    )
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--require-animations", action="store_true")
    parser.add_argument("--require-group-animation", action="store_true")
    parser.add_argument("--require-blank-shapes", action="store_true")
    parser.add_argument("--forbid-visible-page-labels", action="store_true")
    parser.add_argument("--forbid-off-slide-objects", action="store_true")
    parser.add_argument("--report", type=Path)
    args = parser.parse_args(argv)

    report = validate(inspect_deck(args.pptx), args)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0 if report["passes"] else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
