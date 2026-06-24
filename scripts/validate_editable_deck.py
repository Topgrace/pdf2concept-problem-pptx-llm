#!/usr/bin/env python3
"""Validate that a generated PPTX is editable, not a full-problem image deck."""

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


def clean(text: str) -> str:
    return re.sub(r"\s+", " ", text).strip()


def slide_number(path: str) -> int:
    match = re.search(r"slide(\d+)\.xml$", path)
    return int(match.group(1)) if match else 0


def inspect_deck(path: Path) -> dict:
    with zipfile.ZipFile(path) as deck:
        bad_entry = deck.testzip()
        if bad_entry is not None:
            raise ValueError(f"Invalid PPTX zip entry: {bad_entry}")
        slide_paths = sorted(
            [
                name
                for name in deck.namelist()
                if name.startswith("ppt/slides/slide") and name.endswith(".xml")
            ],
            key=slide_number,
        )
        slides = []
        for name in slide_paths:
            root = ET.fromstring(deck.read(name))
            texts = [
                clean(node.text or "")
                for node in root.findall(".//a:t", NS)
                if clean(node.text or "")
            ]
            pictures = root.findall(".//p:pic", NS)
            text_chars = sum(len(text) for text in texts)
            text_shapes = len([text for text in texts if text])
            picture_count = len(pictures)
            editable_enough = text_chars >= 12 or (text_shapes >= 2 and picture_count <= 2)
            image_only = picture_count > 0 and text_chars == 0
            slides.append(
                {
                    "slide": slide_number(name),
                    "text_chars": text_chars,
                    "text_shapes": text_shapes,
                    "picture_count": picture_count,
                    "image_only": image_only,
                    "editable_enough": editable_enough,
                    "sample_text": " | ".join(texts[:4]),
                }
            )
    failures = [
        slide
        for slide in slides
        if slide["image_only"] or not slide["editable_enough"]
    ]
    return {
        "path": str(path),
        "slide_count": len(slides),
        "slides": slides,
        "failures": failures,
        "passes": not failures,
    }


def main(argv: list[str]) -> int:
    parser = argparse.ArgumentParser(
        description="Fail when a PPTX appears to use full problem screenshots instead of editable text/shapes."
    )
    parser.add_argument("pptx", type=Path)
    parser.add_argument("--report", type=Path)
    args = parser.parse_args(argv)

    report = inspect_deck(args.pptx)
    text = json.dumps(report, ensure_ascii=False, indent=2)
    if args.report:
        args.report.parent.mkdir(parents=True, exist_ok=True)
        args.report.write_text(text + "\n", encoding="utf-8")
    print(text)
    return 0 if report["passes"] else 1


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
