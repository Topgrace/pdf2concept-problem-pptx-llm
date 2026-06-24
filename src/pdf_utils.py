from __future__ import annotations

import hashlib
import re
from pathlib import Path


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def parse_page_range(value: str) -> tuple[int, int]:
    match = re.fullmatch(r"\s*(\d+)\s*[-~]\s*(\d+)\s*", value)
    if not match:
        raise SystemExit("--pages must look like 10-19 or 10~19")
    start, end = int(match.group(1)), int(match.group(2))
    if start <= 0 or end < start:
        raise SystemExit(f"Invalid page range: {value}")
    return start, end
