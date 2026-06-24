from __future__ import annotations

from dataclasses import dataclass, field
import re


KOREAN_LABEL_TEXTS = (
    "분모의 소인수:",
    "순환마디:",
    "간단히 표현:",
)

MATH_TOKEN_RE = re.compile(r"\d|[=+\-×÷/()[\]\^⁰¹²³⁴⁵⁶⁷⁸⁹]")
BLANK_TOKEN_RE = re.compile(r"\^\[\s*\]|\[\s*\]|__|□")


@dataclass
class SourceBox:
    """PDF-space bounding box for a detected source object."""

    x0: float
    y0: float
    x1: float
    y1: float

    @property
    def center_x(self) -> float:
        return (self.x0 + self.x1) / 2

    @property
    def center_y(self) -> float:
        return (self.y0 + self.y1) / 2

    def to_dict(self) -> dict:
        return {
            "x0": round(self.x0, 3),
            "y0": round(self.y0, 3),
            "x1": round(self.x1, 3),
            "y1": round(self.y1, 3),
        }


@dataclass
class DisplaySegment:
    """Typed display fragment inside a rendered item row."""

    kind: str
    text: str
    line_index: int = 0
    gap_after_in: float | None = None
    shape: str | None = None

    def to_dict(self) -> dict:
        data = {
            "kind": self.kind,
            "text": self.text,
            "line_index": self.line_index,
        }
        if self.gap_after_in is not None:
            data["gap_after_in"] = round(self.gap_after_in, 3)
        if self.shape:
            data["shape"] = self.shape
        return data


@dataclass
class PracticeItem:
    """Normalized subproblem inventory entry."""

    raw_text: str
    number: int | None = None
    expression_text: str = ""
    blank_count: int = 0
    has_square_blanks: bool = False
    has_parenthesis_blank: bool = False
    fraction_count: int = 0
    has_exponent: bool = False
    row_index: int = 0
    column_index: int = 0
    source_box: SourceBox | None = None
    source_blanks: list[SourceBox] = field(default_factory=list)
    source_lines: list[str] = field(default_factory=list)
    display_lines: list[str] = field(default_factory=list)
    display_segments: list[DisplaySegment] = field(default_factory=list)

    @classmethod
    def from_text(
        cls,
        text: str,
        *,
        row_index: int = 0,
        column_index: int = 0,
        source_box: SourceBox | None = None,
        source_blanks: list[SourceBox] | None = None,
        source_lines: list[str] | None = None,
        display_lines: list[str] | None = None,
        display_segments: list[DisplaySegment] | None = None,
    ) -> "PracticeItem":
        number_match = re.match(r"^\((\d{1,2})\)\s*", text)
        number = int(number_match.group(1)) if number_match else None
        expression = re.sub(r"^\(\d{1,2}\)\s*", "", text).strip()
        source_lines = source_lines or [text]
        display_lines = display_lines or display_lines_from_source_lines(source_lines, expression)
        display_segments = display_segments or display_segments_from_display_lines(display_lines)
        source_blanks = source_blanks if source_blanks is not None else []
        text_square_blank_count = len(BLANK_TOKEN_RE.findall(text))
        square_blank_count = len(source_blanks) if source_blanks else text_square_blank_count
        return cls(
            raw_text=text,
            number=number,
            expression_text=expression,
            blank_count=square_blank_count,
            has_square_blanks=square_blank_count > 0,
            has_parenthesis_blank=bool(re.search(r"\(\s{2,}\)", text)),
            fraction_count=len(
                re.findall(
                    r"(?:-?\([^)]*\)|-?[\w가-힣⁰¹²³⁴⁵⁶⁷⁸⁹\]\)]+)\s*/\s*(?:-?\([^)]*\)|-?[\w가-힣⁰¹²³⁴⁵⁶⁷⁸⁹\[\(]+)",
                    text,
                )
            ),
            has_exponent=bool(re.search(r"[⁰¹²³⁴⁵⁶⁷⁸⁹]|\^\d+", text)),
            row_index=row_index,
            column_index=column_index,
            source_box=source_box,
            source_blanks=source_blanks,
            source_lines=source_lines,
            display_lines=display_lines,
            display_segments=display_segments,
        )

    def to_dict(self) -> dict:
        return {
            "number": self.number,
            "raw_text": self.raw_text,
            "expression_text": self.expression_text,
            "blank_count": self.blank_count,
            "has_square_blanks": self.has_square_blanks,
            "has_parenthesis_blank": self.has_parenthesis_blank,
            "fraction_count": self.fraction_count,
            "has_exponent": self.has_exponent,
            "row_index": self.row_index,
            "column_index": self.column_index,
            "source_blanks": [blank.to_dict() for blank in self.source_blanks],
            "source_lines": self.source_lines,
            "display_lines": self.display_lines,
            "display_segments": [segment.to_dict() for segment in self.display_segments],
            "source_box": self.source_box.to_dict() if self.source_box else None,
        }


@dataclass
class PracticeBlock:
    """Normalized concept-practice block extracted from one source PDF page."""

    page: int
    concept_no: str
    concept_title: str
    practice_no: str
    prompt: str
    items: list[str] = field(default_factory=list)
    item_models: list[PracticeItem] = field(default_factory=list)
    layout_type: str = "vertical_list"
    source_lines: list[str] = field(default_factory=list)

    def __post_init__(self) -> None:
        if not self.item_models:
            self.item_models = [
                PracticeItem.from_text(item, row_index=idx)
                for idx, item in enumerate(self.items)
            ]
        elif not self.items:
            self.items = [item.raw_text for item in self.item_models]
        if not self.layout_type:
            self.layout_type = infer_layout_type(self.item_models)

    def to_inventory(self) -> dict:
        return {
            "page": self.page,
            "concept_no": self.concept_no,
            "concept_title": self.concept_title,
            "practice_no": self.practice_no,
            "prompt": self.prompt,
            "layout_type": self.layout_type,
            "items": [item.to_dict() for item in self.item_models],
        }


def infer_layout_type(items: list[PracticeItem]) -> str:
    if not items:
        return "unknown"
    if any(item.column_index > 0 for item in items):
        return "two_column_grid"
    if len(items) <= 2 and any(item.fraction_count >= 2 for item in items):
        return "worked_stack"
    return "vertical_list"


def display_lines_from_source_lines(source_lines: list[str], fallback_expression: str) -> list[str]:
    cleaned: list[str] = []
    for idx, line in enumerate(source_lines):
        value = line.strip()
        if idx == 0:
            value = re.sub(r"^\(\d{1,2}\)\s*", "", value).strip()
        if value:
            cleaned.append(value)
    return cleaned or [fallback_expression]


def is_korean_label_text(text: str) -> bool:
    return bool(re.search(r"[가-힣]", text)) and MATH_TOKEN_RE.search(text) is None


def split_inline_korean_label(text: str) -> tuple[str, str] | None:
    for label in KOREAN_LABEL_TEXTS:
        index = text.rfind(label)
        if index <= 0:
            continue
        math_text = text[:index].rstrip()
        if MATH_TOKEN_RE.search(math_text):
            return math_text, label
    return None


def display_segments_from_display_lines(display_lines: list[str]) -> list[DisplaySegment]:
    segments: list[DisplaySegment] = []
    for line_index, line in enumerate(display_lines):
        text = line.strip()
        if not text:
            continue
        inline_split = split_inline_korean_label(text)
        if inline_split:
            math_text, label_text = inline_split
            segments.append(DisplaySegment(kind="math", text=math_text, line_index=line_index))
            segments.append(DisplaySegment(kind="korean_label", text=label_text, line_index=line_index))
        elif is_korean_label_text(text):
            segments.append(DisplaySegment(kind="korean_label", text=text, line_index=line_index))
        else:
            segments.append(DisplaySegment(kind="math", text=text, line_index=line_index))
    return segments
