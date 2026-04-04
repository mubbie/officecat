"""Data models for all supported file formats."""

from __future__ import annotations

from dataclasses import dataclass, field


# ── Shared ──


@dataclass
class TableData:
    """A table extracted from any format."""

    headers: list[str]
    rows: list[list[str]]


# ── Word ──


@dataclass
class DocxParagraph:
    text: str
    style: str  # "heading1", "heading2", "heading3", "heading4", "body", "list_item"
    level: int  # nesting depth, 0 = top level


@dataclass
class DocxDocument:
    source: str
    blocks: list  # list of DocxParagraph | TableData, interleaved in document order


# ── PowerPoint ──


@dataclass
class Slide:
    number: int
    title: str | None
    body: list[str]
    notes: str | None
    images: list[str] = field(default_factory=list)


@dataclass
class Presentation:
    source: str
    slides: list[Slide]
    total_slides: int


# ── Excel ──


@dataclass
class Sheet:
    name: str
    headers: list[str]
    rows: list[list[str]]
    total_rows: int


@dataclass
class Workbook:
    source: str
    sheets: list[Sheet]
    sheet_names: list[str]


# ── CSV ──


@dataclass
class CsvData:
    source: str
    headers: list[str]
    rows: list[list[str]]
    total_rows: int
