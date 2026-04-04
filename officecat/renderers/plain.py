"""Plain text renderer (pipe-friendly)."""

from __future__ import annotations

from officecat.models import (
    CsvData,
    DocxDocument,
    DocxParagraph,
    Presentation,
    TableData,
    Workbook,
)

HEADING_PREFIX = {"heading1": "# ", "heading2": "## ", "heading3": "### ", "heading4": "#### "}


def render_docx(doc: DocxDocument) -> None:
    if not doc.blocks:
        print("Document is empty.")
        return

    for block in doc.blocks:
        if isinstance(block, DocxParagraph):
            prefix = HEADING_PREFIX.get(block.style, "")
            if block.style == "list_item":
                print(f"  - {block.text}")
            else:
                print(f"{prefix}{block.text}")
        elif isinstance(block, TableData):
            print()
            if block.headers:
                print("\t".join(block.headers))
            for row in block.rows:
                print("\t".join(row))


def render_pptx(pres: Presentation) -> None:
    if not pres.slides:
        print("Presentation is empty.")
        return

    for slide in pres.slides:
        title = slide.title or "Untitled"
        print(f"--- Slide {slide.number}: {title} ---")
        for line in slide.body:
            print(line)
        for img in slide.images:
            print(img)
        if slide.notes:
            print(f"[Notes] {slide.notes}")
        print()


def _render_table(headers: list[str], rows: list[list[str]], total_rows: int, title: str | None = None) -> None:
    if title:
        print(title)
    if headers:
        print("\t".join(headers))
    for row in rows:
        padded = row + [""] * (len(headers) - len(row))
        print("\t".join(padded[:len(headers)]))
    if total_rows > len(rows):
        print(f"[Showing {len(rows)} of {total_rows} rows]")


def render_xlsx(wb: Workbook) -> None:
    if not wb.sheets:
        print("Workbook is empty.")
        return

    for i, sheet in enumerate(wb.sheets):
        if i > 0:
            print()
        _render_table(sheet.headers, sheet.rows, sheet.total_rows, title=sheet.name)


def render_csv(data: CsvData) -> None:
    if not data.rows and not data.headers:
        print("File is empty.")
        return

    _render_table(data.headers, data.rows, data.total_rows)
