"""Word (.docx) reader."""

from __future__ import annotations

import sys
from pathlib import Path

from officecat.models import DocxDocument, DocxParagraph, TableData


def _classify_style(style_name: str | None) -> str:
    if style_name is None:
        return "body"
    name = style_name.lower()
    if "heading 1" in name or name == "title":
        return "heading1"
    if "heading 2" in name:
        return "heading2"
    if "heading 3" in name:
        return "heading3"
    if "heading 4" in name or "heading 5" in name or "heading 6" in name:
        return "heading4"
    if "list" in name:
        return "list_item"
    return "body"


def read_docx(path: Path, *, head: int | None = None) -> DocxDocument:
    from docx import Document
    from docx.opc.exceptions import PackageNotFoundError

    try:
        doc = Document(str(path))
    except PackageNotFoundError:
        print(
            f"Error: '{path.name}' appears to be corrupt or is not a valid docx file.",
            file=sys.stderr,
        )
        sys.exit(3)
    except Exception as e:
        msg = str(e).lower()
        if "password" in msg or "encrypt" in msg:
            print(
                f"Error: '{path.name}' is password-protected. "
                f"officecat cannot open encrypted files.",
                file=sys.stderr,
            )
            sys.exit(3)
        raise

    blocks: list = []

    # Paragraphs first (known limitation: tables come after, not interleaved)
    count = 0
    for para in doc.paragraphs:
        text = para.text.strip()
        if not text:
            continue

        style = _classify_style(para.style.name if para.style else None)
        level = 0
        if para.paragraph_format and para.paragraph_format.level is not None:
            level = para.paragraph_format.level

        blocks.append(DocxParagraph(text=text, style=style, level=level))
        count += 1
        if head is not None and count >= head:
            break

    # Tables after paragraphs
    for table in doc.tables:
        rows: list[list[str]] = []
        for row in table.rows:
            cells = row.cells
            row_text: list[str] = []
            prev_tc = None
            for cell in cells:
                # Deduplicate merged cells by checking XML element identity
                if prev_tc is not None and cell._tc is prev_tc:
                    continue
                row_text.append(cell.text.strip())
                prev_tc = cell._tc
            rows.append(row_text)

        if rows:
            blocks.append(TableData(headers=rows[0], rows=rows[1:]))

    return DocxDocument(source=path.name, blocks=blocks)
