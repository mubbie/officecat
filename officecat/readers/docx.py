"""Word (.docx) to markdown reader."""

from __future__ import annotations

import sys
from pathlib import Path


def _classify_style(style_name: str | None) -> str | None:
    """Map a docx paragraph style name to a markdown heading prefix or None."""
    if style_name is None:
        return None
    name = style_name.lower()
    if "heading 1" in name or name == "title":
        return "# "
    if "heading 2" in name:
        return "## "
    if "heading 3" in name:
        return "### "
    if "heading 4" in name or "heading 5" in name or "heading 6" in name:
        return "#### "
    if "list" in name:
        return "- "
    return None


def _escape_pipe(text: str) -> str:
    return text.replace("|", "\\|")


def _table_to_markdown(table) -> str:
    """Convert a docx Table object to a markdown pipe table."""
    rows: list[list[str]] = []
    for row in table.rows:
        cells = row.cells
        row_text: list[str] = []
        prev_tc = None
        for cell in cells:
            if prev_tc is not None and cell._tc is prev_tc:
                continue
            row_text.append(_escape_pipe(cell.text.strip()))
            prev_tc = cell._tc
        rows.append(row_text)

    if not rows:
        return ""

    col_count = max(len(r) for r in rows)
    headers = rows[0] + [""] * (col_count - len(rows[0]))

    lines: list[str] = []
    lines.append("| " + " | ".join(headers[:col_count]) + " |")
    lines.append("| " + " | ".join(["---"] * col_count) + " |")

    for row in rows[1:]:
        padded = row + [""] * (col_count - len(row))
        lines.append("| " + " | ".join(padded[:col_count]) + " |")

    return "\n".join(lines)


def to_markdown(path: Path, *, head: int | None = None, **_kwargs) -> str:
    """Convert a docx file to markdown, preserving paragraph/table order."""
    from docx import Document
    from docx.opc.exceptions import PackageNotFoundError
    from docx.table import Table as DocxTable
    from docx.text.paragraph import Paragraph

    try:
        doc = Document(str(path))
    except PackageNotFoundError:
        from officecat.readers._errors import report_open_error

        report_open_error(path)
        raise SystemExit(3)
    except Exception as e:
        msg = str(e).lower()
        if "password" in msg or "encrypt" in msg:
            print(
                f"Error: '{path.name}' is password-protected. "
                f"officecat cannot open encrypted files.",
                file=sys.stderr,
            )
            raise SystemExit(3)
        raise

    blocks: list[str] = []
    block_count = 0

    # Walk body children to preserve interleaved order
    for child in doc.element.body:
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p":
            para = Paragraph(child, doc)
            text = para.text.strip()
            if not text:
                continue

            prefix = _classify_style(para.style.name if para.style else None)
            if prefix:
                blocks.append(f"{prefix}{text}")
            else:
                blocks.append(text)

            block_count += 1
            if head is not None and block_count >= head:
                break

        elif tag == "tbl":
            table = DocxTable(child, doc)
            md_table = _table_to_markdown(table)
            if md_table:
                blocks.append(md_table)
                block_count += 1
                if head is not None and block_count >= head:
                    break

    return "\n\n".join(blocks)
