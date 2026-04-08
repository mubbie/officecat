"""CSV/TSV to markdown reader."""

from __future__ import annotations

import csv
from pathlib import Path

DEFAULT_ROW_CAP = 500


def _escape_pipe(text: str) -> str:
    """Escape pipe characters for markdown tables."""
    return text.replace("|", "\\|")


def _col_letter(index: int) -> str:
    """Convert 0-based index to Excel-style letter (A, B, ... Z, AA, ...)."""
    result = ""
    i = index
    while True:
        result = chr(65 + i % 26) + result
        i = i // 26 - 1
        if i < 0:
            break
    return result


def to_markdown(
    path: Path,
    *,
    head: int | None = None,
    headers: int = 1,
    show_all: bool = False,
    **_kwargs,
) -> str:
    """Convert a CSV/TSV file to a markdown table string."""
    is_tsv = path.suffix.lower() == ".tsv"

    with open(path, newline="", encoding="utf-8-sig") as f:
        if is_tsv:
            delimiter = "\t"
        else:
            first_line = f.readline()
            f.seek(0)
            # Only sniff if comma-split yields a single column
            if len(first_line.split(",")) <= 1:
                try:
                    dialect = csv.Sniffer().sniff(first_line)
                    delimiter = dialect.delimiter
                except csv.Error:
                    delimiter = ","
            else:
                delimiter = ","

        reader = csv.reader(f, delimiter=delimiter)

        header_row: list[str] = []
        rows: list[list[str]] = []
        total_rows = 0
        row_cap = head if head is not None else (None if show_all else DEFAULT_ROW_CAP)

        for i, row in enumerate(reader, start=1):
            if headers > 0 and i == headers:
                header_row = [_escape_pipe(c) for c in row]
                continue

            total_rows += 1
            if row_cap is not None and len(rows) >= row_cap:
                continue  # keep counting total_rows
            rows.append([_escape_pipe(c) for c in row])

    if not header_row and rows:
        col_count = max(len(r) for r in rows)
        header_row = [_col_letter(i) for i in range(col_count)]

    if not header_row and not rows:
        return ""

    return _build_table(header_row, rows, total_rows)


def _build_table(
    headers: list[str], rows: list[list[str]], total_rows: int
) -> str:
    """Build a markdown pipe table from headers and rows."""
    col_count = len(headers)
    lines: list[str] = []

    lines.append("| " + " | ".join(headers) + " |")
    lines.append("| " + " | ".join(["---"] * col_count) + " |")

    for row in rows:
        padded = row + [""] * (col_count - len(row))
        lines.append("| " + " | ".join(padded[:col_count]) + " |")

    if total_rows > len(rows):
        lines.append("")
        lines.append(f"*Showing {len(rows)} of {total_rows:,} rows.*")

    return "\n".join(lines)
