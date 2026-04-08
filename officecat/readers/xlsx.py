"""Excel (.xlsx) to markdown reader using python-calamine."""

from __future__ import annotations

import sys
from pathlib import Path

DEFAULT_ROW_CAP = 500


def _escape_pipe(text: str) -> str:
    return text.replace("|", "\\|")


def _col_letter(index: int) -> str:
    result = ""
    i = index
    while True:
        result = chr(65 + i % 26) + result
        i = i // 26 - 1
        if i < 0:
            break
    return result


def _format_cell(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value == int(value):
        return str(int(value))
    if hasattr(value, "isoformat"):
        return value.isoformat()
    return str(value)


def to_markdown(
    path: Path,
    *,
    head: int | None = None,
    sheet: str | None = None,
    headers: int = 1,
    show_all: bool = False,
    **_kwargs,
) -> str:
    """Convert an xlsx file to markdown tables."""
    from python_calamine import CalamineWorkbook

    try:
        wb = CalamineWorkbook.from_path(str(path))
    except Exception as e:
        msg = str(e).lower()
        if "password" in msg or "encrypt" in msg:
            print(
                f"Error: '{path.name}' is password-protected. "
                f"officecat cannot open encrypted files.",
                file=sys.stderr,
            )
            raise SystemExit(3)
        if "zip" in msg or "invalid" in msg or "corrupt" in msg:
            print(
                f"Error: '{path.name}' appears to be corrupt or invalid.",
                file=sys.stderr,
            )
            raise SystemExit(3)
        raise

    all_sheet_names = wb.sheet_names

    if sheet is not None:
        # Try as 1-based index
        try:
            idx = int(sheet)
            if 1 <= idx <= len(all_sheet_names):
                sheets_to_read = [all_sheet_names[idx - 1]]
            else:
                print(
                    f"Error: Sheet index {idx} out of range. "
                    f"Available: {', '.join(all_sheet_names)}",
                    file=sys.stderr,
                )
                raise SystemExit(1)
        except ValueError:
            if sheet in all_sheet_names:
                sheets_to_read = [sheet]
            else:
                print(
                    f"Error: Sheet '{sheet}' not found. "
                    f"Available: {', '.join(all_sheet_names)}",
                    file=sys.stderr,
                )
                raise SystemExit(1)
    else:
        sheets_to_read = list(all_sheet_names)

    row_cap = head if head is not None else (None if show_all else DEFAULT_ROW_CAP)
    sections: list[str] = []

    for sheet_name in sheets_to_read:
        ws = wb.get_sheet_by_name(sheet_name)
        header_row: list[str] = []
        data_rows: list[list[str]] = []
        total_rows = 0

        for i, row in enumerate(ws.iter_rows(), start=1):
            if headers > 0 and i == headers:
                header_row = [_escape_pipe(_format_cell(c)) for c in row]
                continue

            total_rows += 1
            if row_cap is not None and len(data_rows) >= row_cap:
                continue
            data_rows.append([_escape_pipe(_format_cell(c)) for c in row])

        if not header_row:
            if data_rows:
                col_count = max(len(r) for r in data_rows)
            else:
                col_count = 0
            header_row = [_col_letter(i) for i in range(col_count)]

        section_lines: list[str] = [f"## Sheet: {sheet_name}", ""]

        if not header_row:
            section_lines.append("*Empty sheet.*")
        else:
            col_count = len(header_row)
            section_lines.append("| " + " | ".join(header_row) + " |")
            section_lines.append("| " + " | ".join(["---"] * col_count) + " |")

            for row in data_rows:
                padded = row + [""] * (col_count - len(row))
                section_lines.append(
                    "| " + " | ".join(padded[:col_count]) + " |"
                )

            if total_rows > len(data_rows):
                section_lines.append("")
                section_lines.append(
                    f"*Showing {len(data_rows)} of {total_rows:,} rows.*"
                )

        sections.append("\n".join(section_lines))

    return "\n\n---\n\n".join(sections)
