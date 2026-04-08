"""Reader dispatch — routes a file path to the correct format reader."""

from __future__ import annotations

from pathlib import Path

from officecat.detect import FileFormat, detect_format


def convert(path: Path, **options) -> str:
    """Convert a file to a markdown string.

    Args:
        path: Path to the file.
        **options: Format-specific options (head, sheet, slide, headers, show_all).

    Returns:
        A markdown string.
    """
    fmt = detect_format(path)

    if fmt == FileFormat.CSV:
        from officecat.readers.csv_ import to_markdown
    elif fmt == FileFormat.XLSX:
        from officecat.readers.xlsx import to_markdown
    elif fmt == FileFormat.DOCX:
        from officecat.readers.docx import to_markdown
    elif fmt == FileFormat.PPTX:
        from officecat.readers.pptx import to_markdown

    return to_markdown(path, **options)
