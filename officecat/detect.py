"""File type detection and validation."""

from __future__ import annotations

import sys
from enum import Enum
from pathlib import Path


class FileFormat(Enum):
    DOCX = "docx"
    PPTX = "pptx"
    XLSX = "xlsx"
    CSV = "csv"


EXTENSION_MAP: dict[str, FileFormat] = {
    ".docx": FileFormat.DOCX,
    ".pptx": FileFormat.PPTX,
    ".xlsx": FileFormat.XLSX,
    ".csv": FileFormat.CSV,
    ".tsv": FileFormat.CSV,
}

LEGACY_FORMATS: dict[str, str] = {
    ".doc": "docx",
    ".ppt": "pptx",
    ".xls": "xlsx",
}

SUPPORTED_EXTENSIONS = ", ".join(sorted(EXTENSION_MAP.keys()))


def detect_format(path: Path) -> FileFormat:
    """Detect file format from extension. Exits on error."""
    ext = path.suffix.lower()

    if ext in LEGACY_FORMATS:
        target = LEGACY_FORMATS[ext]
        print(
            f"Error: Legacy binary format ({ext}) is not supported.\n"
            f"Convert to .{target} using LibreOffice: "
            f"libreoffice --headless --convert-to {target} {path.name}\n"
            f"Or use specialized tools like antiword or catdoc.",
            file=sys.stderr,
        )
        sys.exit(2)

    fmt = EXTENSION_MAP.get(ext)
    if fmt is None:
        print(
            f"Error: Unsupported file type '{ext}'. "
            f"Supported: {SUPPORTED_EXTENSIONS}",
            file=sys.stderr,
        )
        sys.exit(2)

    return fmt
