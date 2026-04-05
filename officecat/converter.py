"""Thin wrapper around MarkItDown for file-to-markdown conversion."""

from __future__ import annotations

import sys
from dataclasses import dataclass
from pathlib import Path

SUPPORTED_EXTENSIONS = {".docx", ".pptx", ".xlsx", ".xls", ".csv", ".tsv"}

LEGACY_FORMATS = {
    ".doc": ".docx",
    ".ppt": ".pptx",
}


@dataclass
class ConversionResult:
    source: str
    markdown: str


def validate_format(path: Path) -> None:
    """Check file extension is supported. Exits with friendly error if not."""
    ext = path.suffix.lower()

    if ext in LEGACY_FORMATS:
        target = LEGACY_FORMATS[ext]
        print(
            f"Error: Legacy binary format ({ext}) is not supported.\n"
            f"Convert to {target} using LibreOffice: "
            f"libreoffice --headless --convert-to {target[1:]} {path.name}\n"
            f"Or use specialized tools like antiword or catdoc.",
            file=sys.stderr,
        )
        raise SystemExit(2)

    if ext not in SUPPORTED_EXTENSIONS:
        supported = ", ".join(sorted(SUPPORTED_EXTENSIONS))
        print(
            f"Error: Unsupported file type '{ext}'. Supported: {supported}",
            file=sys.stderr,
        )
        raise SystemExit(2)


def convert(path: Path) -> ConversionResult:
    """Convert a file to markdown using MarkItDown."""
    import warnings

    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        from markitdown import MarkItDown

    try:
        md = MarkItDown()
        result = md.convert(str(path))
    except FileNotFoundError:
        print(f"Error: File '{path}' not found.", file=sys.stderr)
        raise SystemExit(1)
    except Exception as e:
        msg = str(e).lower()
        if "password" in msg or "encrypt" in msg:
            print(
                f"Error: '{path.name}' is password-protected. "
                f"officecat cannot open encrypted files.",
                file=sys.stderr,
            )
            raise SystemExit(3)
        if "bad zip" in msg or "not a zip" in msg or "corrupt" in msg:
            print(
                f"Error: '{path.name}' appears to be corrupt or invalid.",
                file=sys.stderr,
            )
            raise SystemExit(3)
        print(
            f"Error: Could not convert '{path.name}'. Format may not be supported.",
            file=sys.stderr,
        )
        raise SystemExit(2)

    text = result.text_content if result.text_content else ""

    return ConversionResult(source=str(path), markdown=text.strip())
