"""CLI entry point — detect, read, render pipeline."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Annotated, Optional
from zipfile import BadZipFile

import typer

from officecat.detect import FileFormat, detect_format

app = typer.Typer(add_completion=False)


@app.command()
def run(
    file: Annotated[Path, typer.Argument(help="File to view.")],
    rich: Annotated[bool, typer.Option("--rich", "-r", help="Colored formatted dump (non-interactive).")] = False,
    plain: Annotated[bool, typer.Option("--plain", "-p", help="Plain text output, no colors.")] = False,
    json: Annotated[bool, typer.Option("--json", "-j", help="JSON output.")] = False,
    head: Annotated[Optional[int], typer.Option("--head", "-n", help="Show first N items.")] = None,
    list_only: Annotated[bool, typer.Option("--list", "-l", help="List contents only.")] = False,
    sheet: Annotated[Optional[str], typer.Option("--sheet", "-s", help="Select sheet by name or 1-based index (xlsx/csv only).")] = None,
    slide: Annotated[Optional[int], typer.Option("--slide", help="Show only slide N (pptx only).")] = None,
    headers: Annotated[int, typer.Option("--headers", "-h", help="Promote row N as headers (xlsx/csv, default: 1, 0 to disable).")] = 1,
    show_all: Annotated[bool, typer.Option("--all", "-a", help="Disable the default row cap.")] = False,
) -> None:
    """View Office files in the terminal."""
    # ── Validate flags ──
    output_flags = sum([rich, plain, json])
    if output_flags > 1:
        _error("--rich, --plain, and --json are mutually exclusive.")

    if not file.exists():
        _error(f"File '{file}' not found.")

    fmt = detect_format(file)

    if sheet is not None and fmt not in (FileFormat.XLSX, FileFormat.CSV):
        _error("--sheet is only valid for xlsx, csv, and tsv files.")

    if slide is not None and fmt != FileFormat.PPTX:
        _error("--slide is only valid for pptx files.")

    if headers != 1 and fmt not in (FileFormat.XLSX, FileFormat.CSV):
        _error("--headers is only valid for xlsx, csv, and tsv files.")

    # ── Mode resolution ──
    if json:
        mode = "json"
    elif plain:
        mode = "plain"
    elif rich:
        mode = "rich"
    elif list_only:
        mode = "plain"
    elif sys.stdout.isatty():
        mode = "tui"
    else:
        mode = "plain"

    # ── Read data ──
    try:
        data = _read(fmt, file, head=head, sheet=sheet, slide=slide, headers=headers, show_all=show_all)
    except BadZipFile:
        print(
            f"Error: '{file.name}' appears to be corrupt or is not a valid {fmt.value} file.",
            file=sys.stderr,
        )
        raise SystemExit(3)
    except KeyboardInterrupt:
        raise SystemExit(130)

    # ── List mode (always non-interactive) ──
    if list_only:
        _list_contents(fmt, data)
        return

    # ── Render ──
    if mode == "tui":
        from officecat.tui.app import OfficeCatApp

        tui_app = OfficeCatApp(data=data, fmt=fmt)
        tui_app.run()
    elif mode == "rich":
        _render_with("rich", fmt, data)
    elif mode == "json":
        _render_with("json_", fmt, data)
    else:
        _render_with("plain", fmt, data)


def _error(msg: str) -> None:
    print(f"Error: {msg}", file=sys.stderr)
    raise SystemExit(1)


def _read(
    fmt: FileFormat,
    path: Path,
    *,
    head: int | None,
    sheet: str | None,
    slide: int | None,
    headers: int,
    show_all: bool,
) -> object:
    if fmt == FileFormat.XLSX:
        from officecat.readers.xlsx import read_xlsx
        return read_xlsx(path, head=head, sheet=sheet, headers=headers, show_all=show_all)
    elif fmt == FileFormat.DOCX:
        from officecat.readers.docx import read_docx
        return read_docx(path, head=head)
    elif fmt == FileFormat.PPTX:
        from officecat.readers.pptx import read_pptx
        return read_pptx(path, head=head, slide=slide)
    elif fmt == FileFormat.CSV:
        from officecat.readers.csv_ import read_csv
        return read_csv(path, head=head, headers=headers)
    else:
        _error(f"Unhandled format: {fmt}")


def _list_contents(fmt: FileFormat, data: object) -> None:
    from officecat.models import DocxParagraph

    if fmt == FileFormat.XLSX:
        for name in data.sheet_names:
            print(name)
    elif fmt == FileFormat.DOCX:
        for block in data.blocks:
            if isinstance(block, DocxParagraph) and block.style.startswith("heading"):
                depth = block.style[-1]
                indent = "  " * (int(depth) - 1)
                print(f"{indent}{block.text}")
    elif fmt == FileFormat.PPTX:
        for s in data.slides:
            title = s.title or "Untitled"
            print(f"Slide {s.number}: {title}")
    elif fmt == FileFormat.CSV:
        if data.headers:
            print(", ".join(data.headers))


def _render_with(renderer_name: str, fmt: FileFormat, data: object) -> None:
    import importlib
    renderer = importlib.import_module(f"officecat.renderers.{renderer_name}")

    dispatch = {
        FileFormat.XLSX: renderer.render_xlsx,
        FileFormat.DOCX: renderer.render_docx,
        FileFormat.PPTX: renderer.render_pptx,
        FileFormat.CSV: renderer.render_csv,
    }
    dispatch[fmt](data)


def main() -> None:
    app()
