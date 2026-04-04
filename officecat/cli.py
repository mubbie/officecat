"""CLI entry point — detect, read, render pipeline."""

from __future__ import annotations

import sys
from pathlib import Path
from zipfile import BadZipFile

import click

from officecat.detect import FileFormat, detect_format


@click.command()
@click.argument("file", type=click.Path(exists=False))
@click.option("--plain", "-p", "output_plain", is_flag=True, help="Plain text output, no colors/boxes.")
@click.option("--json", "-j", "output_json", is_flag=True, help="JSON output.")
@click.option("--head", "-n", type=int, default=None, help="Show first N items (rows/paragraphs/slides).")
@click.option("--list", "-l", "list_only", is_flag=True, help="List contents only (sheet names, slide titles, heading outline).")
@click.option("--sheet", "-s", default=None, help="Select sheet by name or 1-based index (xlsx only).")
@click.option("--slide", type=int, default=None, help="Show only slide N (pptx only).")
@click.option("--headers", "-h", type=int, default=1, help="Promote row N as column headers (xlsx/csv, default: 1, use 0 to disable).")
@click.option("--all", "-a", "show_all", is_flag=True, help="Disable the default row cap for large spreadsheets.")
def main(
    file: str,
    output_plain: bool,
    output_json: bool,
    head: int | None,
    list_only: bool,
    sheet: str | None,
    slide: int | None,
    headers: int,
    show_all: bool,
) -> None:
    """View Office files in the terminal.

    Supports .docx, .pptx, .xlsx, .csv, and .tsv files.
    """
    # ── Validate flags ──
    if output_plain and output_json:
        print("Error: --plain and --json are mutually exclusive.", file=sys.stderr)
        sys.exit(1)

    path = Path(file)
    if not path.exists():
        print(f"Error: File '{file}' not found.", file=sys.stderr)
        sys.exit(1)

    fmt = detect_format(path)

    # Format-specific flag checks
    if sheet is not None and fmt not in (FileFormat.XLSX, FileFormat.CSV):
        print("Error: --sheet is only valid for xlsx and csv files.", file=sys.stderr)
        sys.exit(1)

    if slide is not None and fmt != FileFormat.PPTX:
        print("Error: --slide is only valid for pptx files.", file=sys.stderr)
        sys.exit(1)

    if headers != 1 and fmt not in (FileFormat.XLSX, FileFormat.CSV):
        print("Error: --headers is only valid for xlsx, csv, and tsv files.", file=sys.stderr)
        sys.exit(1)

    # ── Choose renderer ──
    if output_json:
        from officecat.renderers import json_ as renderer
    elif output_plain or not sys.stdout.isatty():
        from officecat.renderers import plain as renderer
    else:
        from officecat.renderers import rich as renderer

    # ── Read and render ──
    try:
        if fmt == FileFormat.XLSX:
            from officecat.readers.xlsx import read_xlsx

            data = read_xlsx(path, head=head, sheet=sheet, headers=headers, show_all=show_all)

            if list_only:
                for name in data.sheet_names:
                    print(name)
                return

            renderer.render_xlsx(data)

        elif fmt == FileFormat.DOCX:
            from officecat.readers.docx import read_docx

            data = read_docx(path, head=head)

            if list_only:
                from officecat.models import DocxParagraph

                for block in data.blocks:
                    if isinstance(block, DocxParagraph) and block.style.startswith("heading"):
                        depth = block.style[-1]
                        indent = "  " * (int(depth) - 1)
                        print(f"{indent}{block.text}")
                return

            renderer.render_docx(data)

        elif fmt == FileFormat.PPTX:
            from officecat.readers.pptx import read_pptx

            data = read_pptx(path, head=head, slide=slide)

            if list_only:
                for s in data.slides:
                    title = s.title or "Untitled"
                    print(f"Slide {s.number}: {title}")
                return

            renderer.render_pptx(data)

        elif fmt == FileFormat.CSV:
            from officecat.readers.csv_ import read_csv

            data = read_csv(path, head=head, headers=headers)

            if list_only:
                if data.headers:
                    print(", ".join(data.headers))
                return

            renderer.render_csv(data)

    except BadZipFile:
        print(
            f"Error: '{path.name}' appears to be corrupt or is not a valid {fmt.value} file.",
            file=sys.stderr,
        )
        sys.exit(3)
    except KeyboardInterrupt:
        sys.exit(130)
