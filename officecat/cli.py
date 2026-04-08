"""CLI entry point — read and render pipeline."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Annotated, Optional

import typer

app = typer.Typer(add_completion=False)

# Formats that support tabular flags
_TABULAR_FMTS = {"xlsx", "csv"}


@app.command()
def run(
    file: Annotated[
        Path, typer.Argument(help="File to view.")
    ],
    tui: Annotated[
        bool, typer.Option("--tui", "-t", help="Interactive viewer.")
    ] = False,
    plain: Annotated[
        bool, typer.Option("--plain", "-p", help="Raw markdown, no colors.")
    ] = False,
    json: Annotated[
        bool, typer.Option("--json", "-j", help="JSON output.")
    ] = False,
    head: Annotated[
        Optional[int], typer.Option("--head", "-n", help="Show first N lines.")
    ] = None,
    sheet: Annotated[
        Optional[str],
        typer.Option("--sheet", "-s", help="Sheet by name or index (xlsx)."),
    ] = None,
    slide: Annotated[
        Optional[int],
        typer.Option("--slide", help="Show only slide N (pptx)."),
    ] = None,
    headers: Annotated[
        int,
        typer.Option("--headers", "-h", help="Row N as headers (default: 1)."),
    ] = 1,
    show_all: Annotated[
        bool, typer.Option("--all", "-a", help="Disable the row cap.")
    ] = False,
) -> None:
    """View Office files in the terminal.

    Supports .docx, .pptx, .xlsx, .csv, and .tsv files.
    """
    # ── Validate output flags ──
    output_flags = sum([tui, plain, json])
    if output_flags > 1:
        _error("--tui, --plain, and --json are mutually exclusive.")

    if not file.exists():
        _error(f"File '{file}' not found.")

    # ── Validate format ──
    from officecat.detect import detect_format
    fmt = detect_format(file)

    # ── Validate format-specific flags ──
    fmt_name = fmt.value
    if sheet is not None and fmt_name not in ("xlsx", "csv"):
        _error("--sheet is only valid for xlsx, csv, and tsv files.")
    if slide is not None and fmt_name != "pptx":
        _error("--slide is only valid for pptx files.")
    if headers != 1 and fmt_name not in ("xlsx", "csv"):
        _error("--headers is only valid for xlsx, csv, and tsv files.")

    # ── Mode resolution ──
    if json:
        mode = "json"
    elif plain:
        mode = "plain"
    elif tui:
        mode = "tui"
    elif sys.stdout.isatty():
        mode = "rich"
    else:
        mode = "plain"

    # ── Build reader options ──
    reader_opts: dict = {}
    if fmt_name in ("xlsx", "csv"):
        reader_opts["headers"] = headers
        reader_opts["show_all"] = show_all
        if sheet is not None:
            reader_opts["sheet"] = sheet
    if fmt_name == "pptx" and slide is not None:
        reader_opts["slide"] = slide
    if head is not None and fmt_name in ("docx", "pptx"):
        reader_opts["head"] = head

    # For tabular formats, pass head as row limit to reader
    if head is not None and fmt_name in ("xlsx", "csv"):
        reader_opts["head"] = head

    # ── Convert (with spinner for TTY) ──
    from officecat.readers import convert

    if sys.stderr.isatty() and mode != "tui":
        from rich.console import Console
        console = Console(stderr=True)
        with console.status(f"Reading {file.name}...", spinner="dots"):
            markdown = convert(file, **reader_opts)
    else:
        markdown = convert(file, **reader_opts)

    if not markdown or not markdown.strip():
        print("Document is empty.")
        return

    # ── TUI guardrail for large tables ──
    if mode == "tui" and show_all and fmt_name in ("xlsx", "csv"):
        line_count = markdown.count("\n")
        if line_count > 1000:
            # Truncate to ~1000 lines for TUI performance
            lines = markdown.splitlines()
            markdown = "\n".join(lines[:1000])
            markdown += (
                "\n\n*TUI limited to first 1000 lines."
                " Use --plain with --all to view everything.*"
            )

    # ── Render ──
    if mode == "tui":
        from officecat.tui.app import OfficeCatApp
        tui_app = OfficeCatApp(source=str(file), markdown=markdown)
        tui_app.run()
    elif mode == "rich":
        from officecat.renderers import rich as _rich
        _rich.render(markdown, head=head)
    elif mode == "json":
        from officecat.renderers import json_ as _json
        _json.render(str(file), markdown)
    else:
        from officecat.renderers import plain as _plain
        _plain.render(markdown, head=head)


def _error(msg: str) -> None:
    print(f"Error: {msg}", file=sys.stderr)
    raise SystemExit(1)


def main() -> None:
    app()
