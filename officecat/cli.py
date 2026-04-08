"""CLI entry point — convert and render pipeline."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Annotated, Optional

import typer

app = typer.Typer(add_completion=False)


@app.command()
def run(
    file: Annotated[Path, typer.Argument(help="File to view.")],
    tui: Annotated[bool, typer.Option("--tui", "-t", help="Interactive TUI viewer with search and TOC.")] = False,
    plain: Annotated[bool, typer.Option("--plain", "-p", help="Raw markdown text, no colors.")] = False,
    json: Annotated[bool, typer.Option("--json", "-j", help="JSON output.")] = False,
    head: Annotated[Optional[int], typer.Option("--head", "-n", help="Show first N lines.")] = None,
    list_only: Annotated[bool, typer.Option("--list", "-l", help="Show file metadata only.")] = False,
) -> None:
    """View Office files in the terminal.

    Supports .docx, .pptx, .xlsx, .xls, .csv, and .tsv files.
    """
    # ── Validate flags ──
    output_flags = sum([tui, plain, json])
    if output_flags > 1:
        _error("--tui, --plain, and --json are mutually exclusive.")

    if not file.exists():
        _error(f"File '{file}' not found.")

    # ── Validate format ──
    from officecat.converter import validate_format
    validate_format(file)

    # ── Mode resolution ──
    if json:
        mode = "json"
    elif plain:
        mode = "plain"
    elif tui:
        mode = "tui"
    elif list_only:
        mode = "plain"
    elif sys.stdout.isatty():
        mode = "rich"
    else:
        mode = "plain"

    # ── Convert (with spinner for TTY) ──
    from officecat.converter import convert

    if sys.stderr.isatty():
        from rich.console import Console

        console = Console(stderr=True)
        with console.status(f"Converting {file.name}...", spinner="dots"):
            result = convert(file)
    else:
        result = convert(file)

    if not result.markdown:
        print("Document is empty.")
        return

    # ── List mode ──
    if list_only:
        _list_metadata(file, result)
        return

    # ── Render ──
    if mode == "tui":
        from officecat.tui.app import OfficeCatApp

        tui_app = OfficeCatApp(source=result.source, markdown=result.markdown)
        tui_app.run()
    elif mode == "rich":
        _render_rich(result.markdown, head=head)
    elif mode == "json":
        from officecat.renderers.json_ import render_json
        render_json(result.source, result.markdown)
    else:
        from officecat.renderers.plain import render_plain
        render_plain(result.markdown, head=head)


def _error(msg: str) -> None:
    print(f"Error: {msg}", file=sys.stderr)
    raise SystemExit(1)


def _list_metadata(path: Path, result) -> None:
    """Show file metadata and a heading outline."""
    import os

    size = os.path.getsize(path)
    if size < 1024:
        size_str = f"{size} B"
    elif size < 1024 * 1024:
        size_str = f"{size / 1024:.1f} KB"
    else:
        size_str = f"{size / (1024 * 1024):.1f} MB"

    print(f"File: {path.name}")
    print(f"Size: {size_str}")
    print(f"Suffix: {path.suffix}")

    headings = []
    for line in result.markdown.splitlines():
        stripped = line.strip()
        if stripped.startswith("#"):
            level = 0
            for ch in stripped:
                if ch == "#":
                    level += 1
                else:
                    break
            text = stripped[level:].strip()
            if text:
                indent = "  " * (level - 1)
                headings.append(f"{indent}{text}")

    if headings:
        print("\nOutline:")
        for h in headings:
            print(f"  {h}")


def _render_rich(markdown_text: str, *, head: int | None = None) -> None:
    import sys as _sys

    if _sys.platform == "win32" and hasattr(_sys.stdout, "reconfigure"):
        try:
            _sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    from rich.console import Console
    from rich.markdown import Markdown

    text = markdown_text
    if head is not None:
        text = "\n".join(text.splitlines()[:head])

    console = Console()
    console.print(Markdown(text))


def main() -> None:
    app()
