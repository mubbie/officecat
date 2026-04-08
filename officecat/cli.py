"""CLI entry point — read and render pipeline."""

from __future__ import annotations

import sys
from pathlib import Path
from typing import Annotated, Optional

import typer

from officecat import __version__

app = typer.Typer(
    add_completion=False,
    invoke_without_command=True,
    no_args_is_help=True,
)


@app.callback()
def main(
    ctx: typer.Context,
    version: Annotated[
        Optional[bool],
        typer.Option("--version", "-v", help="Show version."),
    ] = None,
) -> None:
    """View Office files in the terminal."""
    if version:
        print(f"officecat {__version__}")
        raise typer.Exit()


@app.command(hidden=True)
def view(
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
        Optional[int],
        typer.Option("--head", "-n", help="Show first N lines."),
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
        typer.Option(
            "--headers", "-h", help="Row N as headers (default: 1)."
        ),
    ] = 1,
    show_all: Annotated[
        bool, typer.Option("--all", "-a", help="Disable the row cap.")
    ] = False,
) -> None:
    """View an Office file."""
    _view_file(
        file,
        tui=tui,
        plain=plain,
        json=json,
        head=head,
        sheet=sheet,
        slide=slide,
        headers=headers,
        show_all=show_all,
    )


@app.command()
def update() -> None:
    """Update officecat to the latest version."""
    import subprocess

    from rich.console import Console

    console = Console(stderr=True)

    console.print(f"[dim]Current version: {__version__}[/dim]")
    console.print("[dim]Checking for updates...[/dim]")

    try:
        result = subprocess.run(
            [
                sys.executable, "-m", "pip", "install",
                "--upgrade", "officecat",
            ],
            capture_output=True,
            text=True,
        )
    except Exception as e:
        console.print(f"[red]Update failed: {e}[/red]")
        raise typer.Exit(1)

    if result.returncode != 0:
        console.print(f"[red]Update failed:[/red]\n{result.stderr}")
        raise typer.Exit(1)

    if "already satisfied" in result.stdout.lower():
        console.print(
            f"[green]Already up to date ({__version__}).[/green]"
        )
    else:
        # Re-check version after upgrade
        try:
            ver_result = subprocess.run(
                [
                    sys.executable, "-c",
                    "from officecat import __version__; "
                    "print(__version__)",
                ],
                capture_output=True,
                text=True,
            )
            new_ver = ver_result.stdout.strip()
        except Exception:
            new_ver = "unknown"
        console.print(
            f"[green]Updated: {__version__} -> {new_ver}[/green]"
        )


def _view_file(
    file: Path,
    *,
    tui: bool,
    plain: bool,
    json: bool,
    head: int | None,
    sheet: str | None,
    slide: int | None,
    headers: int,
    show_all: bool,
) -> None:
    """Core file viewing logic."""
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
    if head is not None and fmt_name in ("xlsx", "csv"):
        reader_opts["head"] = head

    # ── Convert (with spinner for TTY) ──
    from officecat.readers import convert

    if sys.stderr.isatty() and mode != "tui":
        from rich.console import Console

        console = Console(stderr=True)
        with console.status(
            f"Reading {file.name}...", spinner="dots"
        ):
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


def cli_entry() -> None:
    """Entry point — route bare file args to view command."""
    # If the first arg looks like a file (not a subcommand),
    # inject "view" so typer routes it correctly.
    if len(sys.argv) > 1:
        first = sys.argv[1]
        if (
            first not in ("update", "--help", "--version", "-v")
            and not first.startswith("--")
        ):
            sys.argv.insert(1, "view")
    app()
