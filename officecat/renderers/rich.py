"""Rich terminal renderer (default)."""

from __future__ import annotations

from rich.console import Console
from rich.panel import Panel
from rich.rule import Rule
from rich.table import Table
from rich import box

from officecat.models import (
    CsvData,
    DocxDocument,
    DocxParagraph,
    Presentation,
    TableData,
    Workbook,
)

console = Console()


def _render_table_data(td: TableData) -> Table:
    table = Table(box=box.ROUNDED, show_lines=True)
    for h in td.headers:
        table.add_column(h, style="cyan")
    for row in td.rows:
        table.add_row(*row)
    return table


def render_docx(doc: DocxDocument) -> None:
    if not doc.blocks:
        console.print("[dim]Document is empty.[/dim]")
        return

    for block in doc.blocks:
        if isinstance(block, DocxParagraph):
            if block.style == "heading1":
                console.print(Rule(block.text, style="bold cyan"))
            elif block.style == "heading2":
                console.print(f"  [bold bright_blue]{block.text}[/bold bright_blue]")
            elif block.style == "heading3":
                console.print(f"  [bold]{block.text}[/bold]")
            elif block.style == "heading4":
                console.print(f"  [dim bold]{block.text}[/dim bold]")
            elif block.style == "list_item":
                console.print(f"  \u2022 {block.text}")
            else:
                console.print(f"  {block.text}")
        elif isinstance(block, TableData):
            console.print()
            console.print(_render_table_data(block))


def render_pptx(pres: Presentation) -> None:
    if not pres.slides:
        console.print("[dim]Presentation is empty.[/dim]")
        return

    for slide in pres.slides:
        title = f"Slide {slide.number}"
        if slide.title:
            title += f": {slide.title}"

        body_parts: list[str] = []
        for line in slide.body:
            body_parts.append(line)

        for img in slide.images:
            body_parts.append(f"[dim]{img}[/dim]")

        if slide.notes:
            body_parts.append(f"\n[dim italic]Notes: {slide.notes}[/dim italic]")

        content = "\n".join(body_parts) if body_parts else "[dim]Empty slide[/dim]"
        console.print(Panel(content, title=title, border_style="blue"))
        console.print()

    if pres.total_slides > len(pres.slides):
        console.print(
            f"[dim]Showing {len(pres.slides)} of {pres.total_slides} slides. "
            f"Use --head to control how many slides are shown.[/dim]"
        )


def _render_sheet_table(name: str, headers: list[str], rows: list[list[str]], total_rows: int) -> None:
    table = Table(title=name, box=box.ROUNDED)
    for h in headers:
        table.add_column(h, style="cyan")
    for row in rows:
        # Pad short rows
        padded = row + [""] * (len(headers) - len(row))
        table.add_row(*padded[:len(headers)])
    console.print(table)

    if total_rows > len(rows):
        console.print(
            f"[dim]Showing {len(rows)} of {total_rows} rows. "
            f"Use --all to show everything.[/dim]"
        )


def render_xlsx(wb: Workbook) -> None:
    if not wb.sheets:
        console.print("[dim]Workbook is empty.[/dim]")
        return

    for i, sheet in enumerate(wb.sheets):
        if i > 0:
            console.print()
            console.print(Rule(style="dim"))
            console.print()
        _render_sheet_table(sheet.name, sheet.headers, sheet.rows, sheet.total_rows)


def render_csv(data: CsvData) -> None:
    if not data.rows and not data.headers:
        console.print("[dim]File is empty.[/dim]")
        return

    _render_sheet_table(data.source, data.headers, data.rows, data.total_rows)
