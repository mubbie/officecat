"""TUI view for Word documents."""

from __future__ import annotations

from rich.table import Table
from rich import box

from textual.app import ComposeResult
from textual.containers import VerticalScroll
from textual.widgets import Static

from officecat.models import DocxDocument, DocxParagraph, TableData

STYLE_MAP = {
    "heading1": "[bold cyan]{text}[/bold cyan]",
    "heading2": "  [bold bright_blue]{text}[/bold bright_blue]",
    "heading3": "  [bold]{text}[/bold]",
    "heading4": "  [dim bold]{text}[/dim bold]",
    "list_item": "  \u2022 {text}",
    "body": "  {text}",
}


class DocxView(VerticalScroll):
    """Scrollable view of a Word document."""

    def __init__(self, document: DocxDocument, **kwargs: object) -> None:
        super().__init__(**kwargs)
        self._document = document

    def compose(self) -> ComposeResult:
        if not self._document.blocks:
            yield Static("Document is empty.")
            return

        for block in self._document.blocks:
            if isinstance(block, DocxParagraph):
                template = STYLE_MAP.get(block.style, "  {text}")
                markup = template.format(text=block.text)
                yield Static(markup, markup=True)
            elif isinstance(block, TableData):
                table = Table(box=box.ROUNDED, show_lines=True)
                for h in block.headers:
                    table.add_column(h, style="cyan")
                for row in block.rows:
                    table.add_row(*row)
                yield Static(table)
