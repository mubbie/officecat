"""Main Textual TUI application for officecat."""

from __future__ import annotations

from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.widgets import Footer, Header

from officecat.detect import FileFormat


class OfficeCatApp(App):
    """Interactive terminal viewer for Office files."""

    TITLE = "officecat"

    CSS = """
    .status-dim {
        color: $text-muted;
        padding: 0 1;
    }
    """

    BINDINGS = [
        Binding("q", "quit", "Quit"),
        Binding("escape", "quit", "Quit", show=False),
        Binding("/", "search", "Search", show=True),
    ]

    def __init__(self, data: object, fmt: FileFormat, **kwargs: object) -> None:
        super().__init__(**kwargs)
        self._data = data
        self._fmt = fmt

    def compose(self) -> ComposeResult:
        yield Header()
        yield self._build_view()
        yield Footer()

    def _build_view(self):  # noqa: ANN202
        if self._fmt == FileFormat.XLSX:
            from officecat.tui.xlsx_view import XlsxView
            return XlsxView(self._data)
        elif self._fmt == FileFormat.DOCX:
            from officecat.tui.docx_view import DocxView
            return DocxView(self._data)
        elif self._fmt == FileFormat.PPTX:
            from officecat.tui.pptx_view import PptxView
            return PptxView(self._data)
        elif self._fmt == FileFormat.CSV:
            from officecat.tui.xlsx_view import CsvView
            return CsvView(self._data)

    def action_search(self) -> None:
        """Placeholder for search — Textual's built-in search can be wired later."""
        pass
