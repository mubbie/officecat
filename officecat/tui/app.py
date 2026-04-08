"""Textual TUI app for officecat — full-screen markdown viewer."""

from __future__ import annotations

from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.containers import VerticalScroll
from textual.widgets import Footer, Header, Markdown


class OfficeCatApp(App):
    """Interactive terminal viewer for Office files."""

    CSS = """
    #md-scroll {
        height: 1fr;
    }
    """

    BINDINGS = [
        Binding("q", "quit", "Quit"),
    ]

    def __init__(self, source: str, markdown: str, **kwargs: object) -> None:
        self._source = source
        self._markdown = markdown
        super().__init__(**kwargs)
        self.title = f"officecat — {self._source}"

    def compose(self) -> ComposeResult:
        yield Header()
        with VerticalScroll(id="md-scroll"):
            yield Markdown(self._markdown, id="md-view")
        yield Footer()
