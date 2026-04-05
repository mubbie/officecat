"""Textual TUI app for officecat — renders markdown with MarkdownViewer."""

from __future__ import annotations

from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.containers import Horizontal
from textual.screen import ModalScreen
from textual.widgets import Footer, Header, Input, Label, MarkdownViewer


class SearchScreen(ModalScreen[None]):
    """Modal search overlay — appears at the bottom, returns focus cleanly."""

    DEFAULT_CSS = """
    SearchScreen {
        align: center bottom;
    }
    #search-bar {
        dock: bottom;
        width: 100%;
        height: 1;
        background: $surface;
        padding: 0 1;
    }
    #search-bar Label {
        width: auto;
        color: $accent;
    }
    #search-bar Input {
        width: 1fr;
        border: none;
        height: 1;
        padding: 0;
        background: $surface;
    }
    #search-bar .search-count {
        width: auto;
        color: $text-muted;
    }
    """

    BINDINGS = [
        Binding("escape", "close", "Close", show=False),
    ]

    def __init__(self) -> None:
        super().__init__()
        self._matches: list = []
        self._match_index: int = -1

    def compose(self) -> ComposeResult:
        with Horizontal(id="search-bar"):
            yield Label(" / ")
            yield Input(placeholder="Search...", id="search-input")
            yield Label("", id="search-count", classes="search-count")

    def on_mount(self) -> None:
        self.query_one("#search-input", Input).focus()

    def on_input_changed(self, event: Input.Changed) -> None:
        self._clear_highlights()
        query = event.value.strip()
        if not query:
            self._update_count("")
            return
        self._find_matches(query)

    def on_input_submitted(self, event: Input.Submitted) -> None:
        """Enter key — go to next match."""
        self._next_match()

    def on_key(self, event) -> None:
        # Ctrl+N / Ctrl+P for next/prev without leaving input
        if event.key == "ctrl+n":
            self._next_match()
            event.stop()
        elif event.key == "ctrl+p":
            self._prev_match()
            event.stop()

    def action_close(self) -> None:
        self._clear_highlights()
        self.dismiss(None)

    def _find_matches(self, query: str) -> None:
        query_lower = query.lower()
        viewer = self.app.query_one(MarkdownViewer)
        md_widget = viewer.document

        for child in md_widget.children:
            try:
                rendered = child.render()
                text = rendered.plain if hasattr(rendered, "plain") else str(rendered)
            except Exception:
                continue
            if query_lower in text.lower():
                child.add_class("search-match")
                self._matches.append(child)

        if self._matches:
            self._match_index = 0
            self._scroll_to_current()
        else:
            self._update_count("No matches")

    def _next_match(self) -> None:
        if not self._matches:
            return
        self._match_index = (self._match_index + 1) % len(self._matches)
        self._scroll_to_current()

    def _prev_match(self) -> None:
        if not self._matches:
            return
        self._match_index = (self._match_index - 1) % len(self._matches)
        self._scroll_to_current()

    def _scroll_to_current(self) -> None:
        if 0 <= self._match_index < len(self._matches):
            self._matches[self._match_index].scroll_visible(animate=False)
            self._update_count(f"{self._match_index + 1}/{len(self._matches)}")

    def _clear_highlights(self) -> None:
        for w in self._matches:
            w.remove_class("search-match")
        self._matches = []
        self._match_index = -1

    def _update_count(self, text: str) -> None:
        self.query_one("#search-count", Label).update(text)


class OfficeCatApp(App):
    """Interactive terminal viewer for Office files."""

    CSS = """
    .search-match {
        background: $warning 40%;
    }
    MarkdownViewer {
        height: 1fr;
    }
    """

    BINDINGS = [
        Binding("q", "quit", "Quit"),
        Binding("t", "toggle_toc", "TOC", show=True),
        Binding("slash", "search", "Search", show=True),
    ]

    def __init__(self, source: str, markdown: str, **kwargs: object) -> None:
        self._source = source
        self._markdown = markdown
        super().__init__(**kwargs)
        self.title = f"officecat — {self._source}"

    def compose(self) -> ComposeResult:
        yield Header()
        yield MarkdownViewer(
            self._markdown,
            show_table_of_contents=True,
            open_links=False,
            id="viewer",
        )
        yield Footer()

    def action_toggle_toc(self) -> None:
        viewer = self.query_one(MarkdownViewer)
        viewer.show_table_of_contents = not viewer.show_table_of_contents

    def action_search(self) -> None:
        self.push_screen(SearchScreen())
