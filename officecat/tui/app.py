"""Textual TUI app for officecat — renders markdown with MarkdownViewer."""

from __future__ import annotations

from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.widget import Widget
from textual.widgets import Footer, Header, Input, Label, MarkdownViewer


class SearchBar(Widget):
    """Search bar docked at the bottom of the app."""

    DEFAULT_CSS = """
    SearchBar {
        dock: bottom;
        height: 3;
        display: none;
        background: $panel;
        layout: horizontal;
        padding: 0 1;
    }
    SearchBar.open {
        display: block;
    }
    SearchBar #search-prompt {
        width: 3;
        height: 3;
        content-align: left middle;
        color: $accent;
    }
    SearchBar Input {
        width: 1fr;
        color: $text;
        background: $surface;
    }
    SearchBar #search-count {
        width: auto;
        height: 3;
        content-align: left middle;
        color: $text-muted;
        padding: 0 1;
    }
    """

    def compose(self) -> ComposeResult:
        yield Label(" / ", id="search-prompt")
        yield Input(placeholder="Type to search...", id="search-input")
        yield Label("", id="search-count")

    def open(self) -> None:
        self.add_class("open")
        inp = self.query_one("#search-input", Input)
        inp.value = ""
        inp.focus()

    def close(self) -> None:
        self.remove_class("open")
        self.query_one("#search-count", Label).update("")

    @property
    def is_open(self) -> bool:
        return self.has_class("open")

    def set_count(self, text: str) -> None:
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
        Binding("slash", "open_search", "Search", show=True),
    ]

    def __init__(self, source: str, markdown: str, **kwargs: object) -> None:
        self._source = source
        self._markdown = markdown
        self._search_matches: list[Widget] = []
        self._match_index: int = -1
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
        yield SearchBar()
        yield Footer()

    def action_toggle_toc(self) -> None:
        viewer = self.query_one(MarkdownViewer)
        viewer.show_table_of_contents = not viewer.show_table_of_contents

    def action_open_search(self) -> None:
        bar = self.query_one(SearchBar)
        if not bar.is_open:
            bar.open()

    # ── Search input events (bubble up from SearchBar) ──

    def on_input_changed(self, event: Input.Changed) -> None:
        if event.input.id != "search-input":
            return
        self._clear_highlights()
        query = event.value.strip()
        if not query:
            self.query_one(SearchBar).set_count("")
            return
        self._find_matches(query)

    def on_input_submitted(self, event: Input.Submitted) -> None:
        if event.input.id != "search-input":
            return
        self._next_match()

    def on_key(self, event) -> None:
        bar = self.query_one(SearchBar)
        if not bar.is_open:
            return

        if event.key == "escape":
            self._close_search()
            event.stop()
            event.prevent_default()

    def _close_search(self) -> None:
        self._clear_highlights()
        bar = self.query_one(SearchBar)
        bar.close()
        self.query_one(MarkdownViewer).focus()

    def _find_matches(self, query: str) -> None:
        query_lower = query.lower()
        viewer = self.query_one(MarkdownViewer)
        md_widget = viewer.document

        for child in md_widget.children:
            try:
                rendered = child.render()
                text = rendered.plain if hasattr(rendered, "plain") else str(rendered)
            except Exception:
                continue
            if query_lower in text.lower():
                child.add_class("search-match")
                self._search_matches.append(child)

        bar = self.query_one(SearchBar)
        if self._search_matches:
            self._match_index = 0
            self._scroll_to_match()
        else:
            bar.set_count("No matches")

    def _next_match(self) -> None:
        if not self._search_matches:
            return
        self._match_index = (self._match_index + 1) % len(self._search_matches)
        self._scroll_to_match()

    def _prev_match(self) -> None:
        if not self._search_matches:
            return
        self._match_index = (self._match_index - 1) % len(self._search_matches)
        self._scroll_to_match()

    def _scroll_to_match(self) -> None:
        if 0 <= self._match_index < len(self._search_matches):
            self._search_matches[self._match_index].scroll_visible(animate=False)
            self.query_one(SearchBar).set_count(
                f"{self._match_index + 1}/{len(self._search_matches)}"
            )

    def _clear_highlights(self) -> None:
        for w in self._search_matches:
            w.remove_class("search-match")
        self._search_matches = []
        self._match_index = -1
