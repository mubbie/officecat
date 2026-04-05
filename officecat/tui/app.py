"""Textual TUI app for officecat — renders markdown in an interactive viewer."""

from __future__ import annotations

from textual.app import App, ComposeResult
from textual.binding import Binding
from textual.containers import VerticalScroll
from textual.message import Message
from textual.widgets import Footer, Header, Input, Markdown, Static


class SearchBar(Static):
    """Inline search bar that docks at the bottom."""

    DEFAULT_CSS = """
    SearchBar {
        dock: bottom;
        height: auto;
        max-height: 3;
        display: none;
        padding: 0 1;
        background: $surface;
    }
    SearchBar.visible {
        display: block;
    }
    SearchBar Input {
        width: 1fr;
        border: none;
        height: 1;
        padding: 0;
    }
    SearchBar .search-status {
        color: $text-muted;
        height: 1;
    }
    """

    def compose(self) -> ComposeResult:
        yield Input(placeholder="Search...", id="search-input")
        yield Static("", classes="search-status", id="search-status")

    def open(self) -> None:
        self.add_class("visible")
        inp = self.query_one("#search-input", Input)
        inp.value = ""
        inp.focus()
        self._update_status("")

    def close(self) -> None:
        self.remove_class("visible")
        self._update_status("")

    @property
    def is_open(self) -> bool:
        return self.has_class("visible")

    def _update_status(self, text: str) -> None:
        self.query_one("#search-status", Static).update(text)

    def on_input_changed(self, event: Input.Changed) -> None:
        """Notify the app that the search query changed."""
        self.post_message(SearchQueryChanged(event.value))

    def on_key(self, event) -> None:
        if event.key == "escape":
            self.close()
            self.post_message(SearchQueryChanged(""))
            event.stop()
        elif event.key == "enter":
            self.post_message(SearchNextRequested())
            event.stop()


class SearchQueryChanged(Message):
    """Posted when the search query changes."""

    def __init__(self, query: str) -> None:
        super().__init__()
        self.query = query


class SearchNextRequested(Message):
    """Posted when the user presses Enter in the search bar."""
    pass


class OfficeCatApp(App):
    """Interactive terminal viewer for Office files."""

    CSS = """
    #md-scroll {
        scrollbar-gutter: stable;
    }
    .search-match {
        background: $warning 40%;
    }
    """

    BINDINGS = [
        Binding("q", "quit", "Quit"),
        Binding("escape", "escape", "Back", show=False),
        Binding("slash", "open_search", "Search", show=True),
        Binding("n", "next_match", "Next", show=False),
        Binding("shift+n", "prev_match", "Prev", show=False),
    ]

    def __init__(self, source: str, markdown: str, **kwargs: object) -> None:
        self._source = source
        self._markdown = markdown
        self._search_matches: list = []
        self._match_index: int = -1
        super().__init__(**kwargs)
        self.title = f"officecat — {self._source}"

    def compose(self) -> ComposeResult:
        yield Header()
        with VerticalScroll(id="md-scroll"):
            yield Markdown(self._markdown, id="md-view")
        yield SearchBar()
        yield Footer()

    def action_open_search(self) -> None:
        self.query_one(SearchBar).open()

    def action_escape(self) -> None:
        bar = self.query_one(SearchBar)
        if bar.is_open:
            bar.close()
            self._clear_search()
        else:
            self.exit()

    def action_next_match(self) -> None:
        bar = self.query_one(SearchBar)
        if bar.is_open and self._search_matches:
            self._match_index = (self._match_index + 1) % len(self._search_matches)
            self._scroll_to_match()

    def action_prev_match(self) -> None:
        bar = self.query_one(SearchBar)
        if bar.is_open and self._search_matches:
            self._match_index = (self._match_index - 1) % len(self._search_matches)
            self._scroll_to_match()

    def on_search_query_changed(self, event: SearchQueryChanged) -> None:
        self._clear_search()
        query = event.query.strip()
        if not query:
            self.query_one(SearchBar)._update_status("")
            return
        self._find_matches(query)

    def on_search_next_requested(self, event: SearchNextRequested) -> None:
        self.action_next_match()

    def _find_matches(self, query: str) -> None:
        """Find markdown blocks containing the query and highlight them."""
        query_lower = query.lower()
        md_widget = self.query_one("#md-view", Markdown)

        # Search through the rendered child widgets of the Markdown widget
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
            bar._update_status("No matches")

    def _scroll_to_match(self) -> None:
        if 0 <= self._match_index < len(self._search_matches):
            widget = self._search_matches[self._match_index]
            widget.scroll_visible(animate=False)
            bar = self.query_one(SearchBar)
            bar._update_status(
                f"{self._match_index + 1}/{len(self._search_matches)}"
            )

    def _clear_search(self) -> None:
        for w in self._search_matches:
            w.remove_class("search-match")
        self._search_matches = []
        self._match_index = -1
