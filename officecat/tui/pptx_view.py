"""TUI view for PowerPoint presentations."""

from __future__ import annotations

from textual.app import ComposeResult
from textual.containers import Vertical, VerticalScroll
from textual.widgets import Label, Static, TabbedContent, TabPane

from officecat.models import Presentation, Slide


class SlideView(VerticalScroll):
    """A single slide rendered as a scrollable panel."""

    def __init__(self, slide: Slide, **kwargs: object) -> None:
        super().__init__(**kwargs)
        self._slide = slide

    def compose(self) -> ComposeResult:
        s = self._slide
        if s.title:
            yield Static(f"[bold cyan]{s.title}[/bold cyan]\n", markup=True)

        for line in s.body:
            yield Static(f"  {line}")

        for img in s.images:
            yield Static(f"  [dim]{img}[/dim]", markup=True)

        if s.notes:
            yield Static("")
            yield Static(f"  [dim italic]Notes: {s.notes}[/dim italic]", markup=True)


class PptxView(Vertical):
    """Tabbed slide navigation for presentations."""

    def __init__(self, presentation: Presentation, **kwargs: object) -> None:
        super().__init__(**kwargs)
        self._presentation = presentation

    def compose(self) -> ComposeResult:
        pres = self._presentation
        if not pres.slides:
            yield Static("Presentation is empty.")
            return

        total = pres.total_slides
        yield Label(
            f"  {len(pres.slides)} of {total} slides" if len(pres.slides) < total
            else f"  {total} slides",
            classes="status-dim",
        )

        with TabbedContent():
            for slide in pres.slides:
                tab_label = f"Slide {slide.number}"
                if slide.title:
                    tab_label += f": {slide.title}"
                with TabPane(tab_label):
                    yield SlideView(slide)
