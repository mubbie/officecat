"""Rich colored markdown renderer."""

from __future__ import annotations

import sys


def render(markdown_text: str, head: int | None = None) -> None:
    """Print colored markdown to stdout."""
    if sys.platform == "win32" and hasattr(sys.stdout, "reconfigure"):
        try:
            sys.stdout.reconfigure(encoding="utf-8")
        except Exception:
            pass

    from rich.console import Console
    from rich.markdown import Markdown

    text = markdown_text
    if head is not None:
        text = "\n".join(text.splitlines()[:head])

    console = Console()
    console.print(Markdown(text))
