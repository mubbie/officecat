"""Plain text renderer — raw markdown via print()."""

from __future__ import annotations


def render(markdown_text: str, head: int | None = None) -> None:
    """Print raw markdown text to stdout."""
    lines = markdown_text.splitlines()
    if head is not None:
        lines = lines[:head]
    for line in lines:
        print(line)
