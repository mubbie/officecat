"""JSON renderer."""

from __future__ import annotations

import json


def render(source: str, markdown_text: str) -> None:
    """Print JSON with source and markdown keys."""
    output = {"source": source, "markdown": markdown_text}
    print(json.dumps(output, indent=2))
