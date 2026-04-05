"""JSON renderer."""

from __future__ import annotations

import json


def render_json(source: str, markdown_text: str) -> None:
    output = {
        "source": source,
        "markdown": markdown_text,
    }
    print(json.dumps(output, indent=2))
