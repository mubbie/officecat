"""JSON renderer."""

from __future__ import annotations

import dataclasses
import json

from officecat.models import CsvData, DocxDocument, Presentation, Workbook


def _render(obj: object) -> None:
    print(json.dumps(dataclasses.asdict(obj), indent=2))


def render_docx(doc: DocxDocument) -> None:
    _render(doc)


def render_pptx(pres: Presentation) -> None:
    _render(pres)


def render_xlsx(wb: Workbook) -> None:
    _render(wb)


def render_csv(data: CsvData) -> None:
    _render(data)
