"""PowerPoint (.pptx) reader."""

from __future__ import annotations

import sys
from pathlib import Path

from officecat.models import Presentation, Slide


def read_pptx(
    path: Path,
    *,
    head: int | None = None,
    slide: int | None = None,
) -> Presentation:
    from pptx import Presentation as PptxPresentation
    from pptx.enum.shapes import MSO_SHAPE_TYPE
    from pptx.opc.exceptions import PackageNotFoundError

    try:
        prs = PptxPresentation(str(path))
    except PackageNotFoundError:
        print(
            f"Error: '{path.name}' appears to be corrupt or is not a valid pptx file.",
            file=sys.stderr,
        )
        sys.exit(3)
    except Exception as e:
        msg = str(e).lower()
        if "password" in msg or "encrypt" in msg:
            print(
                f"Error: '{path.name}' is password-protected. "
                f"officecat cannot open encrypted files.",
                file=sys.stderr,
            )
            sys.exit(3)
        raise

    total_slides = len(prs.slides)

    if slide is not None:
        if slide < 1 or slide > total_slides:
            print(
                f"Error: Slide {slide} not found. "
                f"Presentation has {total_slides} slides.",
                file=sys.stderr,
            )
            sys.exit(1)

    result_slides: list[Slide] = []

    for i, sld in enumerate(prs.slides, start=1):
        if slide is not None and i != slide:
            continue

        title_shape = sld.shapes.title
        title_text = title_shape.text.strip() if title_shape else None

        body: list[str] = []
        images: list[str] = []

        for shape in sld.shapes:
            # Skip the title shape to avoid duplication
            if shape is title_shape:
                continue

            st = shape.shape_type
            if st == MSO_SHAPE_TYPE.PICTURE:
                images.append(f"[Image: {shape.name}]")
            elif st == MSO_SHAPE_TYPE.GROUP:
                body.append("[Grouped content]")
            elif hasattr(shape, "has_table") and shape.has_table:
                tbl = shape.table
                rows = len(tbl.rows)
                cols = len(tbl.columns)
                body.append(f"[Table: {rows}x{cols}]")
            elif shape.has_text_frame:
                text = shape.text_frame.text.strip()
                if text:
                    body.append(text)

        notes: str | None = None
        if sld.has_notes_slide:
            notes_text = sld.notes_slide.notes_text_frame.text.strip()
            if notes_text:
                notes = notes_text

        result_slides.append(
            Slide(
                number=i,
                title=title_text if title_text else None,
                body=body,
                notes=notes,
                images=images,
            )
        )

        if slide is not None:
            break
        if head is not None and len(result_slides) >= head:
            break

    return Presentation(
        source=path.name,
        slides=result_slides,
        total_slides=total_slides,
    )
