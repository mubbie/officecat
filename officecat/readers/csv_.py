"""CSV/TSV reader."""

from __future__ import annotations

import csv
from pathlib import Path

from officecat.models import CsvData


def read_csv(
    path: Path,
    *,
    head: int | None = None,
    headers: int = 1,
) -> CsvData:
    is_tsv = path.suffix.lower() == ".tsv"

    with open(path, newline="", encoding="utf-8-sig") as f:
        if is_tsv:
            dialect = None
            delimiter = "\t"
        else:
            sample = f.read(8192)
            try:
                dialect = csv.Sniffer().sniff(sample)
                delimiter = dialect.delimiter
            except csv.Error:
                delimiter = ","
            f.seek(0)

        reader = csv.reader(f, delimiter=delimiter)

        header_row: list[str] = []
        rows: list[list[str]] = []
        total_rows = 0

        for i, row in enumerate(reader, start=1):
            if headers > 0 and i == headers:
                header_row = row
                continue

            total_rows += 1
            if head is None or len(rows) < head:
                rows.append(row)

        if not header_row and rows:
            col_count = max(len(r) for r in rows)
            header_row = [chr(65 + i) if i < 26 else f"Col{i+1}" for i in range(col_count)]

    return CsvData(
        source=path.name,
        headers=header_row,
        rows=rows,
        total_rows=total_rows,
    )
