"""TUI view for Excel and CSV data."""

from __future__ import annotations

from textual.app import ComposeResult
from textual.containers import Vertical
from textual.widgets import DataTable, Label, Static, TabbedContent, TabPane

from officecat.models import CsvData, Sheet, Workbook


class SheetView(Vertical):
    """A single sheet rendered as a DataTable with a status bar."""

    def __init__(self, sheet: Sheet, **kwargs: object) -> None:
        super().__init__(**kwargs)
        self._sheet = sheet

    def compose(self) -> ComposeResult:
        table = DataTable(zebra_stripes=True)
        yield table
        if self._sheet.total_rows > len(self._sheet.rows):
            yield Label(
                f"  Showing {len(self._sheet.rows)} of {self._sheet.total_rows} rows. "
                f"Use --all to show everything.",
                classes="status-dim",
            )

    def on_mount(self) -> None:
        table = self.query_one(DataTable)
        for h in self._sheet.headers:
            table.add_column(h, key=h)
        col_count = len(self._sheet.headers)
        for i, row in enumerate(self._sheet.rows):
            padded = row + [""] * (col_count - len(row))
            table.add_row(*padded[:col_count], key=str(i))


class XlsxView(Vertical):
    """Multi-sheet workbook view with tabs."""

    def __init__(self, workbook: Workbook, **kwargs: object) -> None:
        super().__init__(**kwargs)
        self._workbook = workbook

    def compose(self) -> ComposeResult:
        if not self._workbook.sheets:
            yield Static("Workbook is empty.")
            return

        if len(self._workbook.sheets) == 1:
            yield SheetView(self._workbook.sheets[0])
        else:
            with TabbedContent():
                for sheet in self._workbook.sheets:
                    with TabPane(sheet.name):
                        yield SheetView(sheet)


class CsvView(Vertical):
    """CSV data view — single sheet."""

    def __init__(self, csv_data: CsvData, **kwargs: object) -> None:
        super().__init__(**kwargs)
        self._data = csv_data

    def compose(self) -> ComposeResult:
        if not self._data.rows and not self._data.headers:
            yield Static("File is empty.")
            return

        sheet = Sheet(
            name=self._data.source,
            headers=self._data.headers,
            rows=self._data.rows,
            total_rows=self._data.total_rows,
        )
        yield SheetView(sheet)
