"""
Excel export with formatting, auto-column-width, and totals row.

Produces a clean .xlsx workbook with:
  - Bold headers with background fill
  - Auto-adjusted column widths
  - Currency formatting for the Total column
  - A summary sheet with totals and extraction statistics
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Optional

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, numbers
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .extractor import ExtractedInvoice


# ── Style constants ────────────────────────────────────────────────────────

HEADER_FILL = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
TOTAL_ROW_FILL = PatternFill(start_color="DCE6F1", end_color="DCE6F1", fill_type="solid")
TOTAL_ROW_FONT = Font(bold=True, size=11)
CURRENCY_FMT = '#,##0.00"€"'
DATE_FMT = "YYYY-MM-DD"

COLUMNS = [
    ("Filename", 40),
    ("Invoice Number", 22),
    ("Date", 14),
    ("Total", 14),
    ("Supplier Name", 35),
    ("VAT Number", 16),
    ("Errors", 30),
]


# ── Helpers ────────────────────────────────────────────────────────────────


def _style_header(ws: Worksheet, num_cols: int) -> None:
    """Apply header styling to the first row."""
    for col_idx in range(1, num_cols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center")


def _auto_width(ws: Worksheet, widths: List[tuple]) -> None:
    """Set column widths from predefined list."""
    for col_idx, (_, width) in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(col_idx)].width = width


def _write_data_row(
    ws: Worksheet,
    row_idx: int,
    invoice: ExtractedInvoice,
    num_cols: int,
) -> None:
    """Write one invoice's data to a row."""
    data = invoice.to_dict()
    col_keys = [name for name, _ in COLUMNS]

    for col_idx, key in enumerate(col_keys, 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        value = data.get(key)

        if key == "Total" and invoice.total is not None:
            cell.value = invoice.total
            cell.number_format = CURRENCY_FMT
        elif key == "Date" and invoice.date:
            cell.value = invoice.date
        else:
            cell.value = value if value else ""

        cell.alignment = Alignment(vertical="top")


# ── Summary sheet ──────────────────────────────────────────────────────────


def _build_summary(
    wb: Workbook,
    invoices: List[ExtractedInvoice],
    output_path: Path,
) -> None:
    """Create a Summary sheet with statistics and totals."""
    ws = wb.create_sheet(title="Summary", index=0)

    total_invoices = len(invoices)
    totals = [inv.total for inv in invoices if inv.total is not None]
    grand_total = sum(totals)
    successful = sum(1 for inv in invoices if not inv.errors or all(
        "Could not extract" not in e for e in inv.errors
    ))
    errors = total_invoices - successful

    summary_data = [
        ("Metric", "Value"),
        ("Total PDFs processed", total_invoices),
        ("Successfully extracted", successful),
        ("With errors", errors),
        ("", ""),
        ("Grand Total (€)", grand_total),
        ("", ""),
        ("File exported to", str(output_path.resolve())),
        ("Export timestamp", __import__("datetime").datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
    ]

    for row_idx, (label, value) in enumerate(summary_data, 1):
        cell_a = ws.cell(row=row_idx, column=1, value=label)
        cell_b = ws.cell(row=row_idx, column=2, value=value)

        if row_idx == 1:
            cell_a.font = HEADER_FONT
            cell_a.fill = HEADER_FILL
            cell_b.font = HEADER_FONT
            cell_b.fill = HEADER_FILL
        elif label == "Grand Total (€)":
            cell_a.font = TOTAL_ROW_FONT
            cell_b.font = TOTAL_ROW_FONT
            cell_b.number_format = CURRENCY_FMT
            cell_b.fill = TOTAL_ROW_FILL
            cell_a.fill = TOTAL_ROW_FILL

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 50


# ── Main export function ───────────────────────────────────────────────────


def export_to_excel(
    invoices: List[ExtractedInvoice],
    output_path: str | Path,
    include_summary: bool = True,
) -> Path:
    """
    Export extracted invoice data to a formatted .xlsx file.

    Args:
        invoices: List of extracted invoice data.
        output_path: Where to save the .xlsx file.
        include_summary: Whether to add a Summary sheet (default: True).

    Returns:
        Path to the created Excel file.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Invoices"

    num_cols = len(COLUMNS)

    # ── Header row ──────────────────────────────────────────────────────
    for col_idx, (name, _) in enumerate(COLUMNS, 1):
        ws.cell(row=1, column=col_idx, value=name)

    _style_header(ws, num_cols)
    _auto_width(ws, COLUMNS)

    # ── Data rows ────────────────────────────────────────────────────────
    for row_idx, invoice in enumerate(invoices, 2):
        _write_data_row(ws, row_idx, invoice, num_cols)

    # ── Grand total row ──────────────────────────────────────────────────
    total_row = len(invoices) + 2
    totals = [inv.total for inv in invoices if inv.total is not None]
    grand_total = sum(totals)

    ws.cell(row=total_row, column=1, value="GRAND TOTAL").font = TOTAL_ROW_FONT
    ws.cell(row=total_row, column=1).fill = TOTAL_ROW_FILL

    total_cell = ws.cell(row=total_row, column=4)
    total_cell.value = grand_total
    total_cell.font = TOTAL_ROW_FONT
    total_cell.fill = TOTAL_ROW_FILL
    total_cell.number_format = CURRENCY_FMT

    # ── Summary sheet ────────────────────────────────────────────────────
    if include_summary:
        _build_summary(wb, invoices, output_path)

    # ── Freeze top row ───────────────────────────────────────────────────
    ws.freeze_panes = "A2"

    # ── Save ─────────────────────────────────────────────────────────────
    wb.save(str(output_path))
    return output_path.resolve()
