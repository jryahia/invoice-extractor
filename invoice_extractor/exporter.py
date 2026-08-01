"""
Excel and CSV export with formatting, auto-column-width, and totals row.

Produces a clean .xlsx workbook with:
  - Bold headers with background fill
  - Auto-filter on the header row and frozen panes
  - Zebra striping on data rows for readability
  - Auto-adjusted column widths
  - Currency formatting for the Totale column (negatives in red)
  - A summary sheet with totals and extraction statistics

The CSV export uses Italian conventions (``;`` separator, comma decimal,
UTF-8 BOM) so the file opens correctly in an Italian Excel by double-click.

Example:
    >>> from invoice_extractor import export_to_excel, export_to_csv
    >>> export_to_excel(invoices, "output/fatture.xlsx")   # doctest: +SKIP
    >>> export_to_csv(invoices, "output/fatture.csv")      # doctest: +SKIP
"""

from __future__ import annotations

import csv
import logging
from datetime import datetime
from pathlib import Path
from typing import List

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.worksheet import Worksheet

from .extractor import ExtractedInvoice

logger = logging.getLogger("invoice_extractor")


# ── Style constants ────────────────────────────────────────────────────────

HEADER_FILL = PatternFill(start_color="1F4E3D", end_color="1F4E3D", fill_type="solid")
HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
ZEBRA_FILL = PatternFill(start_color="F2F7F4", end_color="F2F7F4", fill_type="solid")
TOTAL_ROW_FILL = PatternFill(start_color="D5E8DC", end_color="D5E8DC", fill_type="solid")
TOTAL_ROW_FONT = Font(bold=True, size=11)
ERROR_FONT = Font(color="B00020")
WARNING_FONT = Font(color="B26A00")

#: Currency format — negatives (note di credito) shown in red with a sign.
CURRENCY_FMT = '#,##0.00" €";[Red]-#,##0.00" €"'
DATE_FMT = "YYYY-MM-DD"

#: (header, width) — headers must match the keys of ``ExtractedInvoice.to_dict``.
COLUMNS = [
    ("File", 38),
    ("Numero Fattura", 20),
    ("Data", 13),
    ("Totale", 15),
    ("Fornitore", 32),
    ("Partita IVA", 16),
    ("Tipo Documento", 17),
    ("Esito", 12),
    ("Note", 46),
]

#: 1-based index of the "Totale" column, used for the grand-total cell.
TOTAL_COL_IDX = [name for name, _ in COLUMNS].index("Totale") + 1
STATUS_COL_IDX = [name for name, _ in COLUMNS].index("Esito") + 1


# ── Helpers ────────────────────────────────────────────────────────────────


def _style_header(ws: Worksheet, num_cols: int) -> None:
    """Apply header styling to the first row."""
    for col_idx in range(1, num_cols + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 22


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
    """Write one invoice's data to a row, with zebra striping and status colour."""
    data = invoice.to_dict()
    col_keys = [name for name, _ in COLUMNS]
    striped = row_idx % 2 == 1  # alternate, starting from the 2nd data row

    for col_idx, key in enumerate(col_keys, 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        value = data.get(key)

        if key == "Totale":
            # 0.0 is a legitimate total, so test for None explicitly.
            cell.value = invoice.total if invoice.total is not None else ""
            cell.number_format = CURRENCY_FMT
            cell.alignment = Alignment(vertical="top", horizontal="right")
        else:
            cell.value = value if value else ""
            cell.alignment = Alignment(vertical="top")

        if key == "Esito":
            if invoice.has_hard_error:
                cell.font = ERROR_FONT
            elif invoice.missing_fields:
                cell.font = WARNING_FONT

        if striped:
            cell.fill = ZEBRA_FILL


# ── Summary sheet ──────────────────────────────────────────────────────────


def _build_summary(
    wb: Workbook,
    invoices: List[ExtractedInvoice],
    output_path: Path,
) -> None:
    """Create a Riepilogo (summary) sheet with statistics and totals."""
    ws = wb.create_sheet(title="Riepilogo", index=0)

    total_invoices = len(invoices)
    totals = [inv.total for inv in invoices if inv.total is not None]
    grand_total = sum(totals)
    successful = sum(1 for inv in invoices if inv.is_ok)
    incomplete = sum(1 for inv in invoices if not inv.has_hard_error and inv.missing_fields)
    failed = sum(1 for inv in invoices if inv.has_hard_error)
    credit_notes = sum(1 for inv in invoices if inv.is_credit_note)

    summary_data = [
        ("Voce", "Valore"),
        ("PDF elaborati", total_invoices),
        ("Estratti correttamente", successful),
        ("Estratti parzialmente", incomplete),
        ("Non leggibili", failed),
        ("Note di credito", credit_notes),
        ("", ""),
        ("Totale complessivo (€)", grand_total),
        ("", ""),
        ("File esportato in", str(output_path.resolve())),
        ("Data esportazione", datetime.now().strftime("%d/%m/%Y %H:%M:%S")),
    ]

    for row_idx, (label, value) in enumerate(summary_data, 1):
        cell_a = ws.cell(row=row_idx, column=1, value=label)
        cell_b = ws.cell(row=row_idx, column=2, value=value)

        if row_idx == 1:
            cell_a.font = HEADER_FONT
            cell_a.fill = HEADER_FILL
            cell_b.font = HEADER_FONT
            cell_b.fill = HEADER_FILL
        elif label == "Totale complessivo (€)":
            cell_a.font = TOTAL_ROW_FONT
            cell_b.font = TOTAL_ROW_FONT
            cell_b.number_format = CURRENCY_FMT
            cell_b.fill = TOTAL_ROW_FILL
            cell_a.fill = TOTAL_ROW_FILL
        elif label == "Non leggibili" and failed:
            cell_b.font = ERROR_FONT

    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 50
    ws.freeze_panes = "A2"


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
        include_summary: Whether to add a Riepilogo sheet (default: True).

    Returns:
        Absolute path to the created Excel file.

    Raises:
        OSError: If the destination cannot be written (e.g. the file is open
            in Excel). The caller is expected to show an Italian message.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Fatture"

    num_cols = len(COLUMNS)

    # ── Header row ──────────────────────────────────────────────────────
    for col_idx, (name, _) in enumerate(COLUMNS, 1):
        ws.cell(row=1, column=col_idx, value=name)

    _style_header(ws, num_cols)
    _auto_width(ws, COLUMNS)

    # ── Data rows ────────────────────────────────────────────────────────
    for row_idx, invoice in enumerate(invoices, 2):
        _write_data_row(ws, row_idx, invoice, num_cols)

    last_data_row = len(invoices) + 1

    # ── Auto-filter — header + data only, never the grand-total row ─────
    ws.auto_filter.ref = f"A1:{get_column_letter(num_cols)}{max(last_data_row, 1)}"

    # ── Grand total row (left out of the filter range on purpose) ───────
    total_row = last_data_row + 2
    totals = [inv.total for inv in invoices if inv.total is not None]
    grand_total = sum(totals)

    label_cell = ws.cell(row=total_row, column=1, value="TOTALE COMPLESSIVO")
    label_cell.font = TOTAL_ROW_FONT
    label_cell.fill = TOTAL_ROW_FILL

    for col_idx in range(2, num_cols + 1):
        ws.cell(row=total_row, column=col_idx).fill = TOTAL_ROW_FILL

    total_cell = ws.cell(row=total_row, column=TOTAL_COL_IDX)
    total_cell.value = grand_total
    total_cell.font = TOTAL_ROW_FONT
    total_cell.number_format = CURRENCY_FMT
    total_cell.alignment = Alignment(horizontal="right")

    # ── Freeze the header row ────────────────────────────────────────────
    ws.freeze_panes = "A2"

    # ── Summary sheet ────────────────────────────────────────────────────
    if include_summary:
        _build_summary(wb, invoices, output_path)

    # ── Save ─────────────────────────────────────────────────────────────
    try:
        wb.save(str(output_path))
    except OSError as exc:
        logger.error("Salvataggio fallito su %s: %s", output_path, exc)
        raise OSError(
            f"Impossibile salvare '{output_path}'. "
            "Verifica che il file non sia già aperto in Excel e di avere i permessi di scrittura."
        ) from exc

    logger.info("Esportate %d fatture in %s", len(invoices), output_path)
    return output_path.resolve()


def export_to_csv(
    invoices: List[ExtractedInvoice],
    output_path: str | Path,
    delimiter: str = ";",
) -> Path:
    """
    Export extracted invoice data to a CSV file with Italian conventions.

    Uses ``;`` as separator and a comma decimal separator, written as UTF-8
    with BOM so an Italian Excel opens it correctly on double-click.

    Args:
        invoices: List of extracted invoice data.
        output_path: Where to save the .csv file.
        delimiter: Field separator (default ``;``, the Italian Excel default).

    Returns:
        Absolute path to the created CSV file.

    Raises:
        OSError: If the destination cannot be written.
    """
    output_path = Path(output_path)
    output_path.parent.mkdir(parents=True, exist_ok=True)

    headers = [name for name, _ in COLUMNS]

    try:
        with output_path.open("w", encoding="utf-8-sig", newline="") as fh:
            writer = csv.writer(fh, delimiter=delimiter, quoting=csv.QUOTE_MINIMAL)
            writer.writerow(headers)

            for inv in invoices:
                data = inv.to_dict()
                row = []
                for key in headers:
                    if key == "Totale":
                        # Italian decimal comma; empty when no total was found.
                        row.append("" if inv.total is None else f"{inv.total:.2f}".replace(".", ","))
                    else:
                        row.append(str(data.get(key) or ""))
                writer.writerow(row)

            totals = [inv.total for inv in invoices if inv.total is not None]
            grand_total = sum(totals)
            total_row = [""] * len(headers)
            total_row[0] = "TOTALE COMPLESSIVO"
            total_row[TOTAL_COL_IDX - 1] = f"{grand_total:.2f}".replace(".", ",")
            writer.writerow([])
            writer.writerow(total_row)
    except OSError as exc:
        logger.error("Salvataggio CSV fallito su %s: %s", output_path, exc)
        raise OSError(
            f"Impossibile salvare '{output_path}'. "
            "Verifica che il file non sia già aperto e di avere i permessi di scrittura."
        ) from exc

    logger.info("Esportate %d fatture in %s", len(invoices), output_path)
    return output_path.resolve()
