"""
PDF parsing and field extraction for Italian invoices.

Uses pdfplumber for text extraction and regex patterns for field matching.
Supports batch processing of multiple PDF files.
"""

from __future__ import annotations

import locale
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Dict, List, Optional, Set

import pdfplumber

from .patterns import (
    DATE_FORMATS,
    ITALIAN_MONTHS,
    PATTERNS,
)


# ── Data model ─────────────────────────────────────────────────────────────


@dataclass
class ExtractedInvoice:
    """Represents data extracted from a single invoice PDF."""

    filename: str
    invoice_number: Optional[str] = None
    date: Optional[str] = None
    total: Optional[float] = None
    supplier_name: Optional[str] = None
    vat_number: Optional[str] = None
    errors: List[str] = field(default_factory=list)

    def to_dict(self) -> Dict[str, object]:
        """Serialize to dict (for Excel export)."""
        return {
            "Filename": self.filename,
            "Invoice Number": self.invoice_number or "",
            "Date": self.date or "",
            "Total": self.total,
            "Supplier Name": self.supplier_name or "",
            "VAT Number": self.vat_number or "",
            "Errors": "; ".join(self.errors) if self.errors else "",
        }


# ── Italian locale setup ───────────────────────────────────────────────────


def _try_set_italian_locale() -> None:
    """Try setting Italian locale for date parsing with month names."""
    for loc in ["it_IT.UTF-8", "it_IT.utf8", "it_IT"]:
        try:
            locale.setlocale(locale.LC_TIME, loc)
            return
        except locale.Error:
            continue


_try_set_italian_locale()


# ── Number parsing helpers ────────────────────────────────────────────────


def parse_italian_number(text: str) -> Optional[float]:
    """
    Parse an Italian-formatted number (e.g. ``1.234,56`` → 1234.56).

    Handles:
      - 1.234,56   (dot = thousand sep, comma = decimal)
      - 1 234,56   (space = thousand sep)
      - 1234.56    (plain decimal)
      - 1234,56    (comma decimal)
    """
    text = text.strip().replace("€", "").strip()
    text = text.replace(" ", "")

    # Italian: 1.234,56
    if re.match(r"^\d{1,3}(\.\d{3})*(,\d{1,2})?$", text):
        text = text.replace(".", "").replace(",", ".")
        return float(text)

    # Comma as decimal: 1234,56
    if re.match(r"^\d+(,\d{1,2})?$", text):
        text = text.replace(",", ".")
        return float(text)

    # Plain decimal: 1234.56
    try:
        return float(text)
    except ValueError:
        return None


def parse_italian_date(text: str) -> Optional[str]:
    """Parse a date string from invoice and return YYYY-MM-DD string."""
    text = text.strip()

    # Try Italian long format: "15 marzo 2024"
    match = re.match(
        r"(\d{1,2})\s*(gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)\s*(\d{4})",
        text,
        re.IGNORECASE,
    )
    if match:
        day, month_name, year = match.groups()
        month = ITALIAN_MONTHS.get(month_name.lower())
        if month:
            return f"{year}-{month}-{int(day):02d}"

    # Try all configured date formats
    for fmt in DATE_FORMATS:
        try:
            dt = datetime.strptime(text, fmt)
            return dt.strftime("%Y-%m-%d")
        except ValueError:
            continue

    return None


# ── Text extraction ────────────────────────────────────────────────────────


def extract_text_from_pdf(pdf_path: Path) -> str:
    """Extract all text from a PDF file using pdfplumber."""
    if not pdf_path.exists():
        raise FileNotFoundError(f"PDF not found: {pdf_path}")

    with pdfplumber.open(str(pdf_path)) as pdf:
        pages_text: List[str] = []
        for page in pdf.pages:
            text = page.extract_text()
            if text:
                pages_text.append(text)
        return "\n".join(pages_text)


# ── Field extraction ──────────────────────────────────────────────────────


def extract_field(text: str, field_name: str) -> Optional[str]:
    """
    Try all patterns for a given field name against the text.

    Returns the first non-None match value, cleaned of leading/trailing
    punctuation and whitespace.
    """
    patterns = PATTERNS.get(field_name, [])
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if match:
            value = match.group(1).strip()
            # Clean leading/trailing punctuation
            value = re.sub(r'^[\s.,;:\-]+|[\s.,;:\-]+$', '', value)
            if value and len(value) >= 1:
                return value
    return None


def extract_data_from_text(text: str) -> Dict[str, object]:
    """
    Run all field extractions on the full document text.
    Returns a dict with keys: invoice_number, date, total, supplier_name, vat_number.
    """
    result: Dict[str, object] = {}

    # String fields
    result["invoice_number"] = extract_field(text, "invoice_number")
    result["supplier_name"] = extract_field(text, "supplier_name")
    result["vat_number"] = extract_field(text, "vat_number")

    # Date — try extraction then parse
    date_str = extract_field(text, "date")
    if date_str:
        parsed = parse_italian_date(date_str)
        result["date"] = parsed or date_str  # fallback to raw string
    else:
        result["date"] = None

    # Total — try extraction then parse number
    total_str = extract_field(text, "total")
    if total_str:
        total_value = parse_italian_number(total_str)
        result["total"] = total_value
    else:
        result["total"] = None

    return result


# ── Full pipeline ──────────────────────────────────────────────────────────


def extract_invoice(pdf_path: Path) -> ExtractedInvoice:
    """Extract all fields from a single PDF invoice file."""
    result = ExtractedInvoice(filename=pdf_path.name)

    try:
        text = extract_text_from_pdf(pdf_path)
    except Exception as exc:
        result.errors.append(f"PDF read error: {exc}")
        return result

    if not text.strip():
        result.errors.append("No text extracted from PDF")
        return result

    data = extract_data_from_text(text)

    result.invoice_number = data.get("invoice_number")  # type: ignore
    result.date = data.get("date")  # type: ignore
    result.total = data.get("total")  # type: ignore
    result.supplier_name = data.get("supplier_name")  # type: ignore
    result.vat_number = data.get("vat_number")  # type: ignore

    # Track which fields were not found
    for key in ["invoice_number", "date", "total", "supplier_name", "vat_number"]:
        if data.get(key) is None:
            result.errors.append(f"Could not extract '{key}'")

    return result


def extract_batch(
    pdf_paths: List[Path],
    show_progress: bool = True,
) -> List[ExtractedInvoice]:
    """
    Extract data from a batch of PDF invoices.

    Args:
        pdf_paths: List of paths to PDF files.
        show_progress: Print progress to stderr (for CLI integration).

    Returns:
        List of ExtractedInvoice dataclass instances.
    """
    results: List[ExtractedInvoice] = []
    total = len(pdf_paths)

    for idx, pdf_path in enumerate(pdf_paths, 1):
        if show_progress:
            print(f"  [{idx}/{total}] Processing: {pdf_path.name}")
        result = extract_invoice(pdf_path)
        results.append(result)

    return results
