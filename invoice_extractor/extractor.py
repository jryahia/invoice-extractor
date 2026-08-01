"""
PDF parsing and field extraction for Italian invoices.

Uses pdfplumber for text extraction and regex patterns for field matching.
Supports batch processing of multiple PDF files.

Example:
    >>> from pathlib import Path
    >>> from invoice_extractor import extract_invoice
    >>> inv = extract_invoice(Path("fattura.pdf"))   # doctest: +SKIP
    >>> inv.total, inv.vat_number                    # doctest: +SKIP
    (1250.0, '01234567890')

Failures are never raised out of :func:`extract_invoice`: an unreadable file
produces an :class:`ExtractedInvoice` carrying an Italian message in
``errors``, so a batch run never aborts halfway through.
"""

from __future__ import annotations

import locale
import logging
import re
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable, Dict, List, Optional

import pdfplumber

from .patterns import (
    DATE_FORMATS,
    FIELD_LABELS_IT,
    ITALIAN_MONTHS,
    PATTERNS,
)

logger = logging.getLogger("invoice_extractor")

#: Fields extracted from every invoice, in display order.
EXTRACTED_FIELDS = ["invoice_number", "date", "total", "supplier_name", "vat_number"]


# ── Errors ─────────────────────────────────────────────────────────────────


class InvoiceExtractionError(Exception):
    """Base class for unrecoverable problems with a single PDF."""


class EmptyFileError(InvoiceExtractionError):
    """The file is zero bytes (or too small to be a PDF)."""


class CorruptPDFError(InvoiceExtractionError):
    """The file is not a readable PDF (bad header, truncated, encrypted)."""


class ScannedPDFError(InvoiceExtractionError):
    """The PDF is valid but contains no text layer — most likely a scan."""


# ── Data model ─────────────────────────────────────────────────────────────


@dataclass
class ExtractedInvoice:
    """
    Data extracted from a single invoice PDF.

    Two distinct failure kinds are tracked separately:

    ``errors``
        Hard failures — the PDF could not be read at all. No field is usable.
    ``missing_fields``
        Soft issues — the PDF was read fine but a field did not match any
        pattern. The other fields are still valid.
    """

    filename: str
    invoice_number: Optional[str] = None
    date: Optional[str] = None
    total: Optional[float] = None
    supplier_name: Optional[str] = None
    vat_number: Optional[str] = None
    errors: List[str] = field(default_factory=list)
    missing_fields: List[str] = field(default_factory=list)

    # ── Status helpers ────────────────────────────────────────────────────

    @property
    def has_hard_error(self) -> bool:
        """True if the PDF could not be read (corrupt, empty, scanned)."""
        return bool(self.errors)

    @property
    def is_ok(self) -> bool:
        """True if the PDF was read and every field was extracted."""
        return not self.errors and not self.missing_fields

    @property
    def is_credit_note(self) -> bool:
        """True if the total is negative — i.e. a nota di credito."""
        return self.total is not None and self.total < 0

    @property
    def status_it(self) -> str:
        """Human-readable Italian status, used in the Excel/CSV export."""
        if self.has_hard_error:
            return "Errore"
        if self.missing_fields:
            return "Incompleto"
        return "OK"

    @property
    def document_type_it(self) -> str:
        """Italian document type inferred from the sign of the total."""
        if self.total is None:
            return ""
        return "Nota di credito" if self.total < 0 else "Fattura"

    @property
    def notes_it(self) -> str:
        """All errors and missing-field warnings joined into one cell."""
        parts = list(self.errors)
        if self.missing_fields:
            labels = [FIELD_LABELS_IT.get(f, f) for f in self.missing_fields]
            parts.append("Campi non trovati: " + ", ".join(labels))
        return " | ".join(parts)

    def to_dict(self) -> Dict[str, object]:
        """Serialize to an Italian-labelled dict (for Excel/CSV export)."""
        return {
            "File": self.filename,
            "Numero Fattura": self.invoice_number or "",
            "Data": self.date or "",
            "Totale": self.total,
            "Fornitore": self.supplier_name or "",
            "Partita IVA": self.vat_number or "",
            "Tipo Documento": self.document_type_it,
            "Esito": self.status_it,
            "Note": self.notes_it,
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
      - -1.234,56 and (1.234,56)  → negative (note di credito)

    Returns None if the text is not a recognisable number.
    """
    text = text.strip().replace("€", "").strip()
    text = text.replace(" ", "").replace(" ", "")

    # Note di credito mark negatives either with a sign or with parentheses.
    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1].strip()
    if text[:1] in ("-", "−"):
        negative = True
        text = text[1:].strip()
    elif text[:1] == "+":
        text = text[1:].strip()

    value: Optional[float] = None

    # Italian: 1.234,56
    if re.match(r"^\d{1,3}(\.\d{3})*(,\d{1,2})?$", text):
        value = float(text.replace(".", "").replace(",", "."))
    # Comma as decimal: 1234,56
    elif re.match(r"^\d+(,\d{1,2})?$", text):
        value = float(text.replace(",", "."))
    else:
        # Plain decimal: 1234.56
        try:
            value = float(text)
        except ValueError:
            return None

    return -value if negative else value


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
    """
    Extract all text from a PDF file using pdfplumber.

    Raises:
        FileNotFoundError: The path does not exist.
        EmptyFileError: The file is empty.
        CorruptPDFError: The file is not a readable/decryptable PDF.
        ScannedPDFError: The PDF parsed but has no text layer (scanned image).
    """
    if not pdf_path.exists():
        raise FileNotFoundError(f"File PDF non trovato: {pdf_path}")

    if pdf_path.stat().st_size == 0:
        raise EmptyFileError(f"Il file è vuoto (0 byte): {pdf_path.name}")

    try:
        with pdfplumber.open(str(pdf_path)) as pdf:
            if not pdf.pages:
                raise CorruptPDFError(f"Il PDF non contiene pagine: {pdf_path.name}")
            pages_text: List[str] = []
            for page in pdf.pages:
                try:
                    text = page.extract_text()
                except Exception as exc:  # a single damaged page must not kill the file
                    logger.warning(
                        "Pagina %s illeggibile in %s: %s", page.page_number, pdf_path.name, exc
                    )
                    continue
                if text:
                    pages_text.append(text)
    except InvoiceExtractionError:
        raise
    except Exception as exc:
        raise CorruptPDFError(
            f"PDF danneggiato o protetto da password: {pdf_path.name} ({exc})"
        ) from exc

    full_text = "\n".join(pages_text)
    if not full_text.strip():
        raise ScannedPDFError(
            f"Nessun testo estraibile da {pdf_path.name}: "
            "probabilmente è una scansione. Serve un OCR."
        )
    return full_text


# ── Field extraction ──────────────────────────────────────────────────────


def extract_field(text: str, field_name: str) -> Optional[str]:
    """
    Try all patterns for a given field name against the text.

    Returns the first non-None match value, cleaned of leading/trailing
    punctuation and whitespace. A leading minus sign is preserved when it
    belongs to a number (note di credito).
    """
    patterns = PATTERNS.get(field_name, [])
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if match:
            value = match.group(1).strip()
            # Clean leading/trailing punctuation, but keep a numeric sign.
            value = re.sub(r"^[\s.,;:]+|[\s.,;:]+$", "", value)
            value = re.sub(r"^-+(?!\s*[\d(])", "", value)
            value = re.sub(r"-+$", "", value).strip()
            if value:
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
        result["total"] = parse_italian_number(total_str)
    else:
        result["total"] = None

    return result


# ── Full pipeline ──────────────────────────────────────────────────────────


def extract_invoice(pdf_path: Path) -> ExtractedInvoice:
    """
    Extract all fields from a single PDF invoice file.

    Never raises: read failures are recorded in ``result.errors`` as an
    Italian message so batch processing can continue.
    """
    pdf_path = Path(pdf_path)
    result = ExtractedInvoice(filename=pdf_path.name)

    try:
        text = extract_text_from_pdf(pdf_path)
    except (InvoiceExtractionError, FileNotFoundError) as exc:
        logger.warning("Estrazione fallita per %s: %s", pdf_path.name, exc)
        result.errors.append(str(exc))
        return result
    except Exception as exc:  # unexpected — still must not abort the batch
        logger.exception("Errore imprevisto su %s", pdf_path.name)
        result.errors.append(f"Errore imprevisto durante la lettura: {exc}")
        return result

    data = extract_data_from_text(text)

    result.invoice_number = data.get("invoice_number")  # type: ignore[assignment]
    result.date = data.get("date")  # type: ignore[assignment]
    result.total = data.get("total")  # type: ignore[assignment]
    result.supplier_name = data.get("supplier_name")  # type: ignore[assignment]
    result.vat_number = data.get("vat_number")  # type: ignore[assignment]

    # Track which fields were not found (soft issues, not hard errors).
    result.missing_fields = [f for f in EXTRACTED_FIELDS if data.get(f) is None]
    if result.missing_fields:
        logger.debug(
            "Campi non trovati in %s: %s", pdf_path.name, ", ".join(result.missing_fields)
        )

    return result


def extract_batch(
    pdf_paths: List[Path],
    show_progress: bool = True,
    progress_callback: Optional[Callable[[int, int, Path], None]] = None,
) -> List[ExtractedInvoice]:
    """
    Extract data from a batch of PDF invoices.

    Args:
        pdf_paths: List of paths to PDF files.
        show_progress: Log one INFO line per processed file.
        progress_callback: Called as ``(index, total, path)`` before each file,
            so a caller can drive its own progress bar (the CLI uses Rich).

    Returns:
        List of :class:`ExtractedInvoice` instances, one per input path.
    """
    results: List[ExtractedInvoice] = []
    total = len(pdf_paths)

    logger.info("Avvio elaborazione di %d file PDF", total)

    for idx, pdf_path in enumerate(pdf_paths, 1):
        if progress_callback is not None:
            progress_callback(idx, total, pdf_path)
        if show_progress:
            logger.info("[%d/%d] Elaborazione: %s", idx, total, pdf_path.name)
        results.append(extract_invoice(pdf_path))

    ok = sum(1 for r in results if r.is_ok)
    logger.info("Elaborazione completata: %d/%d estratti senza problemi", ok, total)

    return results
