"""
Invoice Extractor — B2B PDF Invoice → Excel tool.

Extract structured data (invoice number, date, total, supplier, VAT)
from Italian PDF invoices and export to formatted Excel files.
"""

from .extractor import (
    CorruptPDFError,
    EmptyFileError,
    ExtractedInvoice,
    InvoiceExtractionError,
    ScannedPDFError,
    extract_batch,
    extract_invoice,
)
from .exporter import export_to_csv, export_to_excel

__version__ = "1.0.0"
__all__ = [
    "ExtractedInvoice",
    "extract_invoice",
    "extract_batch",
    "export_to_excel",
    "export_to_csv",
    "InvoiceExtractionError",
    "CorruptPDFError",
    "EmptyFileError",
    "ScannedPDFError",
    "__version__",
]
