"""
Invoice Extractor — B2B PDF Invoice → Excel tool.

Extract structured data (invoice number, date, total, supplier, VAT)
from Italian PDF invoices and export to formatted Excel files.
"""

from .extractor import ExtractedInvoice, extract_batch, extract_invoice
from .exporter import export_to_excel

__version__ = "0.1.0"
__all__ = [
    "ExtractedInvoice",
    "extract_invoice",
    "extract_batch",
    "export_to_excel",
]
