"""
Regex patterns for Italian invoice field extraction.

Supports multiple common formats for:
  - Fatture (invoices)
  - Ricevute fiscali (fiscal receipts)
  - Documenti di trasporto (DDT)
  - Note di credito/debito

Each pattern group is a list of regex strings tried in order.
"""

from typing import Dict, List

# ── Core invoice field patterns ────────────────────────────────────────────

PATTERNS: Dict[str, List[str]] = {
    "invoice_number": [
        # "Fattura N. 123/ABC"
        r"(?:fattura|ricevuta|nota\s*(?:di\s*)?(?:credito|debito)|documento|ddt)\s*[nN]?[.:]?\s*(\S[\w/\-.]{0,40}?)(?:\s|$|\\n)",
        # "Numero Fattura: INV-2024-001"
        r"(?:numero|n[.:])\s*(?:fattura|documento|ricevuta)\s*[.:]?\s*(\S[\w/\-.]{0,40}?)(?:\s|$|\\n)",
        # "N. FATTURA 12345"
        r"[nN][.:]?\s*FATTURA\s*([\w/\-.]{1,40})",
        # "Protocollo: 2024/001"
        r"protocollo[.:]?\s*(\S[\w/\-.]{0,40}?)(?:\s|$|\\n)",
    ],
    "date": [
        # "Data Fattura: 15/03/2024"
        r"(?:data|del)\s*(?:fattura|documento|ricevuta|emissione)\s*[.:]?\s*(\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4})",
        # "Data: 2024-03-15"
        r"data[.:]?\s*(\d{4}[/\-.]\d{1,2}[/\-.]\d{1,2})",
        # "15/03/2024" standalone
        r"(?<!\d)(\d{1,2}[/\-.]\d{1,2}[/\-.]\d{2,4})(?!\d)",
        # Italian long format: "15 marzo 2024"
        r"(\d{1,2})\s*(gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)\s*(\d{4})",
    ],
    "total": [
        # "Totale Fattura EUR 1.250,00" (with EUR/USD currency code)
        r"(?:totale|importo)\s*(?:fattura|documento)?\s*[.:]?\s*(?:€|€|eur|euro|usd)?\s*([\d]{1,3}(?:[. ]\d{3})*(?:,\d{2})?)",
        # "Totale Fattura € 1.234,56"
        r"(?:totale|importo\s*totale|totale\s*documento|totale\s*da\s*pagare|totale\s*fattura)\s*[.:]?\s*[€€]?\s*([\d]{1,3}(?:[. ]\d{3})*(?:,\d{2})?)",
        # "Importo: € 1.234,56"
        r"importo[.:]?\s*[€€]?\s*([\d]{1,3}(?:[. ]\d{3})*(?:,\d{2})?)",
        # "€ 1.234,56" or "EUR 1.234,56"
        r"(?:€|€|eur|euro)\s*([\d]{1,3}(?:[. ]\d{3})*(?:,\d{2})?)",
        # Plain number with 2 decimals near "totale" context
        r"(?:(?:totale|importo|tot\.?)\s*[.:]?\s*)(\d{1,3}(?:[. ]\d{3})*(?:,\d{2}))",
        # "Tot. € 1234.56" — dot as decimal
        r"(?:totale|importo)\s*(?:€|€|eur)?\s*([\d]+(?:,\d{1,2}|\.\d{2}))",
    ],
    "supplier_name": [
        # "Fornitore: Nome Azienda S.r.l."
        r"(?:fornitore|mittente|emittente|dal|prestatore)\s*[.:]?\s*(.+?)(?:\\n|\||via\s|p[.]?\s*[.]?\s*iva|cod\s*[.]?\s*fisc|indirizzo|telefono|email|$)",
        # "Spett.le Nome Azienda" (formal heading)
        r"spett[.]?\s*le\s*(.+?)(?:\\n|\||via\s|c.a\.|p[.]?\s*iva)",
        # "Ragione Sociale: Nome Azienda"
        r"(?:ragione\s*sociale|denominazione|ditta|societ[àa])\s*[.:]?\s*(.+?)(?:\\n|\||via|p[.]?\s*iva|cod\s*fisc|tel)",
        # First line with legal form indicators
        r"^(.+?(?:S\.r\.l\.|S\.p\.A\.|S\.n\.c\.|S\.a\.s\.|Societ[àa]|Ditta))",
    ],
    "vat_number": [
        # "P.IVA: 01234567890"
        r"(?:p[.]?\s*[.]?\s*iva|partita\s*iva|p\s*iva)\s*[.:]?\s*([\d]{11})",
        # "Partita IVA 01234567890"
        r"([\d]{11})(?:\s*$|\s*\n|\s*\||\s*-)",
        # "Cod. Fisc. 01234567890" (also captures VAT if 11 digits)
        r"(?:cod\s*[.]?\s*fisc|codice\s*fiscale)\s*[.:]?\s*([\d]{11})",
    ],
}

# ── Date format strings for parsing ────────────────────────────────────────

DATE_FORMATS: List[str] = [
    "%d/%m/%Y",
    "%d/%m/%y",
    "%d-%m-%Y",
    "%d-%m-%y",
    "%d.%m.%Y",
    "%d.%m.%y",
    "%Y/%m/%d",
    "%Y-%m-%d",
    "%Y.%m.%d",
    "%d %B %Y",  # Italian month names via locale
    "%d %b %Y",
]

# ── Month name mapping (Italian → numeric) ────────────────────────────────

ITALIAN_MONTHS: Dict[str, str] = {
    "gennaio": "01",
    "febbraio": "02",
    "marzo": "03",
    "aprile": "04",
    "maggio": "05",
    "giugno": "06",
    "luglio": "07",
    "agosto": "08",
    "settembre": "09",
    "ottobre": "10",
    "novembre": "11",
    "dicembre": "12",
}
