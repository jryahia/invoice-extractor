"""
Regex patterns for Italian invoice field extraction.

Supports multiple common formats for:
  - Fatture (invoices)
  - Ricevute fiscali (fiscal receipts)
  - Documenti di trasporto (DDT)
  - Note di credito/debito

Each pattern group is a list of regex strings tried in order.
"""

from __future__ import annotations

import json
import logging
import re
from pathlib import Path
from typing import Dict, List

logger = logging.getLogger("invoice_extractor")

# ── Shared regex fragments ─────────────────────────────────────────────────

# An Italian-formatted amount, optionally signed or wrapped in parentheses
# (both are used on note di credito to mark a negative total).
_AMOUNT = r"\(?\s*[-−]?\s*\d{1,3}(?:[. ]\d{3})*(?:,\d{1,2})?\s*\)?"

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
        # "Totale Fattura EUR 1.250,00" / "Totale Fattura EUR -450,00" (credit note)
        rf"(?:totale|importo)\s*(?:fattura|documento)?\s*[.:]?\s*(?:€|€|eur|euro|usd)?\s*({_AMOUNT})",
        # "Totale Fattura € 1.234,56"
        rf"(?:totale|importo\s*totale|totale\s*documento|totale\s*da\s*pagare|totale\s*fattura)\s*[.:]?\s*[€€]?\s*({_AMOUNT})",
        # "Importo: € 1.234,56"
        rf"importo[.:]?\s*[€€]?\s*({_AMOUNT})",
        # "€ 1.234,56" or "EUR 1.234,56"
        rf"(?:€|€|eur|euro)\s*({_AMOUNT})",
        # Plain number with 2 decimals near "totale" context
        r"(?:(?:totale|importo|tot\.?)\s*[.:]?\s*)(\(?\s*[-−]?\s*\d{1,3}(?:[. ]\d{3})*,\d{2}\s*\)?)",
        # "Tot. € 1234.56" — dot as decimal
        r"(?:totale|importo)\s*(?:€|€|eur)?\s*([-−]?\d+(?:,\d{1,2}|\.\d{2}))",
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

# ── Italian labels for the extracted fields ───────────────────────────────

FIELD_LABELS_IT: Dict[str, str] = {
    "invoice_number": "numero fattura",
    "date": "data",
    "total": "totale",
    "supplier_name": "fornitore",
    "vat_number": "partita IVA",
}


# ── Custom pattern loading (--config) ─────────────────────────────────────


def load_patterns_from_config(config_path: Path) -> Dict[str, List[str]]:
    """
    Merge custom regex patterns from a ``config.json`` into :data:`PATTERNS`.

    The config file may contain a ``"patterns"`` object mapping a field name to
    a list of regex strings, and a ``"date_formats"`` list. Custom patterns are
    tried *before* the built-in ones, so a client can override the defaults
    without losing them as a fallback.

    Args:
        config_path: Path to the JSON configuration file.

    Returns:
        The updated :data:`PATTERNS` dictionary.

    Raises:
        ValueError: If the file is not valid JSON or a pattern is not a valid
            regular expression. The message is in Italian (user-facing).

    Example:
        >>> load_patterns_from_config(Path("config.json"))  # doctest: +SKIP
    """
    try:
        raw = json.loads(config_path.read_text(encoding="utf-8"))
    except json.JSONDecodeError as exc:
        raise ValueError(
            f"File di configurazione non valido ({config_path.name}): {exc}"
        ) from exc
    except OSError as exc:
        raise ValueError(
            f"Impossibile leggere il file di configurazione {config_path}: {exc}"
        ) from exc

    custom = raw.get("patterns", {})
    if not isinstance(custom, dict):
        raise ValueError("La chiave 'patterns' deve essere un oggetto JSON.")

    for field_name, regexes in custom.items():
        if field_name not in PATTERNS:
            logger.warning("Campo sconosciuto nella configurazione: '%s' (ignorato)", field_name)
            continue
        if isinstance(regexes, str):
            regexes = [regexes]
        valid: List[str] = []
        for rx in regexes:
            try:
                re.compile(rx)
            except re.error as exc:
                raise ValueError(
                    f"Espressione regolare non valida per '{field_name}': {rx} ({exc})"
                ) from exc
            valid.append(rx)
        # Custom patterns take priority, built-ins stay as fallback.
        PATTERNS[field_name] = valid + [p for p in PATTERNS[field_name] if p not in valid]
        logger.info("Caricati %d pattern personalizzati per '%s'", len(valid), field_name)

    for fmt in raw.get("date_formats", []):
        if fmt not in DATE_FORMATS:
            DATE_FORMATS.insert(0, fmt)

    return PATTERNS
