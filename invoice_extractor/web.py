"""
Optional Gradio web UI for the invoice extractor.

Drag-and-drop interface for uploading PDF invoices, viewing the extracted
data in a preview table, and downloading the Excel/CSV export.

Gradio is an *optional* dependency. Importing this module without it raises
nothing at import time — :func:`main` prints an Italian install hint instead.

Launch with:
    python -m invoice_extractor.web
"""

from __future__ import annotations

import logging
import shutil
import tempfile
import time
from pathlib import Path
from typing import List, Optional, Sequence

from .exporter import COLUMNS, export_to_csv, export_to_excel
from .extractor import extract_batch

logger = logging.getLogger("invoice_extractor")

# ── Optional dependency guard ──────────────────────────────────────────────

try:  # pragma: no cover - depends on the environment
    import gradio as gr

    GRADIO_AVAILABLE = True
    GRADIO_IMPORT_ERROR: Optional[BaseException] = None
except Exception as exc:  # ImportError, or a broken/partial gradio install
    gr = None  # type: ignore[assignment]
    GRADIO_AVAILABLE = False
    GRADIO_IMPORT_ERROR = exc


INSTALL_HINT = """
╔══════════════════════════════════════════════════════════════════╗
║  Interfaccia web non disponibile                                 ║
╚══════════════════════════════════════════════════════════════════╝

Il pacchetto 'gradio' non è installato (o non è installato
correttamente). Per usare l'interfaccia web installalo con:

    pip install "gradio>=4.0.0"

In alternativa puoi usare l'interfaccia da riga di comando, che
non richiede gradio:

    invoice-extract ./invoices/ --output risultati.xlsx
"""


# ── Branding ───────────────────────────────────────────────────────────────

BG = "#0A0F0A"
EMERALD = "#2D7D46"
ORANGE = "#E8803A"

TITLE = "Estrattore Fatture B2B"
TAGLINE = "Da PDF a Excel in pochi secondi — per le piccole imprese italiane"

#: Inline SVG cross mark, drawn in the brand emerald with an orange accent.
LOGO_SVG = f"""
<svg width="46" height="46" viewBox="0 0 48 48" fill="none" xmlns="http://www.w3.org/2000/svg">
  <rect x="1" y="1" width="46" height="46" rx="12" fill="{EMERALD}" fill-opacity="0.14"
        stroke="{EMERALD}" stroke-width="1.5"/>
  <path d="M20 11h8v9h9v8h-9v9h-8v-9h-9v-8h9z" fill="{EMERALD}"/>
  <circle cx="36" cy="12" r="4" fill="{ORANGE}"/>
</svg>
"""

CUSTOM_CSS = f"""
/* ── Estrattore Fatture — tema scuro personalizzato ───────────────────── */

:root, .dark {{
  --ie-bg: {BG};
  --ie-surface: #111A13;
  --ie-surface-2: #16211A;
  --ie-border: #24352A;
  --ie-emerald: {EMERALD};
  --ie-emerald-soft: #3E9C5C;
  --ie-orange: {ORANGE};
  --ie-text: #E6EFE8;
  --ie-muted: #8FA396;
}}

.gradio-container, body, gradio-app {{
  background: var(--ie-bg) !important;
  color: var(--ie-text) !important;
  font-family: "Inter", "Segoe UI", system-ui, -apple-system, sans-serif !important;
  max-width: 1180px !important;
  margin: 0 auto !important;
}}

footer, .built-with, .show-api {{ display: none !important; }}

/* ── Intestazione con logo ─────────────────────────────────────────────── */
#ie-header {{
  display: flex; align-items: center; gap: 16px;
  padding: 26px 28px; margin-bottom: 22px;
  background: linear-gradient(135deg, #101A13 0%, #0C130E 100%);
  border: 1px solid var(--ie-border);
  border-left: 4px solid var(--ie-emerald);
  border-radius: 14px;
}}
#ie-header h1 {{
  margin: 0; font-size: 1.55rem; font-weight: 700;
  letter-spacing: -0.02em; color: var(--ie-text);
}}
#ie-header .ie-tagline {{
  margin: 4px 0 0; font-size: 0.92rem; color: var(--ie-muted);
}}
#ie-header .ie-badge {{
  margin-left: auto; padding: 6px 13px; border-radius: 999px;
  font-size: 0.74rem; font-weight: 600; letter-spacing: 0.06em;
  color: var(--ie-orange); background: rgba(232, 128, 58, 0.12);
  border: 1px solid rgba(232, 128, 58, 0.35);
}}

/* ── Blocchi / pannelli ────────────────────────────────────────────────── */
.gr-block, .gr-box, .gr-panel, .block, .form {{
  background: var(--ie-surface) !important;
  border: 1px solid var(--ie-border) !important;
  border-radius: 12px !important;
  color: var(--ie-text) !important;
}}
.gr-block.gr-box {{ box-shadow: 0 2px 14px rgba(0,0,0,0.35) !important; }}

label, .gr-input-label, span[data-testid="block-info"] {{
  color: var(--ie-muted) !important; font-weight: 600 !important;
  font-size: 0.84rem !important; letter-spacing: 0.01em;
}}

/* ── Area di caricamento (centrale, in evidenza) ──────────────────────── */
#ie-upload {{ margin: 0 auto; }}
#ie-upload .wrap, #ie-upload [data-testid="block-label"] + div {{
  background: var(--ie-surface-2) !important;
}}
#ie-upload .file-preview, #ie-upload .upload-container {{
  min-height: 210px !important;
  border: 2px dashed rgba(45, 125, 70, 0.55) !important;
  border-radius: 14px !important;
  background: radial-gradient(circle at 50% 0%, rgba(45,125,70,0.09), transparent 70%) !important;
  transition: border-color .18s ease, background .18s ease;
}}
#ie-upload .file-preview:hover, #ie-upload .upload-container:hover {{
  border-color: var(--ie-orange) !important;
  background: radial-gradient(circle at 50% 0%, rgba(232,128,58,0.10), transparent 70%) !important;
}}

/* ── Pulsante principale ───────────────────────────────────────────────── */
#ie-submit, button.primary, .gr-button-primary {{
  background: linear-gradient(135deg, var(--ie-emerald) 0%, var(--ie-emerald-soft) 100%) !important;
  color: #FFFFFF !important;
  border: none !important; border-radius: 10px !important;
  font-weight: 700 !important; letter-spacing: 0.02em;
  box-shadow: 0 4px 16px rgba(45, 125, 70, 0.32) !important;
  transition: transform .12s ease, box-shadow .18s ease;
}}
#ie-submit:hover, button.primary:hover {{
  transform: translateY(-1px);
  box-shadow: 0 6px 22px rgba(232, 128, 58, 0.34) !important;
}}
button.secondary, .gr-button-secondary {{
  background: var(--ie-surface-2) !important;
  color: var(--ie-text) !important;
  border: 1px solid var(--ie-border) !important;
  border-radius: 10px !important;
}}

/* ── Tabella risultati ─────────────────────────────────────────────────── */
#ie-results table {{ border-collapse: collapse !important; font-size: 0.87rem; }}
#ie-results thead th {{
  background: var(--ie-emerald) !important; color: #FFFFFF !important;
  font-weight: 700 !important; text-transform: none !important;
  border: none !important; padding: 10px 12px !important;
}}
#ie-results tbody td {{
  background: var(--ie-surface) !important; color: var(--ie-text) !important;
  border-bottom: 1px solid var(--ie-border) !important; padding: 8px 12px !important;
}}
#ie-results tbody tr:nth-child(even) td {{ background: var(--ie-surface-2) !important; }}
#ie-results tbody tr:hover td {{ background: rgba(45, 125, 70, 0.14) !important; }}

/* ── Riquadro riepilogo ────────────────────────────────────────────────── */
#ie-summary {{
  background: var(--ie-surface-2) !important;
  border: 1px solid var(--ie-border) !important;
  border-left: 3px solid var(--ie-orange) !important;
  border-radius: 12px !important; padding: 16px 20px !important;
  color: var(--ie-text) !important; min-height: 96px;
}}
#ie-summary strong {{ color: var(--ie-orange) !important; }}
#ie-summary .ie-total {{ font-size: 1.35rem; font-weight: 700; color: var(--ie-emerald-soft); }}

/* ── Piè di pagina ─────────────────────────────────────────────────────── */
#ie-footer {{
  margin-top: 26px; padding-top: 16px; text-align: center;
  border-top: 1px solid var(--ie-border);
  color: var(--ie-muted); font-size: 0.79rem;
}}
"""

HEADER_HTML = f"""
<div id="ie-header">
  {LOGO_SVG}
  <div>
    <h1>{TITLE}</h1>
    <p class="ie-tagline">{TAGLINE}</p>
  </div>
  <span class="ie-badge">FATTURE · DDT · NOTE DI CREDITO</span>
</div>
"""

FOOTER_HTML = """
<div id="ie-footer">
  Estrattore Fatture B2B · Elaborazione locale, nessun dato lascia il tuo computer ·
  Supporta fatture, ricevute fiscali, DDT e note di credito
</div>
"""

INTRO_MD = """
### Come funziona

1. **Trascina** qui le tue fatture PDF (anche decine insieme).
2. Premi **Estrai dati**.
3. **Scarica** il file Excel o CSV già formattato.

Campi estratti: numero fattura · data · totale · fornitore · partita IVA
"""

EMPTY_SUMMARY = "In attesa dei file… carica una o più fatture PDF per iniziare."

#: Temp working directories older than this are cleaned up on each run.
_TMP_MAX_AGE_SECONDS = 3600
_TMP_PREFIX = "invoice_extract_"


# ── Temp directory handling ────────────────────────────────────────────────


def _cleanup_stale_tmp_dirs(max_age: int = _TMP_MAX_AGE_SECONDS) -> None:
    """
    Remove working directories left behind by earlier runs.

    Each request gets its own directory (so concurrent uploads never collide),
    which means old ones must be swept periodically.
    """
    root = Path(tempfile.gettempdir())
    now = time.time()
    for path in root.glob(f"{_TMP_PREFIX}*"):
        try:
            if path.is_dir() and now - path.stat().st_mtime > max_age:
                shutil.rmtree(path, ignore_errors=True)
        except OSError:  # another worker may have removed it already
            continue


# ── Processing function ───────────────────────────────────────────────────


def process_invoices(files: Optional[Sequence[object]]) -> tuple:
    """
    Process uploaded PDF files.

    Args:
        files: Uploaded file paths as handed over by ``gr.File`` (filepath mode).

    Returns:
        ``(table_dict, xlsx_path, csv_path, summary_markdown)`` — the shapes
        expected by the four Gradio output components.
    """
    empty_table = {"headers": [name for name, _ in COLUMNS], "data": []}

    if not files:
        return empty_table, None, None, "⚠️ Nessun file caricato. Trascina almeno una fattura PDF."

    _cleanup_stale_tmp_dirs()

    # One directory per request → concurrent uploads cannot overwrite each other.
    tmp_dir = Path(tempfile.mkdtemp(prefix=_TMP_PREFIX))
    pdf_paths: List[Path] = []
    skipped: List[str] = []

    for f in files:
        src_name = getattr(f, "name", f)
        try:
            src = Path(str(src_name))
            if src.suffix.lower() != ".pdf":
                skipped.append(src.name)
                continue
            dest = tmp_dir / src.name
            shutil.copy2(src, dest)
            pdf_paths.append(dest)
        except OSError as exc:
            logger.warning("Impossibile leggere il file caricato %s: %s", src_name, exc)
            skipped.append(str(src_name))

    if not pdf_paths:
        return (
            empty_table,
            None,
            None,
            "❌ Nessun file PDF valido trovato. Carica file con estensione `.pdf`.",
        )

    invoices = extract_batch(pdf_paths, show_progress=False)

    # ── Export both formats so the user can pick ────────────────────────
    xlsx_path: Optional[str] = None
    csv_path: Optional[str] = None
    export_error = ""
    try:
        xlsx_path = str(export_to_excel(invoices, tmp_dir / "fatture_estratte.xlsx"))
        csv_path = str(export_to_csv(invoices, tmp_dir / "fatture_estratte.csv"))
    except OSError as exc:
        logger.error("Esportazione fallita: %s", exc)
        export_error = f"\n\n❌ **Errore di esportazione:** {exc}"

    # ── Preview table ────────────────────────────────────────────────────
    headers = [name for name, _ in COLUMNS]
    rows = []
    for inv in invoices:
        data = inv.to_dict()
        rows.append(
            [
                ("" if inv.total is None else f"{inv.total:,.2f} €") if key == "Totale"
                else str(data.get(key) or "")
                for key in headers
            ]
        )
    table = {"headers": headers, "data": rows}

    # ── Summary ──────────────────────────────────────────────────────────
    grand_total = sum(inv.total for inv in invoices if inv.total is not None)
    ok = sum(1 for inv in invoices if inv.is_ok)
    incomplete = sum(1 for inv in invoices if not inv.has_hard_error and inv.missing_fields)
    failed = sum(1 for inv in invoices if inv.has_hard_error)
    credit_notes = sum(1 for inv in invoices if inv.is_credit_note)

    lines = [
        f"**Fatture elaborate:** {len(invoices)}",
        f"**Estratte correttamente:** {ok}",
    ]
    if incomplete:
        lines.append(f"**Estratte parzialmente:** {incomplete}")
    if failed:
        lines.append(f"**File non leggibili:** {failed}")
    if credit_notes:
        lines.append(f"**Note di credito:** {credit_notes}")
    if skipped:
        lines.append(f"**File ignorati (non PDF):** {', '.join(skipped)}")
    lines.append(f'<span class="ie-total">Totale complessivo: € {grand_total:,.2f}</span>')
    lines.append("_Scarica il file Excel o CSV qui sotto._")

    return table, xlsx_path, csv_path, "\n\n".join(lines) + export_error


# ── UI build ───────────────────────────────────────────────────────────────


def build_ui() -> "gr.Blocks":
    """
    Build and return the Gradio interface.

    Raises:
        ImportError: If gradio is not installed. Use :func:`main` for a
            friendly Italian message instead.
    """
    if not GRADIO_AVAILABLE:
        raise ImportError(INSTALL_HINT) from GRADIO_IMPORT_ERROR

    theme = gr.themes.Base(
        primary_hue=gr.themes.colors.emerald,
        secondary_hue=gr.themes.colors.orange,
        neutral_hue=gr.themes.colors.gray,
        font=[gr.themes.GoogleFont("Inter"), "Segoe UI", "sans-serif"],
    )

    with gr.Blocks(title=TITLE, theme=theme, css=CUSTOM_CSS, analytics_enabled=False) as demo:
        gr.HTML(HEADER_HTML)

        # ── Upload area, prominently centred ────────────────────────────
        with gr.Row():
            with gr.Column(scale=1, min_width=260):
                gr.Markdown(INTRO_MD)
            with gr.Column(scale=2, min_width=380, elem_id="ie-upload"):
                file_input = gr.File(
                    label="Trascina qui le fatture PDF",
                    file_count="multiple",
                    file_types=[".pdf"],
                    type="filepath",
                )
                submit_btn = gr.Button(
                    "Estrai dati", variant="primary", size="lg", elem_id="ie-submit"
                )
                clear_btn = gr.Button("Svuota", variant="secondary", size="sm")

        # ── Results below ────────────────────────────────────────────────
        gr.Markdown("### Risultati")
        with gr.Row():
            with gr.Column(scale=1, min_width=260):
                output_summary = gr.Markdown(EMPTY_SUMMARY, elem_id="ie-summary")
                output_xlsx = gr.File(label="Scarica Excel (.xlsx)", interactive=False)
                output_csv = gr.File(label="Scarica CSV (.csv)", interactive=False)
            with gr.Column(scale=3, min_width=420, elem_id="ie-results"):
                output_table = gr.Dataframe(
                    headers=[name for name, _ in COLUMNS],
                    label="Anteprima dati estratti",
                    interactive=False,
                    wrap=True,
                )

        gr.HTML(FOOTER_HTML)

        submit_btn.click(
            fn=process_invoices,
            inputs=file_input,
            outputs=[output_table, output_xlsx, output_csv, output_summary],
        )
        clear_btn.click(
            fn=lambda: (
                None,
                {"headers": [name for name, _ in COLUMNS], "data": []},
                None,
                None,
                EMPTY_SUMMARY,
            ),
            inputs=None,
            outputs=[file_input, output_table, output_xlsx, output_csv, output_summary],
        )

    return demo


# ── Entry point ────────────────────────────────────────────────────────────


def main() -> int:
    """
    Launch the Gradio web UI.

    Returns a non-zero exit code (and prints an install hint in Italian) when
    gradio is not available, instead of raising a raw ImportError.
    """
    if not GRADIO_AVAILABLE:
        print(INSTALL_HINT)
        print(f"Dettaglio tecnico: {GRADIO_IMPORT_ERROR}\n")
        return 1

    logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")
    demo = build_ui()
    demo.launch(server_name="0.0.0.0", server_port=7860, share=False, show_api=False)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
