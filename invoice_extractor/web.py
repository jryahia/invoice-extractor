"""
Optional Gradio web UI for the invoice extractor.

Provides a simple drag-and-drop interface for uploading PDF invoices,
viewing extracted data in a table, and downloading the Excel export.

Launch with: python -m invoice_extractor.web
"""

from __future__ import annotations

import tempfile
from pathlib import Path
from typing import List

import gradio as gr

from .extractor import ExtractedInvoice, extract_batch
from .exporter import export_to_excel


# ── Constants ──────────────────────────────────────────────────────────────

TITLE = "B2B Invoice Extractor"
DESCRIPTION = """
Upload one or more Italian PDF invoices (fatture, ricevute fiscali) to extract:
- Invoice Number
- Date
- Total (€)
- Supplier Name
- VAT Number (Partita IVA)

Download the results as a formatted Excel (.xlsx) file.
"""


# ── Processing function ───────────────────────────────────────────────────


def process_invoices(files: List[tempfile.NamedTemporaryFile]) -> tuple:
    """
    Process uploaded PDF files and return:
      - DataFrame-ready data as list of dicts
      - Path to output Excel file
      - Summary text
    """
    if not files:
        return [], None, "No files uploaded."

    # Save uploaded files to temp directory
    tmp_dir = Path(tempfile.mkdtemp(prefix="invoice_extract_"))
    pdf_paths: List[Path] = []

    for f in files:
        dest = tmp_dir / Path(f.name).name
        data = Path(f.name).read_bytes() if hasattr(f, "name") else bytes()
        # Gradio hands us NamedTemporaryFile objects
        try:
            dest.write_bytes(Path(f.name).read_bytes())
        except Exception as exc:
            # Fallback: try reading the file object
            try:
                import shutil
                shutil.copy2(f.name, dest)
            except Exception as exc2:
                return [], None, f"Error reading {f.name}: {exc2}"

        if dest.suffix.lower() == ".pdf":
            pdf_paths.append(dest)

    if not pdf_paths:
        return [], None, "No valid PDF files found."

    # Extract
    invoices = extract_batch(pdf_paths, show_progress=False)

    # Export
    output_path = tmp_dir / "extracted_invoices.xlsx"
    export_to_excel(invoices, output_path)

    # Build table data
    table_data = [inv.to_dict() for inv in invoices]

    # Summary
    totals = [inv.total for inv in invoices if inv.total is not None]
    grand_total = sum(totals)
    success = sum(1 for inv in invoices if not inv.errors or all(
        "Could not extract" not in e for e in inv.errors
    ))
    summary_text = (
        f"**Processed:** {len(invoices)} invoices\\n"
        f"**Successfully extracted:** {success}\\n"
        f"**Grand Total:** €{grand_total:,.2f}\\n"
        f"**Download:** Click the file below"
    )

    return table_data, str(output_path), summary_text


# ── UI build ───────────────────────────────────────────────────────────────


def build_ui() -> gr.Blocks:
    """Build and return the Gradio interface."""
    with gr.Blocks(title=TITLE, theme=gr.themes.Soft()) as demo:
        gr.Markdown(f"# {TITLE}")
        gr.Markdown(DESCRIPTION)

        with gr.Row():
            with gr.Column(scale=1):
                file_input = gr.File(
                    label="Upload PDF Invoices",
                    file_count="multiple",
                    file_types=[".pdf"],
                    type="filepath",
                )
                submit_btn = gr.Button("Extract Data", variant="primary", size="lg")

            with gr.Column(scale=2):
                output_summary = gr.Markdown("Awaiting upload...")
                output_table = gr.Dataframe(
                    label="Extracted Data",
                    interactive=False,
                    wrap=True,
                )
                output_file = gr.File(
                    label="Download Excel Export",
                    interactive=False,
                )

        submit_btn.click(
            fn=process_invoices,
            inputs=file_input,
            outputs=[output_table, output_file, output_summary],
        )

    return demo


# ── Entry point ────────────────────────────────────────────────────────────


def main() -> None:
    """Launch the Gradio web UI."""
    demo = build_ui()
    demo.launch(
        server_name="0.0.0.0",
        server_port=7860,
        share=False,
    )


if __name__ == "__main__":
    main()
