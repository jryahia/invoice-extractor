#!/usr/bin/env python3
"""
Typer CLI entry point for the B2B Invoice → Excel extractor.

Usage:
    invoice-extract --help
    invoice-extract /path/to/invoices/ --output ./output.xlsx
    invoice-extract inv1.pdf inv2.pdf --recursive
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from typing import List, Optional

import typer
from rich.console import Console
from rich.progress import BarColumn, Progress, TextColumn
from rich.table import Table

from .extractor import extract_batch, extract_invoice
from .exporter import export_to_excel

app = typer.Typer(
    name="invoice-extract",
    help="Extract data from Italian B2B PDF invoices to Excel.",
    add_completion=False,
    rich_markup_mode="rich",
)
console = Console()
err_console = Console(stderr=True)


# ── Callbacks ──────────────────────────────────────────────────────────────


def _validate_path(ctx: typer.Context, param: typer.CallbackParam, value: Optional[Path]) -> Optional[Path]:
    """Ensure input path exists."""
    if value is not None and not value.exists():
        raise typer.BadParameter(f"Path does not exist: {value}")
    return value


# ── Commands ───────────────────────────────────────────────────────────────


@app.callback(invoke_without_command=True)
def main(
    ctx: typer.Context,
    input_path: Optional[Path] = typer.Argument(
        None,
        exists=False,
        help="Path to PDF file or directory of PDFs.",
        callback=_validate_path,
        show_default=False,
    ),
    output: Optional[Path] = typer.Option(
        None,
        "--output", "-o",
        help="Output .xlsx file path.",
        path_type=Path,
    ),
    recursive: bool = typer.Option(
        False,
        "--recursive", "-r",
        help="Search PDFs recursively in directory.",
    ),
    no_summary: bool = typer.Option(
        False,
        "--no-summary",
        help="Skip the summary sheet in Excel output.",
    ),
    config: Optional[Path] = typer.Option(
        None,
        "--config", "-c",
        help="Path to config.json (for custom patterns).",
        callback=_validate_path,
    ),
    list_files: bool = typer.Option(
        False,
        "--list-files",
        help="Only list found PDF files without extracting.",
    ),
) -> None:
    """
    Extract fields from Italian PDF invoices and export to Excel.

    Supports: fatture, ricevute fiscali, DDT, note di credito/debito.

    Examples:
        invoice-extract ./invoices/
        invoice-extract ./invoices/ --output ./output.xlsx --recursive
        invoice-extract inv1.pdf inv2.pdf inv3.pdf
    """
    if ctx.invoked_subcommand is not None:
        return

    # ── Resolve input paths ──────────────────────────────────────────────
    if input_path is None:
        # Try default ./invoices directory
        default_dir = Path("./invoices")
        if default_dir.exists():
            input_path = default_dir
        else:
            console.print("[red]Error:[/red] No input path provided and no default ./invoices/ directory found.")
            console.print("Usage: invoice-extract PATH [OPTIONS]")
            raise typer.Exit(1)

    pdf_files: List[Path] = []
    if input_path.is_file():
        if input_path.suffix.lower() == ".pdf":
            pdf_files = [input_path]
        else:
            err_console.print(f"[red]Error:[/red] Not a PDF file: {input_path}")
            raise typer.Exit(1)
    elif input_path.is_dir():
        pattern = "**/*.pdf" if recursive else "*.pdf"
        pdf_files = sorted(input_path.glob(pattern))
        if not pdf_files:
            err_console.print(f"[yellow]Warning:[/yellow] No PDF files found in {input_path}.")
            raise typer.Exit(0)

    if not pdf_files:
        err_console.print("[red]Error:[/red] No PDF files specified.")
        raise typer.Exit(1)

    # ── List-only mode ─────────────────────────────────────────────────
    if list_files:
        table = Table(title=f"Found {len(pdf_files)} PDF file(s)")
        table.add_column("#", style="dim")
        table.add_column("Filename", style="cyan")
        table.add_column("Size", justify="right")
        for idx, f in enumerate(pdf_files, 1):
            size = f.stat().st_size
            size_str = f"{size:,} bytes" if size < 1024 * 1024 else f"{size / 1024 / 1024:.1f} MB"
            table.add_row(str(idx), str(f), size_str)
        console.print(table)
        raise typer.Exit(0)

    # ── Default output path ──────────────────────────────────────────────
    if output is None:
        output = Path("./output/extracted_invoices.xlsx")

    # ── Extraction ─────────────────────────────────────────────────────
    err_console.print(f"[bold]Processing {len(pdf_files)} PDF file(s)...[/bold]")

    invoices = extract_batch(pdf_files)

    # ── Export ───────────────────────────────────────────────────────────
    output_path = export_to_excel(
        invoices=invoices,
        output_path=output,
        include_summary=not no_summary,
    )

    # ── Results summary ─────────────────────────────────────────────────
    success = sum(1 for inv in invoices if not inv.errors or all(
        "Could not extract" not in e for e in inv.errors
    ))
    errors = len(invoices) - success
    totals = [inv.total for inv in invoices if inv.total is not None]
    grand_total = sum(totals)

    # Results table
    table = Table(title="Extraction Results")
    table.add_column("Metric", style="bold cyan")
    table.add_column("Value", justify="right")
    table.add_row("Total invoices", str(len(invoices)))
    table.add_row("Successfully extracted", str(success))
    table.add_row("With errors", str(errors) if errors else "[green]0[/green]")
    table.add_row("Grand Total", f"€{grand_total:,.2f}")
    table.add_row("Output file", str(output_path))
    console.print(table)

    # Detail table for errors
    error_invoices = [inv for inv in invoices if inv.errors and any(
        "Could not extract" in e for e in inv.errors
    )]
    if error_invoices:
        err_table = Table(title="Invoices with Extraction Issues", style="yellow")
        err_table.add_column("File", style="yellow")
        err_table.add_column("Issues")
        for inv in error_invoices:
            issues = [e for e in inv.errors if "Could not extract" in e]
            err_table.add_row(inv.filename, "; ".join(issues))
        err_console.print(err_table)


# ── Standalone command ─────────────────────────────────────────────────────


@app.command()
def batch(
    input_dir: Path = typer.Argument(
        ...,
        help="Directory containing PDF invoices.",
        exists=True,
        file_okay=False,
        dir_okay=True,
    ),
    output: Optional[Path] = typer.Option(
        None,
        "--output", "-o",
        help="Output .xlsx file path.",
        path_type=Path,
    ),
    recursive: bool = typer.Option(
        False,
        "--recursive", "-r",
        help="Search PDFs recursively.",
    ),
) -> None:
    """Batch extract all PDF invoices from a directory."""
    pdf_files = sorted(input_dir.glob("**/*.pdf" if recursive else "*.pdf"))

    if not pdf_files:
        err_console.print(f"[yellow]No PDF files found in {input_dir}.[/yellow]")
        raise typer.Exit(0)

    out = output or Path(f"./output/invoices_batch_{input_dir.name}.xlsx")
    invoices = extract_batch(pdf_files)
    export_to_excel(invoices, out)

    console.print(f"[green]Exported {len(invoices)} invoices to {out.resolve()}[/green]")


@app.command()
def single(
    pdf: Path = typer.Argument(
        ...,
        help="Single PDF invoice file.",
        exists=True,
        file_okay=True,
        dir_okay=False,
    ),
    output: Optional[Path] = typer.Option(
        None,
        "--output", "-o",
        help="Output .xlsx file path.",
        path_type=Path,
    ),
) -> None:
    """Extract data from a single PDF invoice."""
    inv = extract_invoice(pdf)
    out = output or Path(f"./output/{pdf.stem}_extracted.xlsx")
    export_to_excel([inv], out)

    console.print(f"[green]Exported to {out.resolve()}[/green]")
    for key, val in inv.to_dict().items():
        if key != "Errors":
            console.print(f"  [cyan]{key}:[/cyan] {val}")
    if inv.errors:
        for err in inv.errors:
            err_console.print(f"  [yellow]Warning:[/yellow] {err}")


# ── Entry point ────────────────────────────────────────────────────────────


def _entry_point() -> None:
    """Allow ``python -m invoice_extractor run`` or ``invoice-extract``."""
    app()


if __name__ == "__main__":
    _entry_point()
