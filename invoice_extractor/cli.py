#!/usr/bin/env python3
"""
Typer CLI entry point for the B2B Invoice → Excel extractor.

Usage:
    invoice-extract --help
    invoice-extract --version
    invoice-extract /path/to/invoices/ --output ./output.xlsx
    invoice-extract /path/to/invoices/ --format csv --recursive
    invoice-extract single fattura.pdf
    invoice-extract batch ./dir/

All user-facing output is in Italian: the tool is sold to Italian small
businesses. Diagnostics go to stderr, results to stdout.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import List, Optional

import typer
from rich.console import Console
from rich.logging import RichHandler
from rich.progress import (
    BarColumn,
    MofNCompleteColumn,
    Progress,
    SpinnerColumn,
    TextColumn,
    TimeElapsedColumn,
)
from rich.table import Table

from . import __version__
from .exporter import export_to_csv, export_to_excel
from .extractor import ExtractedInvoice, extract_batch, extract_invoice
from .patterns import FIELD_LABELS_IT, load_patterns_from_config


# ── Custom TyperGroup: default command ─────────────────────────────────────
#
# A Click *group* stops parsing options at the first non-option token
# (``Group.allow_interspersed_args = False``); everything after it is handed to
# the subcommand resolver.  A plain *command* does allow interspersed options.
# So the extraction options live on the ``run`` command and the group simply
# inserts ``run`` whenever the command line does not already start with a
# subcommand name or with one of the group's own options.  That keeps both
# documented forms working:
#
#     invoice-extract ./fatture/ --output x.xlsx   →  run ./fatture/ --output …
#     invoice-extract batch ./fatture/ -o x.xlsx   →  untouched


class _DefaultCommandGroup(typer.core.TyperGroup):
    """A TyperGroup that routes anything but a subcommand to a default command."""

    default_command = "run"

    def parse_args(self, ctx: typer.Context, args: list[str]) -> list[str]:
        if not args or not self._is_group_level(ctx, args[0]):
            args = [self.default_command, *args]
        return super().parse_args(ctx, args)

    def _is_group_level(self, ctx: typer.Context, token: str) -> bool:
        """True if the first token belongs to the group (subcommand or option)."""
        if token in self.commands:
            return True
        return any(
            token in (*param.opts, *param.secondary_opts) for param in self.get_params(ctx)
        )


# Build the Typer app with our custom group class
app = typer.Typer(
    name="invoice-extract",
    help="Estrae i dati dalle fatture PDF italiane e le esporta in Excel o CSV.",
    add_completion=False,
    rich_markup_mode="rich",
    cls=_DefaultCommandGroup,
)
console = Console()
err_console = Console(stderr=True)
logger = logging.getLogger("invoice_extractor")


# ── Logging ────────────────────────────────────────────────────────────────


def _setup_logging(verbose: bool) -> None:
    """Route library logging through Rich, on stderr."""
    logging.basicConfig(
        level=logging.DEBUG if verbose else logging.WARNING,
        format="%(message)s",
        datefmt="[%X]",
        handlers=[RichHandler(console=err_console, rich_tracebacks=True, show_path=verbose)],
        force=True,
    )


# ── Callbacks ──────────────────────────────────────────────────────────────


def _version_callback(value: bool) -> None:
    """Print the version and exit (``--version``)."""
    if value:
        console.print(f"[bold]invoice-extract[/bold] versione [cyan]{__version__}[/cyan]")
        raise typer.Exit(0)


# ── Shared helpers ─────────────────────────────────────────────────────────


def _collect_pdfs(input_path: Path, recursive: bool) -> List[Path]:
    """Resolve a file or directory into a sorted list of PDF paths."""
    if input_path.is_file():
        if input_path.suffix.lower() != ".pdf":
            err_console.print(f"[red]Errore:[/red] Il file non è un PDF: {input_path}")
            raise typer.Exit(1)
        return [input_path]

    pattern = "**/*.pdf" if recursive else "*.pdf"
    return sorted(p for p in input_path.glob(pattern) if p.is_file())


def _run_extraction(pdf_files: List[Path]) -> List[ExtractedInvoice]:
    """Extract a batch of PDFs while showing a Rich progress bar."""
    with Progress(
        SpinnerColumn(style="green"),
        TextColumn("[bold]Estrazione fatture[/bold]"),
        BarColumn(bar_width=None, complete_style="green", finished_style="green"),
        MofNCompleteColumn(),
        TextColumn("[dim]{task.fields[current]}[/dim]"),
        TimeElapsedColumn(),
        console=err_console,
        transient=True,
    ) as progress:
        task_id = progress.add_task("estrazione", total=len(pdf_files), current="")

        def _on_progress(idx: int, total: int, path: Path) -> None:
            progress.update(task_id, completed=idx - 1, current=path.name)

        invoices = extract_batch(pdf_files, show_progress=False, progress_callback=_on_progress)
        progress.update(task_id, completed=len(pdf_files), current="completato")

    return invoices


def _export(
    invoices: List[ExtractedInvoice],
    output: Path,
    fmt: str,
    include_summary: bool = True,
) -> List[Path]:
    """Export in the requested format(s). Returns the written paths."""
    fmt = fmt.lower()
    if fmt not in {"xlsx", "csv", "both"}:
        err_console.print(
            f"[red]Errore:[/red] Formato non valido: '{fmt}'. Usa 'xlsx', 'csv' oppure 'both'."
        )
        raise typer.Exit(1)

    written: List[Path] = []
    try:
        if fmt in ("xlsx", "both"):
            written.append(
                export_to_excel(invoices, output.with_suffix(".xlsx"), include_summary=include_summary)
            )
        if fmt in ("csv", "both"):
            written.append(export_to_csv(invoices, output.with_suffix(".csv")))
    except OSError as exc:
        err_console.print(f"[red]Errore di salvataggio:[/red] {exc}")
        raise typer.Exit(1)

    return written


def _print_results(invoices: List[ExtractedInvoice], written: List[Path]) -> None:
    """Print the results table and any per-file warnings."""
    ok = sum(1 for inv in invoices if inv.is_ok)
    incomplete = sum(1 for inv in invoices if not inv.has_hard_error and inv.missing_fields)
    failed = sum(1 for inv in invoices if inv.has_hard_error)
    credit_notes = sum(1 for inv in invoices if inv.is_credit_note)
    grand_total = sum(inv.total for inv in invoices if inv.total is not None)

    table = Table(title="Risultati estrazione", title_style="bold green")
    table.add_column("Voce", style="bold cyan")
    table.add_column("Valore", justify="right")
    table.add_row("Fatture elaborate", str(len(invoices)))
    table.add_row("Estratte correttamente", f"[green]{ok}[/green]")
    table.add_row("Estratte parzialmente", f"[yellow]{incomplete}[/yellow]" if incomplete else "0")
    table.add_row("File non leggibili", f"[red]{failed}[/red]" if failed else "0")
    if credit_notes:
        table.add_row("Note di credito", str(credit_notes))
    table.add_row("Totale complessivo", f"€ {grand_total:,.2f}")
    for path in written:
        table.add_row("File esportato", str(path))
    console.print(table)

    # ── Per-file detail for anything that did not extract cleanly ───────
    problem_invoices = [inv for inv in invoices if not inv.is_ok]
    if problem_invoices:
        err_table = Table(title="File con problemi", title_style="yellow")
        err_table.add_column("File", style="yellow", overflow="fold")
        err_table.add_column("Dettaglio", overflow="fold")
        for inv in problem_invoices:
            if inv.has_hard_error:
                detail = "[red]" + " | ".join(inv.errors) + "[/red]"
            else:
                labels = [FIELD_LABELS_IT.get(f, f) for f in inv.missing_fields]
                detail = "Campi non trovati: " + ", ".join(labels)
            err_table.add_row(inv.filename, detail)
        err_console.print(err_table)


# ── Commands ───────────────────────────────────────────────────────────────


@app.callback()
def main(
    version: bool = typer.Option(
        False,
        "--version",
        help="Mostra la versione ed esce.",
        callback=_version_callback,
        is_eager=True,
    ),
) -> None:
    """Opzioni comuni a tutti i comandi."""


@app.command(
    "run",
    short_help="Estrae le fatture da un percorso (comando predefinito, può essere omesso).",
)
def run(
    input_path: Optional[Path] = typer.Argument(
        None,
        exists=False,
        help="Percorso di un file PDF o di una cartella di PDF.",
        show_default=False,
    ),
    output: Optional[Path] = typer.Option(
        None,
        "--output", "-o",
        help="Percorso del file di output.",
        path_type=Path,
    ),
    fmt: str = typer.Option(
        "xlsx",
        "--format", "-f",
        help="Formato di esportazione: xlsx, csv oppure both.",
    ),
    recursive: bool = typer.Option(
        False,
        "--recursive", "-r",
        help="Cerca i PDF anche nelle sottocartelle.",
    ),
    no_summary: bool = typer.Option(
        False,
        "--no-summary",
        help="Non generare il foglio di riepilogo nell'Excel.",
    ),
    config: Optional[Path] = typer.Option(
        None,
        "--config", "-c",
        help="Percorso di config.json con pattern personalizzati.",
    ),
    list_files: bool = typer.Option(
        False,
        "--list-files",
        help="Elenca solo i PDF trovati, senza estrarre.",
    ),
    verbose: bool = typer.Option(
        False,
        "--verbose", "-v",
        help="Mostra i messaggi di log dettagliati.",
    ),
) -> None:
    """
    Estrae i campi dalle fatture PDF italiane e le esporta in Excel o CSV.

    Comando predefinito: il nome 'run' può essere omesso.

    Supporta: fatture, ricevute fiscali, DDT, note di credito/debito.

    Esempi:
        invoice-extract ./invoices/
        invoice-extract ./invoices/ --output ./risultati.xlsx --recursive
        invoice-extract ./invoices/ --format both
    """
    _setup_logging(verbose)

    # ── Custom patterns ──────────────────────────────────────────────────
    if config is not None:
        if not config.exists():
            err_console.print(f"[red]Errore:[/red] Il percorso non esiste: {config}")
            raise typer.Exit(1)
        try:
            load_patterns_from_config(config)
            console.print(f"[dim]Pattern personalizzati caricati da {config}[/dim]")
        except ValueError as exc:
            err_console.print(f"[red]Errore di configurazione:[/red] {exc}")
            raise typer.Exit(1)

    # ── Resolve input paths ──────────────────────────────────────────────
    if input_path is None:
        default_dir = Path("./invoices")
        if default_dir.exists():
            input_path = default_dir
        else:
            err_console.print(
                "[red]Errore:[/red] Nessun percorso indicato e cartella predefinita "
                "./invoices/ non trovata."
            )
            err_console.print("Uso: invoice-extract PERCORSO [OPZIONI]")
            raise typer.Exit(1)
    elif not input_path.exists():
        err_console.print(f"[red]Errore:[/red] Il percorso non esiste: {input_path}")
        raise typer.Exit(1)

    pdf_files = _collect_pdfs(input_path, recursive)

    if not pdf_files:
        err_console.print(
            f"[yellow]Attenzione:[/yellow] Nessun file PDF trovato in {input_path}."
            + ("" if recursive else " Prova con --recursive per cercare nelle sottocartelle.")
        )
        raise typer.Exit(0)

    # ── List-only mode ─────────────────────────────────────────────────
    if list_files:
        table = Table(title=f"Trovati {len(pdf_files)} file PDF")
        table.add_column("#", style="dim")
        table.add_column("File", style="cyan", overflow="fold")
        table.add_column("Dimensione", justify="right")
        for idx, f in enumerate(pdf_files, 1):
            size = f.stat().st_size
            size_str = f"{size:,} byte" if size < 1024 * 1024 else f"{size / 1024 / 1024:.1f} MB"
            table.add_row(str(idx), str(f), size_str)
        console.print(table)
        raise typer.Exit(0)

    # ── Default output path ──────────────────────────────────────────────
    if output is None:
        output = Path("./output/fatture_estratte.xlsx")

    # ── Extraction ─────────────────────────────────────────────────────
    invoices = _run_extraction(pdf_files)

    # ── Export + results ─────────────────────────────────────────────────
    written = _export(invoices, output, fmt, include_summary=not no_summary)
    _print_results(invoices, written)


# ── Standalone commands ────────────────────────────────────────────────────


@app.command()
def batch(
    input_dir: Path = typer.Argument(
        ...,
        help="Cartella contenente le fatture PDF.",
        exists=True,
        file_okay=False,
        dir_okay=True,
    ),
    output: Optional[Path] = typer.Option(
        None,
        "--output", "-o",
        help="Percorso del file di output.",
        path_type=Path,
    ),
    fmt: str = typer.Option(
        "xlsx",
        "--format", "-f",
        help="Formato di esportazione: xlsx, csv oppure both.",
    ),
    recursive: bool = typer.Option(
        False,
        "--recursive", "-r",
        help="Cerca i PDF anche nelle sottocartelle.",
    ),
    verbose: bool = typer.Option(False, "--verbose", "-v", help="Log dettagliati."),
) -> None:
    """Estrae in blocco tutte le fatture PDF di una cartella."""
    _setup_logging(verbose)

    pdf_files = _collect_pdfs(input_dir, recursive)
    if not pdf_files:
        err_console.print(f"[yellow]Nessun file PDF trovato in {input_dir}.[/yellow]")
        raise typer.Exit(0)

    out = output or Path(f"./output/fatture_{input_dir.name}.xlsx")
    invoices = _run_extraction(pdf_files)
    written = _export(invoices, out, fmt)
    _print_results(invoices, written)


@app.command()
def single(
    pdf: Path = typer.Argument(
        ...,
        help="Singola fattura PDF.",
        exists=True,
        file_okay=True,
        dir_okay=False,
    ),
    output: Optional[Path] = typer.Option(
        None,
        "--output", "-o",
        help="Percorso del file di output.",
        path_type=Path,
    ),
    fmt: str = typer.Option(
        "xlsx",
        "--format", "-f",
        help="Formato di esportazione: xlsx, csv oppure both.",
    ),
    verbose: bool = typer.Option(False, "--verbose", "-v", help="Log dettagliati."),
) -> None:
    """Estrae i dati da una singola fattura PDF."""
    _setup_logging(verbose)

    if pdf.suffix.lower() != ".pdf":
        err_console.print(f"[red]Errore:[/red] Il file non è un PDF: {pdf}")
        raise typer.Exit(1)

    inv = extract_invoice(pdf)
    out = output or Path(f"./output/{pdf.stem}_estratto.xlsx")
    written = _export([inv], out, fmt)

    if inv.has_hard_error:
        for err in inv.errors:
            err_console.print(f"[red]Errore:[/red] {err}")
    else:
        table = Table(title=f"Dati estratti — {inv.filename}", title_style="bold green")
        table.add_column("Campo", style="bold cyan")
        table.add_column("Valore", overflow="fold")
        for key, val in inv.to_dict().items():
            if key not in ("Note", "File"):
                table.add_row(key, "" if val is None else str(val))
        console.print(table)

        if inv.missing_fields:
            labels = [FIELD_LABELS_IT.get(f, f) for f in inv.missing_fields]
            err_console.print(f"[yellow]Attenzione:[/yellow] campi non trovati: {', '.join(labels)}")

    for path in written:
        console.print(f"[green]Esportato in {path}[/green]")


# ── Entry point ────────────────────────────────────────────────────────────


def _entry_point() -> None:
    """Allow ``python -m invoice_extractor`` or ``invoice-extract``."""
    app()


if __name__ == "__main__":
    _entry_point()
