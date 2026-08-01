"""
Test suite for the invoice extractor.

Builds valid PDF invoices on the fly (no binary fixtures to keep in git) and
runs them through the full pipeline: PDF → extraction → Excel/CSV export.

Run with:
    python -m pytest test_smoke.py -v
"""

from __future__ import annotations

from pathlib import Path
from typing import List, Sequence

import pytest

from invoice_extractor import (
    ExtractedInvoice,
    export_to_csv,
    export_to_excel,
    extract_invoice,
)
from invoice_extractor.extractor import (
    extract_batch,
    extract_text_from_pdf,
    parse_italian_date,
    parse_italian_number,
)


# ── Minimal PDF builder ────────────────────────────────────────────────────


def build_pdf(pages: Sequence[Sequence[str]]) -> bytes:
    """
    Build a valid multi-page PDF whose pages contain the given text lines.

    A page with no lines produces a valid page with no text layer — the same
    thing pdfplumber sees for a scanned document.
    """
    n = len(pages)
    font_id = 3 + 2 * n
    objs: dict[int, bytes] = {}

    kids = " ".join(f"{3 + i} 0 R" for i in range(n))
    objs[1] = b"<< /Type /Catalog /Pages 2 0 R >>"
    objs[2] = f"<< /Type /Pages /Kids [{kids}] /Count {n} >>".encode("latin-1")

    for i, lines in enumerate(pages):
        page_id, content_id = 3 + i, 3 + n + i
        objs[page_id] = (
            f"<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] "
            f"/Contents {content_id} 0 R "
            f"/Resources << /Font << /F1 {font_id} 0 R >> >> >>"
        ).encode("latin-1")

        stream = "BT\n/F1 12 Tf\n72 720 Td\n"
        for idx, line in enumerate(lines):
            escaped = line.replace("\\", r"\\").replace("(", r"\(").replace(")", r"\)")
            if idx:
                stream += "0 -20 Td\n"
            stream += f"({escaped}) Tj\n"
        stream += "ET"
        body = stream.encode("latin-1")
        objs[content_id] = (
            b"<< /Length " + str(len(body)).encode() + b" >>\nstream\n" + body + b"\nendstream"
        )

    objs[font_id] = b"<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>"

    out = b"%PDF-1.4\n"
    offsets: dict[int, int] = {}
    for obj_id in sorted(objs):
        offsets[obj_id] = len(out)
        out += f"{obj_id} 0 obj\n".encode("latin-1") + objs[obj_id] + b"\nendobj\n"

    xref_offset = len(out)
    max_id = max(objs)
    out += f"xref\n0 {max_id + 1}\n".encode("latin-1")
    out += b"0000000000 65535 f \n"
    for obj_id in range(1, max_id + 1):
        out += f"{offsets[obj_id]:010d} 00000 n \n".encode("latin-1")
    out += (
        f"trailer\n<< /Size {max_id + 1} /Root 1 0 R >>\nstartxref\n{xref_offset}\n%%EOF\n"
    ).encode("latin-1")
    return out


def write_pdf(path: Path, pages: Sequence[Sequence[str]]) -> Path:
    """Write a generated PDF to ``path`` and return it."""
    path.write_bytes(build_pdf(pages))
    return path


# ── Invoice fixtures ───────────────────────────────────────────────────────

STANDARD_INVOICE = [
    "FATTURA N. INV-2024-001",
    "Data: 15/03/2024",
    "Fornitore: Beta S.r.l.",
    "P.IVA: 01234567890",
    "Totale Fattura: EUR 1.250,00",
    "Descrizione: Servizi di consulenza",
    "Imponibile: EUR 1.000,00",
    "IVA 22%: EUR 250,00",
]

ZERO_AMOUNT_INVOICE = [
    "FATTURA N. INV-2024-002",
    "Data: 01/04/2024",
    "Fornitore: Gamma S.p.A.",
    "P.IVA: 09876543210",
    "Totale Fattura: EUR 0,00",
    "Descrizione: Fornitura in omaggio",
]

CREDIT_NOTE = [
    "NOTA DI CREDITO N. NC-2024-014",
    "Data: 20/05/2024",
    "Fornitore: Delta S.r.l.",
    "P.IVA: 01122334455",
    "Totale Fattura: EUR -450,00",
    "Causale: Storno fattura INV-2024-001",
]

NO_VAT_INVOICE = [
    "FATTURA N. INV-2024-003",
    "Data: 10/06/2024",
    "Fornitore: Epsilon S.n.c.",
    "Totale Fattura: EUR 320,50",
]

MULTIPAGE_INVOICE = [
    [
        "FATTURA N. INV-2024-777",
        "Data: 02/07/2024",
        "Fornitore: Omega S.p.A.",
        "P.IVA: 05566778899",
        "Dettaglio prestazioni - segue a pagina 2",
    ],
    [
        "Riepilogo pagina 2",
        "Imponibile: EUR 2.000,00",
        "IVA 22%: EUR 440,00",
        "Totale Fattura: EUR 2.440,00",
    ],
]


@pytest.fixture
def standard_pdf(tmp_path: Path) -> Path:
    return write_pdf(tmp_path / "fattura_standard.pdf", [STANDARD_INVOICE])


# ── Number parsing ─────────────────────────────────────────────────────────


class TestParseItalianNumber:
    @pytest.mark.parametrize(
        "text,expected",
        [
            ("1.234,56", 1234.56),
            ("1 234,56", 1234.56),
            ("1234,56", 1234.56),
            ("1234.56", 1234.56),
            ("0,00", 0.0),
            ("€ 1.250,00", 1250.00),
            ("-450,00", -450.00),
            ("-1.234,56", -1234.56),
            ("(1.234,56)", -1234.56),  # parenthesised negative on note di credito
            ("+99,90", 99.90),
        ],
    )
    def test_parses(self, text: str, expected: float) -> None:
        assert parse_italian_number(text) == pytest.approx(expected)

    @pytest.mark.parametrize("text", ["", "abc", "n/d", "-"])
    def test_rejects_garbage(self, text: str) -> None:
        assert parse_italian_number(text) is None


class TestParseItalianDate:
    @pytest.mark.parametrize(
        "text,expected",
        [
            ("15/03/2024", "2024-03-15"),
            ("15-03-2024", "2024-03-15"),
            ("2024-03-15", "2024-03-15"),
            ("15 marzo 2024", "2024-03-15"),
            ("1 dicembre 2023", "2023-12-01"),
        ],
    )
    def test_parses(self, text: str, expected: str) -> None:
        assert parse_italian_date(text) == expected

    def test_rejects_garbage(self) -> None:
        assert parse_italian_date("non una data") is None


# ── Extraction ─────────────────────────────────────────────────────────────


class TestStandardInvoice:
    def test_all_fields_extracted(self, standard_pdf: Path) -> None:
        inv = extract_invoice(standard_pdf)

        assert inv.invoice_number == "INV-2024-001"
        assert inv.date == "2024-03-15"
        assert inv.total == pytest.approx(1250.00)
        assert inv.supplier_name in ("Beta S.r.l.", "Beta S.r.l")
        assert inv.vat_number == "01234567890"

    def test_no_errors_and_status_ok(self, standard_pdf: Path) -> None:
        inv = extract_invoice(standard_pdf)

        assert inv.errors == []
        assert inv.missing_fields == []
        assert inv.is_ok
        assert not inv.has_hard_error
        assert inv.status_it == "OK"
        assert inv.document_type_it == "Fattura"
        assert not inv.is_credit_note


class TestZeroAmountInvoice:
    """A 0,00 total is a real value, not a missing field."""

    def test_zero_total_is_extracted(self, tmp_path: Path) -> None:
        pdf = write_pdf(tmp_path / "zero.pdf", [ZERO_AMOUNT_INVOICE])
        inv = extract_invoice(pdf)

        assert inv.total == 0.0
        assert inv.total is not None
        assert "total" not in inv.missing_fields
        assert inv.is_ok
        assert not inv.is_credit_note

    def test_zero_total_written_to_excel(self, tmp_path: Path) -> None:
        from openpyxl import load_workbook

        pdf = write_pdf(tmp_path / "zero.pdf", [ZERO_AMOUNT_INVOICE])
        out = tmp_path / "zero.xlsx"
        export_to_excel([extract_invoice(pdf)], out)

        ws = load_workbook(out)["Fatture"]
        headers = [c.value for c in ws[1]]
        assert ws.cell(row=2, column=headers.index("Totale") + 1).value == 0


class TestCreditNote:
    """Note di credito carry a negative total and must survive the pipeline."""

    def test_negative_total_extracted(self, tmp_path: Path) -> None:
        pdf = write_pdf(tmp_path / "nota_credito.pdf", [CREDIT_NOTE])
        inv = extract_invoice(pdf)

        assert inv.total == pytest.approx(-450.00)
        assert inv.is_credit_note
        assert inv.document_type_it == "Nota di credito"
        assert not inv.has_hard_error

    def test_negative_total_in_grand_total(self, tmp_path: Path) -> None:
        from openpyxl import load_workbook

        invoices = [
            extract_invoice(write_pdf(tmp_path / "fattura.pdf", [STANDARD_INVOICE])),
            extract_invoice(write_pdf(tmp_path / "nota.pdf", [CREDIT_NOTE])),
        ]
        out = tmp_path / "misto.xlsx"
        export_to_excel(invoices, out)

        ws = load_workbook(out)["Fatture"]
        # Grand total sits two rows below the last data row.
        total_row = len(invoices) + 3
        assert ws.cell(row=total_row, column=1).value == "TOTALE COMPLESSIVO"
        headers = [c.value for c in ws[1]]
        grand = ws.cell(row=total_row, column=headers.index("Totale") + 1).value
        assert grand == pytest.approx(1250.00 - 450.00)


class TestMissingVatNumber:
    def test_missing_vat_is_soft_issue(self, tmp_path: Path) -> None:
        pdf = write_pdf(tmp_path / "senza_piva.pdf", [NO_VAT_INVOICE])
        inv = extract_invoice(pdf)

        assert inv.vat_number is None
        assert "vat_number" in inv.missing_fields
        # Not a hard failure: the rest of the invoice is still usable.
        assert inv.errors == []
        assert not inv.has_hard_error
        assert not inv.is_ok
        assert inv.status_it == "Incompleto"
        assert inv.invoice_number == "INV-2024-003"
        assert inv.total == pytest.approx(320.50)

    def test_note_mentions_missing_field_in_italian(self, tmp_path: Path) -> None:
        pdf = write_pdf(tmp_path / "senza_piva.pdf", [NO_VAT_INVOICE])
        inv = extract_invoice(pdf)

        assert "partita IVA" in inv.notes_it


class TestMultiPageInvoice:
    def test_text_from_all_pages_is_used(self, tmp_path: Path) -> None:
        pdf = write_pdf(tmp_path / "multipagina.pdf", MULTIPAGE_INVOICE)

        text = extract_text_from_pdf(pdf)
        assert "Riepilogo pagina 2" in text
        assert "FATTURA N. INV-2024-777" in text

    def test_fields_span_pages(self, tmp_path: Path) -> None:
        pdf = write_pdf(tmp_path / "multipagina.pdf", MULTIPAGE_INVOICE)
        inv = extract_invoice(pdf)

        # Number/date/VAT come from page 1, the total only from page 2.
        assert inv.invoice_number == "INV-2024-777"
        assert inv.vat_number == "05566778899"
        assert inv.total == pytest.approx(2440.00)
        assert inv.is_ok


# ── Robustness ─────────────────────────────────────────────────────────────


class TestBrokenInputs:
    """Bad files must produce an Italian message, never an exception."""

    def test_corrupt_pdf(self, tmp_path: Path) -> None:
        pdf = tmp_path / "corrotto.pdf"
        pdf.write_bytes(b"Questo non e' un PDF, solo testo casuale.")

        inv = extract_invoice(pdf)

        assert inv.has_hard_error
        assert inv.status_it == "Errore"
        assert inv.total is None

    def test_empty_file(self, tmp_path: Path) -> None:
        pdf = tmp_path / "vuoto.pdf"
        pdf.write_bytes(b"")

        inv = extract_invoice(pdf)

        assert inv.has_hard_error
        assert "vuoto" in inv.notes_it.lower()

    def test_scanned_pdf_without_text_layer(self, tmp_path: Path) -> None:
        pdf = write_pdf(tmp_path / "scansione.pdf", [[]])

        inv = extract_invoice(pdf)

        assert inv.has_hard_error
        assert "scansione" in inv.notes_it.lower()
        assert "ocr" in inv.notes_it.lower()

    def test_missing_file(self, tmp_path: Path) -> None:
        inv = extract_invoice(tmp_path / "inesistente.pdf")

        assert inv.has_hard_error
        assert "non trovato" in inv.notes_it.lower()

    def test_batch_continues_past_a_broken_file(self, tmp_path: Path) -> None:
        good = write_pdf(tmp_path / "buona.pdf", [STANDARD_INVOICE])
        bad = tmp_path / "rotta.pdf"
        bad.write_bytes(b"non un pdf")

        results = extract_batch([bad, good], show_progress=False)

        assert len(results) == 2
        assert results[0].has_hard_error
        assert results[1].is_ok

    def test_progress_callback_is_called_per_file(self, tmp_path: Path) -> None:
        pdfs = [
            write_pdf(tmp_path / "a.pdf", [STANDARD_INVOICE]),
            write_pdf(tmp_path / "b.pdf", [CREDIT_NOTE]),
        ]
        seen: List[str] = []

        extract_batch(
            pdfs,
            show_progress=False,
            progress_callback=lambda idx, total, path: seen.append(path.name),
        )

        assert seen == ["a.pdf", "b.pdf"]


# ── Excel export ───────────────────────────────────────────────────────────


class TestExcelExport:
    @pytest.fixture
    def workbook_path(self, tmp_path: Path) -> Path:
        invoices = [
            extract_invoice(write_pdf(tmp_path / "a.pdf", [STANDARD_INVOICE])),
            extract_invoice(write_pdf(tmp_path / "b.pdf", [ZERO_AMOUNT_INVOICE])),
            extract_invoice(write_pdf(tmp_path / "c.pdf", [NO_VAT_INVOICE])),
        ]
        out = tmp_path / "export.xlsx"
        export_to_excel(invoices, out)
        return out

    def test_file_is_created_and_not_empty(self, workbook_path: Path) -> None:
        assert workbook_path.exists()
        assert workbook_path.stat().st_size > 0

    def test_sheets_are_italian(self, workbook_path: Path) -> None:
        from openpyxl import load_workbook

        wb = load_workbook(workbook_path)
        assert "Fatture" in wb.sheetnames
        assert "Riepilogo" in wb.sheetnames

    def test_frozen_panes(self, workbook_path: Path) -> None:
        from openpyxl import load_workbook

        assert load_workbook(workbook_path)["Fatture"].freeze_panes == "A2"

    def test_auto_filter_covers_header_and_data_only(self, workbook_path: Path) -> None:
        from openpyxl import load_workbook

        ws = load_workbook(workbook_path)["Fatture"]
        # 3 invoices → header row 1 + rows 2..4. The grand total (row 6) must
        # stay outside the filter range or sorting would drag it along.
        assert ws.auto_filter.ref is not None
        assert ws.auto_filter.ref.endswith("4")

    def test_zebra_striping_on_alternate_rows(self, workbook_path: Path) -> None:
        from openpyxl import load_workbook

        ws = load_workbook(workbook_path)["Fatture"]
        assert ws.cell(row=3, column=1).fill.start_color.rgb != ws.cell(
            row=2, column=1
        ).fill.start_color.rgb

    def test_summary_reports_counts(self, workbook_path: Path) -> None:
        from openpyxl import load_workbook

        ws = load_workbook(workbook_path)["Riepilogo"]
        rows = {ws.cell(row=r, column=1).value: ws.cell(row=r, column=2).value
                for r in range(1, ws.max_row + 1)}
        assert rows["PDF elaborati"] == 3
        assert rows["Estratti correttamente"] == 2  # the no-VAT one is partial
        assert rows["Estratti parzialmente"] == 1

    def test_export_without_summary(self, tmp_path: Path) -> None:
        from openpyxl import load_workbook

        inv = extract_invoice(write_pdf(tmp_path / "a.pdf", [STANDARD_INVOICE]))
        out = tmp_path / "senza_riepilogo.xlsx"
        export_to_excel([inv], out, include_summary=False)

        assert load_workbook(out).sheetnames == ["Fatture"]

    def test_empty_invoice_list_still_exports(self, tmp_path: Path) -> None:
        out = tmp_path / "vuoto.xlsx"
        export_to_excel([], out)

        assert out.exists()

    def test_creates_missing_output_directory(self, tmp_path: Path) -> None:
        inv = extract_invoice(write_pdf(tmp_path / "a.pdf", [STANDARD_INVOICE]))
        out = tmp_path / "nuova" / "sotto" / "export.xlsx"
        export_to_excel([inv], out)

        assert out.exists()


# ── CSV export ─────────────────────────────────────────────────────────────


class TestCsvExport:
    def test_csv_uses_italian_conventions(self, tmp_path: Path) -> None:
        inv = extract_invoice(write_pdf(tmp_path / "a.pdf", [STANDARD_INVOICE]))
        out = tmp_path / "export.csv"
        export_to_csv([inv], out)

        raw = out.read_bytes()
        assert raw.startswith(b"\xef\xbb\xbf")  # UTF-8 BOM for Italian Excel

        text = raw.decode("utf-8-sig")
        assert text.splitlines()[0].startswith("File;Numero Fattura;Data;Totale")
        assert "1250,00" in text  # comma decimal separator

    def test_csv_keeps_negative_totals(self, tmp_path: Path) -> None:
        inv = extract_invoice(write_pdf(tmp_path / "nota.pdf", [CREDIT_NOTE]))
        out = tmp_path / "nota.csv"
        export_to_csv([inv], out)

        assert "-450,00" in out.read_text(encoding="utf-8-sig")

    def test_csv_has_grand_total_row(self, tmp_path: Path) -> None:
        invoices = [
            extract_invoice(write_pdf(tmp_path / "a.pdf", [STANDARD_INVOICE])),
            extract_invoice(write_pdf(tmp_path / "b.pdf", [NO_VAT_INVOICE])),
        ]
        out = tmp_path / "export.csv"
        export_to_csv(invoices, out)

        text = out.read_text(encoding="utf-8-sig")
        assert "TOTALE COMPLESSIVO" in text
        assert "1570,50" in text


# ── Serialization ──────────────────────────────────────────────────────────


class TestToDict:
    def test_keys_match_export_columns(self) -> None:
        from invoice_extractor.exporter import COLUMNS

        keys = list(ExtractedInvoice(filename="x.pdf").to_dict().keys())
        assert keys == [name for name, _ in COLUMNS]

    def test_empty_invoice_serializes(self) -> None:
        data = ExtractedInvoice(filename="x.pdf").to_dict()

        assert data["File"] == "x.pdf"
        assert data["Totale"] is None
        assert data["Esito"] == "Incompleto" or data["Esito"] == "OK"


# ── Custom patterns (--config) ─────────────────────────────────────────────


class TestConfigPatterns:
    def test_custom_pattern_takes_priority(self, tmp_path: Path) -> None:
        import json

        from invoice_extractor import patterns as patterns_mod
        from invoice_extractor.patterns import load_patterns_from_config

        original = dict(patterns_mod.PATTERNS)
        try:
            cfg = tmp_path / "config.json"
            cfg.write_text(
                json.dumps({"patterns": {"invoice_number": [r"CODICE\s+(\S+)"]}}),
                encoding="utf-8",
            )
            load_patterns_from_config(cfg)

            assert patterns_mod.PATTERNS["invoice_number"][0] == r"CODICE\s+(\S+)"
            # Built-ins are kept as a fallback.
            assert len(patterns_mod.PATTERNS["invoice_number"]) > 1
        finally:
            patterns_mod.PATTERNS.clear()
            patterns_mod.PATTERNS.update(original)

    def test_invalid_json_raises_italian_error(self, tmp_path: Path) -> None:
        from invoice_extractor.patterns import load_patterns_from_config

        cfg = tmp_path / "config.json"
        cfg.write_text("{ non json", encoding="utf-8")

        with pytest.raises(ValueError, match="non valido"):
            load_patterns_from_config(cfg)

    def test_shipped_config_is_loadable(self) -> None:
        from invoice_extractor import patterns as patterns_mod
        from invoice_extractor.patterns import load_patterns_from_config

        config_path = Path(__file__).parent / "config.json"
        if not config_path.exists():
            pytest.skip("config.json non presente")

        original = dict(patterns_mod.PATTERNS)
        try:
            load_patterns_from_config(config_path)
        finally:
            patterns_mod.PATTERNS.clear()
            patterns_mod.PATTERNS.update(original)


# ── Optional web UI ────────────────────────────────────────────────────────


class TestWebModuleIsOptional:
    def test_import_never_fails(self) -> None:
        """web.py must import cleanly even without gradio installed."""
        from invoice_extractor import web

        assert isinstance(web.GRADIO_AVAILABLE, bool)

    def test_install_hint_is_italian_and_actionable(self) -> None:
        from invoice_extractor import web

        assert "pip install" in web.INSTALL_HINT
        assert "gradio" in web.INSTALL_HINT

    def test_build_ui_raises_clear_error_without_gradio(self) -> None:
        from invoice_extractor import web

        if web.GRADIO_AVAILABLE:
            pytest.skip("gradio disponibile: il fallback non è applicabile")

        with pytest.raises(ImportError, match="pip install"):
            web.build_ui()

        assert web.main() == 1


# ── CLI ────────────────────────────────────────────────────────────────────


class TestCli:
    @pytest.fixture
    def runner(self):
        from typer.testing import CliRunner

        return CliRunner()

    def test_version_flag(self, runner) -> None:
        from invoice_extractor import __version__
        from invoice_extractor.cli import app

        result = runner.invoke(app, ["--version"])

        assert result.exit_code == 0
        assert __version__ in result.stdout

    def test_help_is_italian(self, runner) -> None:
        from invoice_extractor.cli import app

        result = runner.invoke(app, ["--help"])

        assert result.exit_code == 0
        assert "fatture" in result.stdout.lower()

    def test_end_to_end_directory_extraction(self, runner, tmp_path: Path) -> None:
        from invoice_extractor.cli import app

        src = tmp_path / "fatture"
        src.mkdir()
        write_pdf(src / "a.pdf", [STANDARD_INVOICE])
        write_pdf(src / "b.pdf", [CREDIT_NOTE])
        out = tmp_path / "out" / "risultati.xlsx"

        result = runner.invoke(app, [str(src), "--output", str(out)])

        assert result.exit_code == 0, result.output
        assert out.exists()

    def test_format_both_writes_xlsx_and_csv(self, runner, tmp_path: Path) -> None:
        from invoice_extractor.cli import app

        src = tmp_path / "fatture"
        src.mkdir()
        write_pdf(src / "a.pdf", [STANDARD_INVOICE])
        out = tmp_path / "risultati.xlsx"

        result = runner.invoke(app, [str(src), "--output", str(out), "--format", "both"])

        assert result.exit_code == 0, result.output
        assert out.with_suffix(".xlsx").exists()
        assert out.with_suffix(".csv").exists()

    def test_invalid_format_is_rejected(self, runner, tmp_path: Path) -> None:
        from invoice_extractor.cli import app

        src = tmp_path / "fatture"
        src.mkdir()
        write_pdf(src / "a.pdf", [STANDARD_INVOICE])

        result = runner.invoke(app, [str(src), "--format", "pdf"])

        assert result.exit_code == 1

    def test_nonexistent_path_is_rejected(self, runner, tmp_path: Path) -> None:
        from invoice_extractor.cli import app

        result = runner.invoke(app, [str(tmp_path / "manca")])

        assert result.exit_code != 0

    def test_single_command(self, runner, tmp_path: Path) -> None:
        from invoice_extractor.cli import app

        pdf = write_pdf(tmp_path / "a.pdf", [STANDARD_INVOICE])
        out = tmp_path / "singola.xlsx"

        result = runner.invoke(app, ["single", str(pdf), "--output", str(out)])

        assert result.exit_code == 0, result.output
        assert out.exists()

    def test_batch_command(self, runner, tmp_path: Path) -> None:
        from invoice_extractor.cli import app

        src = tmp_path / "fatture"
        src.mkdir()
        write_pdf(src / "a.pdf", [STANDARD_INVOICE])
        out = tmp_path / "blocco.xlsx"

        result = runner.invoke(app, ["batch", str(src), "--output", str(out)])

        assert result.exit_code == 0, result.output
        assert out.exists()

    def test_list_files_does_not_extract(self, runner, tmp_path: Path) -> None:
        from invoice_extractor.cli import app

        src = tmp_path / "fatture"
        src.mkdir()
        write_pdf(src / "a.pdf", [STANDARD_INVOICE])

        result = runner.invoke(app, [str(src), "--list-files"])

        assert result.exit_code == 0
        assert not (tmp_path / "output").exists()

    def test_empty_directory_exits_cleanly(self, runner, tmp_path: Path) -> None:
        from invoice_extractor.cli import app

        src = tmp_path / "vuota"
        src.mkdir()

        result = runner.invoke(app, [str(src)])

        assert result.exit_code == 0
