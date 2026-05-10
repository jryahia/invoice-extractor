"""
Quick smoke test: create a sample Italian invoice PDF and run extraction.
"""

import tempfile
from pathlib import Path

# ── Create a minimal test PDF (hand-crafted valid PDF) ─────────────────────

SAMPLE_PDF = b"""%PDF-1.4
1 0 obj
<< /Type /Catalog /Pages 2 0 R >>
endobj
2 0 obj
<< /Type /Pages /Kids [3 0 R] /Count 1 >>
endobj
3 0 obj
<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792]
   /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>
endobj
4 0 obj
<< /Length 500 >>
stream
BT
/F1 12 Tf
100 700 Td
(FATTURA N. INV-2024-001) Tj
0 -20 Td
(Data: 15/03/2024) Tj
0 -20 Td
(Fornitore: Beta S.r.l.) Tj
0 -20 Td
(P.IVA: 01234567890) Tj
0 -40 Td
(Totale Fattura: EUR 1.250,00) Tj
0 -20 Td
(Descrizione: Servizi di consulenza) Tj
0 -20 Td
(Imponibile: EUR 1.000,00) Tj
0 -20 Td
(IVA 22%: EUR 250,00) Tj
ET
endstream
endobj
5 0 obj
<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>
endobj
xref
0 6
0000000000 65535 f 
0000000009 00000 n 
0000000056 00000 n 
0000000111 00000 n 
0000000267 00000 n 
0000000815 00000 n 
trailer
<< /Size 6 /Root 1 0 R >>
startxref
867
%%EOF
"""


def test_extraction():
    """Test extraction against a minimal valid PDF invoice."""
    from invoice_extractor import extract_invoice, export_to_excel

    # Write sample PDF
    tmp = Path(tempfile.mktemp(suffix=".pdf"))
    tmp.write_bytes(SAMPLE_PDF)

    # Extract
    inv = extract_invoice(tmp)
    print(f"File: {inv.filename}")
    print(f"  Invoice #: {inv.invoice_number}")
    print(f"  Date: {inv.date}")
    print(f"  Total: {inv.total}")
    print(f"  Supplier: {inv.supplier_name}")
    print(f"  VAT: {inv.vat_number}")
    print(f"  Errors: {inv.errors}")

    # Export
    out = tmp.with_suffix(".xlsx")
    export_to_excel([inv], out)
    print(f"\nExported to: {out}")
    print(f"File size: {out.stat().st_size} bytes")

    # Cleanup
    tmp.unlink()
    out.unlink()

    # Assertions
    assert inv.invoice_number == "INV-2024-001", f"Got: {inv.invoice_number}"
    assert inv.date == "2024-03-15", f"Got: {inv.date}"
    assert inv.total == 1250.00, f"Got: {inv.total}"
    assert inv.supplier_name in ("Beta S.r.l.", "Beta S.r.l"), f"Got: {inv.supplier_name}"
    assert inv.vat_number == "01234567890", f"Got: {inv.vat_number}"
    assert len(inv.errors) == 0, f"Got errors: {inv.errors}"

    print("\n✅ All tests passed!")


if __name__ == "__main__":
    test_extraction()
