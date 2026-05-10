# Invoice Extractor — B2B PDF → Excel

Extract structured data from **Italian PDF invoices** (fatture, ricevute fiscali
, DDT, note di credito/debito) and export to a clean, formatted Excel (.xlsx) 
file with totals and summary.

Built for Italian small businesses dealing with batch invoice processing.

## Features

- 📄 **Batch processing** — drop in 50+ PDFs at once
- 🔍 **Smart extraction** — invoice number, date, total, supplier name, VAT (Partita IVA)
- 🇮🇹 **Italian invoice support** — patterns for fatture, ricevute, DDT, note
- 📊 **Formatted Excel export** — headers, column widths, currency format, totals row, summary sheet
- 🖥️ **CLI** — typer-based with rich progress output
- 🌐 **Web UI** — optional Gradio drag-and-drop interface
- 📝 **Configurable** — custom regex patterns via config.json

## Installation

```bash
# Clone the repo
git clone git@github.com:jryahia/invoice-extractor.git
cd invoice-extractor

# Install
pip install -r requirements.txt
```

Or install as a package:

```bash
pip install -e .
```

## Usage

### CLI

```bash
# Process a directory of PDFs
invoice-extract ./invoices/

# Process recursively
invoice-extract ./invoices/ --recursive

# Specify output file
invoice-extract ./invoices/ -o results.xlsx

# Single invoice
invoice-extract single invoice.pdf

# List PDFs without extracting
invoice-extract ./invoices/ --list-files
```

### Web UI

```bash
python -m invoice_extractor.web
```

Opens at `http://localhost:7860`. Drag-and-drop PDFs, click "Extract Data", 
download the Excel file.

### Python API

```python
from invoice_extractor import extract_batch, export_to_excel

# Extract from PDFs
invoices = extract_batch([Path("inv1.pdf"), Path("inv2.pdf")])

# Export to Excel
export_to_excel(invoices, "output.xlsx")
```

## Project Structure

```
invoice_extractor/
├── invoice_extractor/
│   ├── __init__.py         # Package init, version
│   ├── __main__.py         # python -m entry point
│   ├── cli.py              # Typer CLI (invoice-extract command)
│   ├── exporter.py         # Excel export with openpyxl
│   ├── extractor.py        # PDF parsing + field extraction
│   ├── patterns.py         # Regex patterns for Italian invoices
│   └── web.py              # Gradio web UI (optional)
├── config.json             # Example configuration
├── requirements.txt        # Python dependencies
└── README.md               # This file
```

## Configuration

Customize extraction patterns in `config.json`:

```json
{
  "patterns": {
    "invoice_number": ["your custom regex here"],
    ...
  },
  "date_formats": ["%d/%m/%Y", ...]
}
```

Pass custom config: `invoice-extract ./invoices/ -c my_config.json`

## Requirements

- Python 3.10+
- pdfplumber — PDF text extraction
- openpyxl — Excel export
- typer — CLI framework
- rich — pretty terminal output
- gradio — optional web UI

## License

MIT
