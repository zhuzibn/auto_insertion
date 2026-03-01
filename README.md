# Auto Insertion - Transaction Data Import Tool

Python CLI tool that parses transaction data from multiple platform exports (JD, WeChat, Alipay, XLS, PDF) and inserts them into Excel worksheets grouped by date.

## Features

- **Multi-format support**: Parses exports from JD, WeChat, Alipay, XLS, and PDF files
- **Date-based grouping**: Automatically organizes transactions by date in Excel worksheets  
- **Idempotent inserts**: Uses SHA-1 fingerprinting to avoid duplicate entries
- **Flexible column mapping**: Supports dynamic columns beyond Excel's column Z limit
- **Robust parsing**: Handles various date formats and numeric conventions
- **Legacy repair**: Fix previously inserted positive expense amounts with `--repair-jd-legacy-sign`

## Installation

```bash
# Clone the repository
git clone <repository-url>
cd auto_insertion

# Install optional dependencies (for PDF and legacy XLS support)
pip install -r requirements.txt
```

## Usage

### Basic Usage

```bash
python3 insert_transactions_by_date.py
```

### Advanced Options

```bash
# Custom source directory and workbook
python3 insert_transactions_by_date.py \
  --source-dir /path/to/source/files \
  --workbook /path/to/workbook.xlsx \
  --sheet "Sheet Name"

# Parse only (dry run)
python3 insert_transactions_by_date.py --dry-run

# Check dependencies
python3 insert_transactions_by_date.py --deps-check

# Repair legacy JD expense sign issues
python3 insert_transactions_by_date.py --repair-jd-legacy-sign
```

### CLI Arguments

| Argument | Default | Description |
|----------|---------|-------------|
| `--source-dir` | `/mnt/r/money_record/auto_insertion_source_files` | Directory containing transaction export files |
| `--workbook` | `/mnt/r/money_record/money_2026.xlsx` | Excel workbook to modify |
| `--sheet` | `2026每日饮食表` | Worksheet name to insert transactions |
| `--in-place` | `True` | Overwrite workbook with backup |
| `--dry-run` | `False` | Parse only, no write operations |
| `--deps-check` | - | Check optional dependency availability |
| `--repair-jd-legacy-sign` | `False` | Fix legacy positive JD expense entries |
| `--selftest-roundtrip-inline-str` | - | Run inline string roundtrip self-test |
| `--selftest-dynamic-cols` | - | Run dynamic column handling self-test |

## Testing

```bash
# Run all tests
python3 -m unittest -v

# Run specific test file
python3 -m unittest tests.test_helpers

# Run specific test class
python3 -m unittest tests.test_helpers.ColumnLetterIndexTests -v

# Run specific test method
python3 -m unittest tests.test_helpers.ColumnLetterIndexTests.test_col_index_to_letter_boundaries -v
```

## Self-Tests

```bash
# Test inline string roundtrip
python3 insert_transactions_by_date.py --selftest-roundtrip-inline-str

# Test dynamic column handling
python3 insert_transactions_by_date.py --selftest-dynamic-cols
```

## Project Structure

```
./
├── insert_transactions_by_date.py    # Main CLI script
├── requirements.txt                 # Optional dependencies
├── tests/
│   ├── __init__.py
│   └── test_helpers.py              # Unit tests for helper functions
└── README.md                        # This file
```

## Supported Platforms

- **JD (京东)**: CSV exports with transaction data (supports both `收支` and `收/支` headers)
- **WeChat (微信)**: Transaction records
- **Alipay (支付宝)**: Payment history exports  
- **XLS**: Legacy Excel format (requires xlrd==1.2.0)
- **PDF**: Text-based transaction records (requires pdfplumber)

## How It Works

1. **File Discovery**: Scans source directory for supported file formats
2. **Parser Dispatch**: Classifies CSV files and routes to appropriate parser
3. **Transaction Parsing**: Extracts structured transaction data (date, amount, payment method, etc.)
4. **Fingerprinting**: Computes SHA-1 hash for each transaction to detect duplicates
5. **Excel Insertion**: Writes transactions to worksheet, skipping existing entries
6. **Backup**: Creates `.backup` of original workbook before modifications

## Excel Column Mapping

| Column | Content |
|--------|---------|
| A | 日期 (Date) |
| D | 金额 (Amount) - stored as numeric value |
| F | 支付方式 (Payment Method) |
| G | tx_fingerprint (SHA-1 hash) |
| H+ | Extra fields (dynamic platform-specific data) |

## Transaction Format

Each transaction contains:
- `date`: Normalized to `YYYY-MM-DD`
- `amount`: Formatted with 2 decimals (expenses negative, income positive)
- `payment_method`: Payment channel
- `merchant`: Business name (if available)
- `order_id`: Transaction identifier (if available)  
- `source_file`: Original filename
- `source_row`: Row number in source file
- `extra_fields`: Additional platform-specific data

## Optional Dependencies

Install via: `pip install -r requirements.txt`

- **`pdfplumber`**: Required for PDF text extraction and transaction parsing
- **`xlrd==1.2.0`**: Required for legacy .xls file support (specific version for compatibility)

Core functionality works without these dependencies - they're only needed for PDF and legacy XLS file support.