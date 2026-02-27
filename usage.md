# How to Use Script with Different Files
## Quick Start (Default Setup)
Just place your files in the source directory and run:

**Add new files to this directory:**

/mnt/r/money_record/auto_insertion_source_files/

**Run the script**

python3 insert_transactions_by_date.py

That's it. The script automatically:

- Discovers all supported files
- Parses them according to their format
- Inserts transactions grouped by date
- Creates a backup before overwriting
---
Supported File Types
| Extension                                                    | Format                         | Parser Used                 | Requirements                    |
| ------------------------------------------------------------ | ------------------------------ | --------------------------- | ------------------------------- |
| .csv                                                         | JD/Alipay exports              | Header-based auto-detection | None                            |
| .xlsx                                                        | WeChat exports                 | First worksheet             | None                            |
| .xls                                                         | Legacy Excel (bank statements) | xlrd                        | pip install -r requirements.txt |
| .pdf                                                         | Bank statements (text-only)    | pdfplumber                  | pip install -r requirements.txt |
| Note: CSV files are sniffed for known schemas (JD vs Alipay). Unknown CSV schemas are skipped with warning. |                                |                             |                                 |
---
## Optional: Custom Paths

**Different source directory**

python3 insert_transactions_by_date.py --source-dir /path/to/your/files

**Different workbook**

python3 insert_transactions_by_date.py --workbook /path/to/output.xlsx

**Different sheet name**

python3 insert_transactions_by_date.py --sheet "MySheet"

## Best Practices

1. Install optional dependencies once:
      pip install -r requirements.txt
   
2. Check dependency availability:
      python3 insert_transactions_by_date.py --deps-check
   
3. Test before writing (dry-run):
      python3 insert_transactions_by_date.py --dry-run
   
4. Re-running is safe:
   - Same files are skipped via tx_fingerprint (idempotent)
   - Only new transactions are inserted
---
## Adding New Files

1. Copy files to /mnt/r/money_record/auto_insertion_source_files/
2. Run script:
      python3 insert_transactions_by_date.py
   3. Check output:
   - FILE_RESULT: lines show per-file parsing status
   - TOTAL: line shows summary
   - BACKUP_CREATED: shows backup location (when --in-place)
---
## What Gets Parsed
Each transaction must have:

- Valid date (normalized to YYYY-MM-DD)
- Valid amount (float with sign)
What gets inserted:
- Column A: 日期
- Column D: 金额
- Column F: 支付方式
- Column G: fp:<fingerprint> (for idempotency)
- Column H+: Extra fields (merchant, order_id, raw columns, etc.)
---
## Troubleshooting
**File not parsed? Check warnings:**
python3 insert_transactions_by_date.py --dry-run

Look for SKIP_* messages

**Missing dependencies?**
python3 insert_transactions_by_date.py --deps-check

Install: pip install -r requirements.txt

**Need to see what will happen without writing?**
python3 insert_transactions_by_date.py --dry-run