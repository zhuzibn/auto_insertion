#!/usr/bin/env python3
import argparse
import csv
import datetime as dt
import hashlib
import importlib.util
import os
import re
import shutil
import sys
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Protocol, cast
import xml.etree.ElementTree as ET


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_DOC_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships"


@dataclass
class Tx:
    date: str
    amount: float
    payment_method: str
    platform: str
    merchant: str
    order_id: str
    source_file: str
    source_path: str
    source_row: str
    extra_fields: list[tuple[str, str]]


class TxLike(Protocol):
    date: str
    amount: float
    payment_method: str
    platform: str
    merchant: str
    order_id: str
    source_file: str
    source_row: str


class CliArgs(Protocol):
    source_dir: str
    workbook: str
    sheet: str
    in_place: bool
    dry_run: bool
    deps_check: bool
    selftest_roundtrip_inline_str: bool
    selftest_dynamic_cols: bool


def dep_available(name: str) -> bool:
    return importlib.util.find_spec(name) is not None


def col_index_to_letter(index: int) -> str:
    if index <= 0:
        raise ValueError("column index must be >= 1")
    chars: list[str] = []
    n = index
    while n > 0:
        n, rem = divmod(n - 1, 26)
        chars.append(chr(ord("A") + rem))
    return "".join(reversed(chars))


def col_letter_to_index(letter: str) -> int:
    token = (letter or "").strip().upper()
    if not re.fullmatch(r"[A-Z]+", token):
        raise ValueError(f"invalid column letter: {letter}")
    out = 0
    for ch in token:
        out = out * 26 + (ord(ch) - ord("A") + 1)
    return out


def normalize_date(value: object | None) -> str | None:
    if value is None:
        return None
    text = str(value).strip()
    if not text:
        return None

    text = (
        text.replace("年", "-")
        .replace("月", "-")
        .replace("日", "")
        .replace("/", "-")
        .replace(".", "-")
    )
    text = re.sub(r"\s+", " ", text)
    for fmt in (
        "%Y-%m-%d",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%Y%m%d",
        "%Y-%m-%dT%H:%M:%S",
    ):
        try:
            return dt.datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            pass

    m = re.match(r"^(\d{4}-\d{1,2}-\d{1,2})", text)
    if m:
        try:
            return dt.datetime.strptime(m.group(1), "%Y-%m-%d").strftime("%Y-%m-%d")
        except ValueError:
            return None
    m = re.match(r"^(\d{8})", text)
    if m:
        try:
            return dt.datetime.strptime(m.group(1), "%Y%m%d").strftime("%Y-%m-%d")
        except ValueError:
            return None
    return None


def parse_amount(raw: object | None) -> float | None:
    if raw is None:
        return None
    text = str(raw).strip().replace(",", "")
    if not text:
        return None
    m = re.search(r"[-+]?\d+(?:\.\d+)?", text)
    if not m:
        return None
    try:
        return float(m.group(0))
    except ValueError:
        return None


def csv_rows(file_path: str) -> list[list[str]]:
    for enc in ("utf-8-sig", "gbk"):
        try:
            with open(file_path, "r", encoding=enc, newline="") as f:
                return list(csv.reader(f))
        except Exception:  # noqa: BLE001
            pass
    with open(file_path, "r", encoding="gbk", errors="replace", newline="") as f:
        rows = list(csv.reader(f))
    if rows:
        return rows
    raise RuntimeError(f"Cannot read CSV: {file_path}")


def classify_csv(file_path: str) -> str:
    rows = csv_rows(file_path)
    sample = "\n".join(",".join(row) for row in rows[:40])
    if "京东交易流水" in sample or "京东账号名" in sample:
        return "jd_csv"
    if (
        "支付宝账户" in sample
        or "支付宝交易明细" in sample
        or "交易创建时间" in sample
        or "交易订单号" in sample
        or ("收/支" in sample and "交易对方" in sample)
    ):
        return "alipay_csv"
    return "unknown_csv"


def parse_jd_csv(file_path: str) -> tuple[list[Tx], list[str], dict[str, int]]:
    rows = csv_rows(file_path)
    if not rows:
        return (
            [],
            [f"SKIP_EMPTY_FILE:{file_path}"],
            {"rows_parsed": 0, "rows_skipped": 0},
        )

    header_idx = 0
    for i, row in enumerate(rows[:40]):
        blob = ",".join(row)
        if "交易时间" in blob and "金额" in blob:
            header_idx = i
            break

    header = rows[header_idx]
    idx = {h.strip(): i for i, h in enumerate(header)}

    txs: list[Tx] = []
    warnings: list[str] = []
    skipped = 0

    for row_no, row in enumerate(rows[header_idx + 1 :], start=header_idx + 2):
        if not any(str(v).strip() for v in row):
            continue

        def cell(name: str) -> str:
            col = idx.get(name, -1)
            return row[col] if 0 <= col < len(row) else ""

        date = normalize_date(cell("交易时间"))
        amount = parse_amount(cell("金额"))
        if date is None or amount is None:
            skipped += 1
            missing: list[str] = []
            if date is None:
                missing.append("date")
            if amount is None:
                missing.append("amount")
            warnings.append(
                f"SKIP_JD_INVALID_ROW:{file_path}:{row_no}:missing={','.join(missing)}"
            )
            continue

        direction = cell("收支")
        if "支出" in direction and amount > 0:
            amount = -amount
        if "收入" in direction and amount < 0:
            amount = abs(amount)

        merchant = cell("交易对方")
        order_id = cell("订单号")
        txs.append(
            Tx(
                date=date,
                amount=amount,
                payment_method=cell("支付方式") or "京东",
                platform="jd",
                merchant=merchant,
                order_id=order_id,
                source_file=os.path.basename(file_path),
                source_path=file_path,
                source_row=str(row_no),
                extra_fields=[
                    ("merchant", merchant),
                    ("order_id", order_id),
                    ("tx_type", cell("交易类型")),
                ],
            )
        )

    return txs, warnings, {"rows_parsed": len(txs), "rows_skipped": skipped}


def parse_alipay_csv(file_path: str) -> tuple[list[Tx], list[str], dict[str, int]]:
    rows = csv_rows(file_path)
    if not rows:
        return (
            [],
            [f"SKIP_EMPTY_FILE:{file_path}"],
            {"rows_parsed": 0, "rows_skipped": 0},
        )

    header_idx = 0
    for i, row in enumerate(rows[:40]):
        row_set = {x.strip() for x in row}
        if ("交易创建时间" in row_set or "交易时间" in row_set) and (
            "金额" in row_set or "金额（元）" in row_set
        ):
            header_idx = i
            break

    header = rows[header_idx]
    idx = {h.strip(): i for i, h in enumerate(header)}

    def first_key(*options: str) -> str:
        for option in options:
            if option in idx:
                return option
        return ""

    date_key = first_key("交易创建时间", "交易时间")
    amount_key = first_key("金额（元）", "金额")
    pay_key = first_key("付款方式", "收/付款方式")
    order_key = first_key("订单号", "交易订单号")

    txs: list[Tx] = []
    warnings: list[str] = []
    skipped = 0

    for row_no, row in enumerate(rows[header_idx + 1 :], start=header_idx + 2):
        if not any(str(v).strip() for v in row):
            continue

        def cell(name: str) -> str:
            col = idx.get(name, -1)
            return row[col] if 0 <= col < len(row) else ""

        date = normalize_date(cell(date_key))
        amount = parse_amount(cell(amount_key))
        if date is None or amount is None:
            skipped += 1
            missing: list[str] = []
            if date is None:
                missing.append("date")
            if amount is None:
                missing.append("amount")
            warnings.append(
                f"SKIP_ALIPAY_INVALID_ROW:{file_path}:{row_no}:missing={','.join(missing)}"
            )
            continue

        direction = cell("收/支")
        if "支" in direction and amount > 0:
            amount = -amount
        if "收" in direction and amount < 0:
            amount = abs(amount)

        merchant = cell("交易对方")
        order_id = cell(order_key)
        txs.append(
            Tx(
                date=date,
                amount=amount,
                payment_method=cell(pay_key) or "支付宝",
                platform="alipay",
                merchant=merchant,
                order_id=order_id,
                source_file=os.path.basename(file_path),
                source_path=file_path,
                source_row=str(row_no),
                extra_fields=[("merchant", merchant), ("order_id", order_id)],
            )
        )

    return txs, warnings, {"rows_parsed": len(txs), "rows_skipped": skipped}


def read_shared_strings(zf: zipfile.ZipFile) -> list[str]:
    try:
        root = ET.fromstring(zf.read("xl/sharedStrings.xml"))
    except KeyError:
        return []

    strings: list[str] = []
    for si in root.findall(f"{{{NS_MAIN}}}si"):
        parts: list[str] = []
        t = si.find(f"{{{NS_MAIN}}}t")
        if t is not None and t.text is not None:
            parts.append(t.text)
        for run in si.findall(f"{{{NS_MAIN}}}r"):
            rt = run.find(f"{{{NS_MAIN}}}t")
            if rt is not None and rt.text is not None:
                parts.append(rt.text)
        strings.append("".join(parts))
    return strings


def parse_sheet_xml(
    xml_bytes: bytes, shared_strings: list[str]
) -> dict[int, dict[str, str]]:
    root = ET.fromstring(xml_bytes)
    sheet_data = root.find(f"{{{NS_MAIN}}}sheetData")
    out: dict[int, dict[str, str]] = {}
    if sheet_data is None:
        return out

    for row in sheet_data.findall(f"{{{NS_MAIN}}}row"):
        row_num = int(row.attrib.get("r", "0"))
        row_cells: dict[str, str] = out.setdefault(row_num, {})
        for cell in row.findall(f"{{{NS_MAIN}}}c"):
            ref = cell.attrib.get("r", "")
            col = re.sub(r"\d", "", ref)
            if not col:
                continue

            ctype = cell.attrib.get("t", "")
            val = ""
            if ctype == "s":
                v = cell.find(f"{{{NS_MAIN}}}v")
                if v is not None and v.text is not None:
                    try:
                        idx = int(v.text)
                    except ValueError:
                        idx = -1
                    if 0 <= idx < len(shared_strings):
                        val = shared_strings[idx]
            elif ctype == "inlineStr":
                t = cell.find(f"{{{NS_MAIN}}}is/{{{NS_MAIN}}}t")
                if t is not None and t.text is not None:
                    val = t.text
            else:
                v = cell.find(f"{{{NS_MAIN}}}v")
                if v is not None and v.text is not None:
                    val = v.text

            if val != "":
                row_cells[col] = val

    return out


def parse_wechat_xlsx(file_path: str) -> tuple[list[Tx], list[str], dict[str, int]]:
    try:
        with zipfile.ZipFile(file_path, "r") as zf:
            shared = read_shared_strings(zf)
            sheet_xml = zf.read("xl/worksheets/sheet1.xml")
    except KeyError as exc:
        return (
            [],
            [f"SKIP_XLSX_READ_ERROR:{file_path}:{exc}"],
            {"rows_parsed": 0, "rows_skipped": 0},
        )
    except Exception as exc:  # noqa: BLE001
        return (
            [],
            [f"SKIP_XLSX_READ_ERROR:{file_path}:{exc}"],
            {"rows_parsed": 0, "rows_skipped": 0},
        )

    rows_data = parse_sheet_xml(sheet_xml, shared)
    header_row = None
    header_map: dict[str, str] = {}
    for row_num in sorted(rows_data):
        vals = rows_data[row_num]
        if any("交易时间" in str(v) for v in vals.values()):
            header_row = row_num
            header_map = {str(v).strip(): c for c, v in vals.items()}
            break

    if header_row is None:
        return (
            [],
            [f"SKIP_XLSX_UNKNOWN_SCHEMA:{file_path}"],
            {"rows_parsed": 0, "rows_skipped": 0},
        )

    def get_val(row: dict[str, str], key: str) -> str:
        col = header_map.get(key)
        return row.get(col, "") if col else ""

    txs: list[Tx] = []
    warnings: list[str] = []
    skipped = 0
    for row_num in sorted(k for k in rows_data if k > header_row):
        row = rows_data[row_num]
        date = normalize_date(get_val(row, "交易时间"))
        amount = parse_amount(get_val(row, "金额(元)") or get_val(row, "金额"))
        if date is None or amount is None:
            if any(row.values()):
                skipped += 1
                missing: list[str] = []
                if date is None:
                    missing.append("date")
                if amount is None:
                    missing.append("amount")
                warnings.append(
                    f"SKIP_WECHAT_INVALID_ROW:{file_path}:{row_num}:missing={','.join(missing)}"
                )
            continue

        direction = get_val(row, "收/支")
        if "支出" in direction and amount > 0:
            amount = -amount
        if "收入" in direction and amount < 0:
            amount = abs(amount)

        merchant = get_val(row, "交易对方")
        order_id = get_val(row, "商户单号") or get_val(row, "交易单号")
        txs.append(
            Tx(
                date=date,
                amount=amount,
                payment_method=get_val(row, "支付方式") or "微信",
                platform="wechat",
                merchant=merchant,
                order_id=order_id,
                source_file=os.path.basename(file_path),
                source_path=file_path,
                source_row=str(row_num),
                extra_fields=[("merchant", merchant), ("order_id", order_id)],
            )
        )

    return txs, warnings, {"rows_parsed": len(txs), "rows_skipped": skipped}


def parse_xls_transactions(
    file_path: str,
) -> tuple[list[Tx], list[str], dict[str, int]]:
    try:
        import xlrd  # type: ignore[import-not-found]
    except Exception:  # noqa: BLE001
        return (
            [],
            [f"SKIP_XLS_MISSING_DEP:{file_path}"],
            {"rows_parsed": 0, "rows_skipped": 0},
        )

    try:
        book = xlrd.open_workbook(file_path)
        sheet = book.sheet_by_index(0)
    except Exception as exc:  # noqa: BLE001
        return (
            [],
            [f"SKIP_XLS_READ_ERROR:{file_path}:{exc}"],
            {"rows_parsed": 0, "rows_skipped": 0},
        )

    nrows = int(sheet.nrows)
    ncols = int(sheet.ncols)

    def cell_text(r: int, c: int) -> str:
        if c < 0 or c >= ncols or r < 0 or r >= nrows:
            return ""
        ctype = sheet.cell_type(r, c)
        value = sheet.cell_value(r, c)
        if ctype in (xlrd.XL_CELL_EMPTY, xlrd.XL_CELL_BLANK):
            return ""
        if ctype == xlrd.XL_CELL_DATE:
            try:
                as_dt = dt.datetime(*xlrd.xldate_as_tuple(float(value), book.datemode))
                return as_dt.strftime("%Y-%m-%d %H:%M:%S")
            except Exception:  # noqa: BLE001
                return str(value).strip()
        if ctype == xlrd.XL_CELL_NUMBER:
            try:
                num = float(value)
            except Exception:  # noqa: BLE001
                return str(value).strip()
            if num.is_integer():
                return str(int(num))
            return str(num)
        return str(value).strip()

    def norm_header(text: str) -> str:
        return re.sub(r"\s+", "", text)

    header_row = None
    header_values: list[str] = []
    for r in range(min(nrows, 80)):
        vals = [cell_text(r, c) for c in range(ncols)]
        normalized = {norm_header(v) for v in vals if v}
        has_date = "交易日期" in normalized
        has_amount = ("交易金额" in normalized) or ("发生额" in normalized)
        if has_date and has_amount:
            header_row = r
            header_values = vals
            break

    if header_row is None:
        return (
            [],
            [f"SKIP_XLS_UNKNOWN_SCHEMA:{file_path}"],
            {"rows_parsed": 0, "rows_skipped": 0},
        )

    header_idx = {
        norm_header(v): i for i, v in enumerate(header_values) if norm_header(v)
    }

    date_col = header_idx.get("交易日期", -1)
    amount_col = header_idx.get("交易金额", header_idx.get("发生额", -1))

    if date_col < 0 or amount_col < 0:
        return (
            [],
            [f"SKIP_XLS_UNKNOWN_SCHEMA:{file_path}"],
            {"rows_parsed": 0, "rows_skipped": 0},
        )

    merchant_cols: list[int] = []
    for key in ("摘要", "对方账号与户名", "交易地点/附言"):
        col = header_idx.get(key)
        if col is not None:
            merchant_cols.append(col)

    txs: list[Tx] = []
    skipped = 0
    for r in range(header_row + 1, nrows):
        values = [cell_text(r, c) for c in range(ncols)]
        if not any(values):
            continue
        if any("合计" in v or "总计" in v for v in values if v):
            skipped += 1
            continue

        raw_date = values[date_col] if date_col < len(values) else ""
        raw_amount = values[amount_col] if amount_col < len(values) else ""
        date = normalize_date(raw_date)
        amount = parse_amount(raw_amount)
        if date is None or amount is None:
            skipped += 1
            continue

        merchant = ""
        for c in merchant_cols:
            if c < len(values) and values[c]:
                merchant = values[c]
                break

        extras = [("merchant", merchant), ("order_id", "")]
        for name, col in sorted(header_idx.items(), key=lambda it: it[1]):
            extras.append((f"xls_{name}", values[col] if col < len(values) else ""))

        txs.append(
            Tx(
                date=date,
                amount=amount,
                payment_method="建行",
                platform="xls",
                merchant=merchant,
                order_id="",
                source_file=os.path.basename(file_path),
                source_path=file_path,
                source_row=str(r + 1),
                extra_fields=extras,
            )
        )

    return txs, [], {"rows_parsed": len(txs), "rows_skipped": skipped}


def infer_pdf_payment_method(file_path: str) -> str:
    name = os.path.basename(file_path)
    if "上海银行" in name:
        return "上海银行"
    if "农业银行" in name:
        return "农业银行"
    if "广发" in name:
        return "广发银行"
    if "工商银行" in name or "工行" in name:
        return "工商银行"
    return "PDF"


def parse_pdf_transactions(
    file_path: str,
) -> tuple[list[Tx], list[str], dict[str, int]]:
    try:
        import pdfplumber  # type: ignore[import-not-found]
    except Exception:  # noqa: BLE001
        return (
            [],
            [f"SKIP_PDF_MISSING_DEP:{file_path}"],
            {"rows_parsed": 0, "rows_skipped": 0},
        )

    date_line_re = re.compile(r"^(\d{4}-\d{2}-\d{2}|\d{8})")
    amount_token_re = re.compile(r"[-+]?\d[\d,]*\.\d{2}")
    skip_keywords = ("合计", "笔数", "期初余额", "期末余额", "最低还款", "账单")
    source_file = os.path.basename(file_path)
    payment_method = infer_pdf_payment_method(file_path)

    txs: list[Tx] = []
    warnings: list[str] = []
    skipped = 0
    source_row = 0

    try:
        with pdfplumber.open(file_path) as pdf:
            page_lines: list[tuple[str, bool]] = []
            for page in pdf.pages:
                plain_text = page.extract_text() or ""
                layout_text = page.extract_text(layout=True) or ""
                header_blob = plain_text + "\n" + layout_text
                has_shanghai_header = "交易金额" in header_blob and (
                    "交易日期" in header_blob or "记账日期" in header_blob
                )
                selected_text = layout_text if layout_text.strip() else plain_text
                for raw_line in selected_text.splitlines():
                    page_lines.append((raw_line.rstrip("\n"), has_shanghai_header))

            for raw_line, shanghai_mode in page_lines:
                source_row += 1
                line = raw_line.strip()
                if not line:
                    continue
                if (not date_line_re.match(line)) and any(
                    k in line for k in skip_keywords
                ):
                    continue
                if not date_line_re.match(line):
                    continue

                amount_tokens = amount_token_re.findall(line)
                if not amount_tokens:
                    skipped += 1
                    warnings.append(f"SKIP_PDF_AMBIGUOUS_AMOUNT:{file_path}:{line}")
                    continue

                date = normalize_date(line)
                amount = parse_amount(cast(str, amount_tokens[0]))
                if date is None or amount is None:
                    skipped += 1
                    warnings.append(f"SKIP_PDF_AMBIGUOUS_AMOUNT:{file_path}:{line}")
                    continue

                memo = ""
                merchant = ""
                if shanghai_mode:
                    cols = re.split(r"\s{2,}", line)
                    if len(cols) >= 5:
                        head_parts = cols[0].split(None, 1)
                        if len(head_parts) >= 2:
                            memo = head_parts[1].strip()
                    if len(cols) >= 6:
                        merchant = cols[4].strip()
                    elif len(cols) >= 5:
                        merchant = cols[4].strip()

                if not memo:
                    m = re.match(
                        r"^(\d{4}-\d{2}-\d{2}|\d{8})(?:\s+\d{2}:\d{2}(?::\d{2})?)?\s*(.*)$",
                        line,
                    )
                    memo = (m.group(2) if m else line).strip()

                if not merchant:
                    amount_match = amount_token_re.search(memo)
                    if amount_match and memo and not memo[0].isdigit():
                        memo = memo[: amount_match.start()].strip()

                merchant = merchant or memo
                txs.append(
                    Tx(
                        date=date,
                        amount=amount,
                        payment_method=payment_method,
                        platform="pdf",
                        merchant=merchant,
                        order_id="",
                        source_file=source_file,
                        source_path=file_path,
                        source_row=str(source_row),
                        extra_fields=[
                            ("merchant", merchant),
                            ("order_id", ""),
                            ("raw_line", line),
                        ],
                    )
                )
    except Exception as exc:  # noqa: BLE001
        return (
            [],
            [f"SKIP_PDF_READ_ERROR:{file_path}:{exc}"],
            {"rows_parsed": 0, "rows_skipped": 0},
        )

    return txs, warnings, {"rows_parsed": len(txs), "rows_skipped": skipped}


def parser_dispatch(
    file_path: str,
    has_pdf: bool,
    has_xlrd: bool,
) -> tuple[str, list[Tx], list[str], dict[str, int]]:
    ext = Path(file_path).suffix.lower()
    if ext == ".csv":
        kind = classify_csv(file_path)
        if kind == "jd_csv":
            txs, warnings, stats = parse_jd_csv(file_path)
            return "jd_csv", txs, warnings, stats
        if kind == "alipay_csv":
            txs, warnings, stats = parse_alipay_csv(file_path)
            return "alipay_csv", txs, warnings, stats
        return (
            "unknown_csv",
            [],
            [f"SKIP_CSV_UNKNOWN_SCHEMA:{file_path}"],
            {"rows_parsed": 0, "rows_skipped": 0},
        )

    if ext == ".xlsx":
        txs, warnings, stats = parse_wechat_xlsx(file_path)
        return "wechat_xlsx", txs, warnings, stats

    if ext == ".xls":
        if not has_xlrd:
            return (
                "xls",
                [],
                [f"SKIP_XLS_MISSING_DEP:{file_path}"],
                {"rows_parsed": 0, "rows_skipped": 0},
            )
        txs, warnings, stats = parse_xls_transactions(file_path)
        return "xls", txs, warnings, stats

    if ext == ".pdf":
        if not has_pdf:
            return (
                "pdf",
                [],
                [f"SKIP_PDF_MISSING_DEP:{file_path}"],
                {"rows_parsed": 0, "rows_skipped": 0},
            )
        txs, warnings, stats = parse_pdf_transactions(file_path)
        return "pdf", txs, warnings, stats

    return (
        "unsupported",
        [],
        [f"SKIP_UNSUPPORTED:{file_path}"],
        {"rows_parsed": 0, "rows_skipped": 0},
    )


def resolve_sheet_path(workbook_path: str, sheet_name: str) -> tuple[str, list[str]]:
    with zipfile.ZipFile(workbook_path, "r") as zf:
        wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
        rels_root = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))

    rel_by_id: dict[str, str] = {}
    for rel in rels_root.findall(f"{{{NS_PKG_REL}}}Relationship"):
        rel_id = rel.attrib.get("Id", "")
        target = rel.attrib.get("Target", "")
        if rel_id and target:
            rel_by_id[rel_id] = target

    sheets = wb_root.find(f"{{{NS_MAIN}}}sheets")
    if sheets is None:
        raise RuntimeError("Workbook has no sheets")

    available: list[str] = []
    for sheet in sheets.findall(f"{{{NS_MAIN}}}sheet"):
        name = sheet.attrib.get("name", "")
        if name:
            available.append(name)

        rid = sheet.attrib.get(f"{{{NS_DOC_REL}}}id", "")
        if name == sheet_name:
            target = rel_by_id.get(rid)
            if not target:
                raise RuntimeError(
                    f"Missing relationship target for sheet: {sheet_name}"
                )
            if target.startswith("/"):
                return target.lstrip("/"), available
            return "xl/" + target.replace("\\", "/"), available

    raise ValueError(
        f"Sheet '{sheet_name}' not found. Available sheets: {', '.join(available)}"
    )


def read_sheet_structure(
    workbook_path: str, sheet_xml_path: str
) -> dict[int, dict[str, str]]:
    with zipfile.ZipFile(workbook_path, "r") as zf:
        shared_strings = read_shared_strings(zf)
        xml_bytes = zf.read(sheet_xml_path)
    return parse_sheet_xml(xml_bytes, shared_strings)


def create_worksheet_xml(rows_data: dict[int, dict[str, str]]) -> bytes:
    root = ET.Element("worksheet", xmlns=NS_MAIN)

    max_row = max(rows_data) if rows_data else 1
    max_col_idx = 6
    for row in rows_data.values():
        for col in row:
            col_idx = col_letter_to_index(col)
            if col_idx > max_col_idx:
                max_col_idx = col_idx

    max_col = col_index_to_letter(max_col_idx)
    _ = ET.SubElement(root, "dimension", ref=f"A1:{max_col}{max_row}")

    cols = ET.SubElement(root, "cols")
    for i in range(1, 7):
        _ = ET.SubElement(
            cols, "col", min=str(i), max=str(i), width="12", customWidth="1"
        )
    if max_col_idx >= 7:
        _ = ET.SubElement(
            cols,
            "col",
            min="7",
            max=str(max_col_idx),
            width="20",
            customWidth="1",
        )

    sheet_data = ET.SubElement(root, "sheetData")
    for row_no in sorted(rows_data):
        row_el = ET.SubElement(sheet_data, "row", r=str(row_no))
        for _, col in sorted((col_letter_to_index(c), c) for c in rows_data[row_no]):
            cell = ET.SubElement(row_el, "c", r=f"{col}{row_no}", t="inlineStr")
            is_el = ET.SubElement(cell, "is")
            t_el = ET.SubElement(is_el, "t")
            t_el.text = str(rows_data[row_no][col])

    return cast(bytes, ET.tostring(root, encoding="utf-8", xml_declaration=True))


def tx_fingerprint(tx: TxLike) -> str:
    payload = "|".join(
        [
            tx.platform,
            tx.date,
            f"{tx.amount:.2f}",
            tx.payment_method,
            tx.merchant,
            tx.order_id,
            tx.source_file,
            tx.source_row,
        ]
    )
    return "fp:" + hashlib.sha1(payload.encode("utf-8")).hexdigest()


def index_existing_fingerprints(rows_data: dict[int, dict[str, str]]) -> set[str]:
    found: set[str] = set()
    fp_re = re.compile(r"fp:[0-9a-f]{40}")
    for row in rows_data.values():
        g = str(row.get("G", "")).strip()
        if g.startswith("fp:"):
            found.add(g)
        for fp in cast(list[str], fp_re.findall(g)):
            found.add(fp)
    return found


def build_date_blocks(
    rows_data: dict[int, dict[str, str]],
) -> dict[str, dict[str, int]]:
    if not rows_data:
        return {}

    max_row = max(rows_data)
    blocks: dict[str, dict[str, int]] = {}
    current_date: str | None = None
    current_start: int | None = None

    for row_no in range(2, max_row + 1):
        raw = str(rows_data.get(row_no, {}).get("A", "")).strip()
        date = normalize_date(raw) if raw else None
        if date:
            if date == current_date:
                continue
            if current_date is not None and current_start is not None:
                blocks[current_date] = {"start": current_start, "end": row_no - 1}
            current_date = date
            current_start = row_no

    if current_date is not None and current_start is not None:
        blocks[current_date] = {"start": current_start, "end": max_row}
    return blocks


def insert_blank_row(rows_data: dict[int, dict[str, str]], row_idx: int) -> None:
    for row_no in sorted(list(rows_data.keys()), reverse=True):
        if row_no >= row_idx:
            rows_data[row_no + 1] = rows_data.pop(row_no)
    rows_data[row_idx] = {}


def write_tx_to_row(
    rows_data: dict[int, dict[str, str]], row_no: int, tx: Tx, fp: str
) -> None:
    row = rows_data.setdefault(row_no, {})
    row["A"] = tx.date
    row["D"] = f"{tx.amount:.2f}"
    row["F"] = tx.payment_method
    row["G"] = fp

    extras = list(tx.extra_fields) + [
        ("source_file", tx.source_file),
        ("source_path", tx.source_path),
        ("source_type", Path(tx.source_file).suffix.lower().lstrip(".")),
    ]
    col_idx = 8
    for key, value in extras:
        row[col_index_to_letter(col_idx)] = f"{key}={value}"
        col_idx += 1


def find_insert_position_for_new_date(
    rows_data: dict[int, dict[str, str]], tx_date: str
) -> int:
    if not rows_data:
        return 2
    max_row = max(rows_data)
    for row_no in range(2, max_row + 1):
        raw = str(rows_data.get(row_no, {}).get("A", "")).strip()
        date = normalize_date(raw) if raw else None
        if date and tx_date < date:
            return row_no
    return max_row + 1


def insert_transactions(
    rows_data: dict[int, dict[str, str]], txs: list[Tx]
) -> tuple[int, int]:
    blocks = build_date_blocks(rows_data)
    existing_fps = index_existing_fingerprints(rows_data)
    inserted = 0
    dup_skipped = 0

    for tx in sorted(
        txs, key=lambda item: (item.date, item.source_file, item.source_row)
    ):
        fp = tx_fingerprint(tx)
        if fp in existing_fps:
            dup_skipped += 1
            continue

        if tx.date in blocks:
            insert_at = blocks[tx.date]["end"] + 1
            insert_blank_row(rows_data, insert_at)
            write_tx_to_row(rows_data, insert_at, tx, fp)
            for date_key in list(blocks.keys()):
                if date_key == tx.date:
                    blocks[date_key]["end"] += 1
                elif blocks[date_key]["start"] >= insert_at:
                    blocks[date_key]["start"] += 1
                    blocks[date_key]["end"] += 1
        else:
            date_row = find_insert_position_for_new_date(rows_data, tx.date)
            insert_blank_row(rows_data, date_row)
            rows_data[date_row]["A"] = tx.date

            insert_blank_row(rows_data, date_row + 1)
            write_tx_to_row(rows_data, date_row + 1, tx, fp)

            for date_key in list(blocks.keys()):
                if blocks[date_key]["start"] >= date_row:
                    blocks[date_key]["start"] += 2
                    blocks[date_key]["end"] += 2
            blocks[tx.date] = {"start": date_row, "end": date_row + 1}

        existing_fps.add(fp)
        inserted += 1

    return inserted, dup_skipped


def discover_source_files(source_dir: str) -> tuple[list[str], list[str], int]:
    supported = {".csv", ".xlsx", ".xls", ".pdf"}
    entries = sorted(os.listdir(source_dir))
    selected: list[str] = []
    ignored: list[str] = []
    for name in entries:
        full = os.path.join(source_dir, name)
        if not os.path.isfile(full):
            continue
        if Path(name).suffix.lower() in supported:
            selected.append(full)
        else:
            ignored.append(full)
    return selected, ignored, len(entries)


def write_in_place(
    workbook_path: str,
    sheet_xml_path: str,
    rows_data: dict[int, dict[str, str]],
) -> str:
    ts = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_path = f"{workbook_path}.backup_{ts}"
    temp_path = f"{workbook_path}.tmp_{ts}"
    _ = shutil.copy2(workbook_path, backup_path)

    with zipfile.ZipFile(workbook_path, "r") as zin:
        with zipfile.ZipFile(temp_path, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)
                if info.filename == sheet_xml_path:
                    data = create_worksheet_xml(rows_data)
                zout.writestr(info, data)

    with zipfile.ZipFile(temp_path, "r") as zf:
        _ = ET.fromstring(zf.read("xl/workbook.xml"))
        _ = ET.fromstring(zf.read(sheet_xml_path))

    os.replace(temp_path, workbook_path)
    return backup_path


def run_selftest_roundtrip_inline_str() -> int:
    rows_data = {
        1: {"A": "date", "D": "amount", "F": "pay", "G": "extra"},
        2: {"A": "2026-02-24", "D": "-12.30", "F": "wx", "G": "hello_inline"},
    }
    xml = create_worksheet_xml(rows_data)
    parsed = parse_sheet_xml(xml, [])
    if parsed.get(2, {}).get("G") != "hello_inline":
        print("SELFTEST_FAIL:inlineStr not preserved")
        return 1
    print("SELFTEST_OK")
    return 0


def run_selftest_dynamic_cols() -> int:
    row = {"A": "2026-02-24", "D": "1.00", "F": "wx"}
    for i in range(30):
        row[col_index_to_letter(7 + i)] = f"k{i}=v{i}"
    rows_data = {1: {"A": "h"}, 2: row}
    xml = create_worksheet_xml(rows_data)
    text = xml.decode("utf-8", errors="replace")
    m = re.search(r"<dimension ref=\"A1:([A-Z]+)\d+\"", text)
    if not m:
        print("SELFTEST_FAIL:missing_dimension")
        return 1
    max_col = m.group(1)
    if col_letter_to_index(max_col) < col_letter_to_index("AA"):
        print(f"SELFTEST_FAIL:max_col_too_small:{max_col}")
        return 1
    if "AA2" not in text:
        print("SELFTEST_FAIL:missing_AA2")
        return 1
    print(f"SELFTEST_OK max_col={max_col}")
    return 0


def setup_cli() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Insert transactions grouped by date")
    _ = parser.add_argument(
        "--source-dir", default="/mnt/r/money_record/auto_insertion_source_files"
    )
    _ = parser.add_argument("--workbook", default="/mnt/r/money_record/money_2026.xlsx")
    _ = parser.add_argument("--sheet", default="2026每日饮食表")
    _ = parser.add_argument("--in-place", action="store_true", default=True)
    _ = parser.add_argument("--dry-run", action="store_true", default=False)
    _ = parser.add_argument("--deps-check", action="store_true", default=False)
    _ = parser.add_argument(
        "--selftest-roundtrip-inline-str", action="store_true", default=False
    )
    _ = parser.add_argument(
        "--selftest-dynamic-cols", action="store_true", default=False
    )
    return parser.parse_args()


def main() -> int:
    args = cast(CliArgs, cast(object, setup_cli()))
    source_dir_arg = args.source_dir
    workbook_arg = args.workbook
    sheet_arg = args.sheet
    in_place = args.in_place
    dry_run = args.dry_run
    deps_check = args.deps_check
    selftest_inline = args.selftest_roundtrip_inline_str
    selftest_dynamic = args.selftest_dynamic_cols

    if selftest_inline:
        return run_selftest_roundtrip_inline_str()
    if selftest_dynamic:
        return run_selftest_dynamic_cols()

    has_pdf = dep_available("pdfplumber")
    has_xlrd = dep_available("xlrd")

    if deps_check:
        print(f"DEP pdfplumber: {'YES' if has_pdf else 'NO'}")
        print(f"DEP xlrd: {'YES' if has_xlrd else 'NO'}")
        return 0

    source_dir = os.path.abspath(source_dir_arg)
    workbook = os.path.abspath(workbook_arg)

    try:
        sheet_xml_path, _ = resolve_sheet_path(workbook, sheet_arg)
    except Exception as exc:  # noqa: BLE001
        print(str(exc), file=sys.stderr)
        return 1

    print(f"RESOLVED_SOURCE_DIR:{source_dir}")
    print(f"RESOLVED_WORKBOOK:{workbook}")
    print(f"RESOLVED_SHEET:{sheet_arg}")
    print(f"RESOLVED_SHEET_XML:{sheet_xml_path}")

    try:
        selected, ignored, seen = discover_source_files(source_dir)
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR_DISCOVERY:{exc}", file=sys.stderr)
        return 1

    print(f"DISCOVERED_FILES:{seen}")
    print(f"SELECTED_FILES:{len(selected)}")
    for ignored_file in ignored:
        print(f"IGNORED_FILE:{ignored_file} reason=unsupported_extension")

    try:
        rows_data = read_sheet_structure(workbook, sheet_xml_path)
    except Exception as exc:  # noqa: BLE001
        print(f"ERROR_READ_SHEET:{exc}", file=sys.stderr)
        return 1

    totals = {
        "files": 0,
        "parsed": 0,
        "inserted": 0,
        "dup_skipped": 0,
        "skipped": 0,
        "warnings": 0,
    }
    all_warnings: list[str] = []

    for file_path in selected:
        parser_id, txs, warnings, stats = parser_dispatch(file_path, has_pdf, has_xlrd)
        parsed = int(stats.get("rows_parsed", 0))
        skipped = int(stats.get("rows_skipped", 0))
        inserted, dup_skipped = insert_transactions(rows_data, txs)
        warn_count = len(warnings)

        totals["files"] += 1
        totals["parsed"] += parsed
        totals["inserted"] += inserted
        totals["dup_skipped"] += dup_skipped
        totals["skipped"] += skipped
        totals["warnings"] += warn_count
        all_warnings.extend(warnings)

        result_line = (
            f"FILE_RESULT:{os.path.basename(file_path)} "
            f"type={Path(file_path).suffix.lower().lstrip('.')} parser={parser_id} "
            f"parsed={parsed} inserted={inserted} dup_skipped={dup_skipped} "
            f"skipped={skipped} warnings={warn_count}"
        )
        print(result_line)

    total_line = (
        f"TOTAL: files={totals['files']} parsed={totals['parsed']} "
        f"inserted={totals['inserted']} dup_skipped={totals['dup_skipped']} "
        f"skipped={totals['skipped']} warnings={totals['warnings']}"
    )
    print(total_line)

    unique_sorted_warnings = sorted(set(all_warnings))
    for warning in unique_sorted_warnings[:50]:
        print(warning)
    if len(unique_sorted_warnings) > 50:
        print("WARNINGS_TRUNCATED:true")

    if dry_run:
        print("DRY_RUN:true")
        return 0

    if in_place:
        try:
            backup_path = write_in_place(workbook, sheet_xml_path, rows_data)
        except Exception as exc:  # noqa: BLE001
            print(f"ERROR_WRITE:{exc}", file=sys.stderr)
            return 1
        print(f"BACKUP_CREATED:{backup_path}")
        return 0

    print("DRY_RUN:true")
    return 0


if __name__ == "__main__":
    sys.exit(main())
