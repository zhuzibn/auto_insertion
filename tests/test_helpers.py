import importlib
import io
import os
import types
import zipfile
from dataclasses import dataclass, replace
from pathlib import Path
import sys
import tempfile
from typing import Any, Protocol, cast
import unittest
from unittest.mock import patch
from contextlib import redirect_stdout

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))


class ScriptModule(Protocol):
    Tx: Any
    argparse: Any
    NS_MAIN: str

    def create_worksheet_xml(self, rows_data: dict[int, dict[str, str]]) -> bytes: ...

    def classify_csv(self, file_path: str) -> str: ...

    def col_index_to_letter(self, index: int) -> str: ...

    def col_letter_to_index(self, letter: str) -> int: ...

    def normalize_date(self, value: str) -> str | None: ...

    def tx_fingerprint(self, tx: object) -> str: ...

    def build_date_blocks(
        self, rows_data: dict[int, dict[str, str]]
    ) -> dict[str, dict[str, int]]: ...

    def parse_jd_csv(
        self, file_path: str
    ) -> tuple[list[object], list[str], dict[str, int]]: ...

    def parse_pdf_transactions(
        self, file_path: str
    ) -> tuple[list[object], list[str], dict[str, int]]: ...

    def parse_wechat_xlsx(
        self, file_path: str
    ) -> tuple[list[object], list[str], dict[str, int]]: ...

    def insert_transactions(
        self,
        rows_data: dict[int, dict[str, str]],
        txs: list[object],
        repair_jd_legacy_sign: bool = False,
    ) -> tuple[int, int, int]: ...

    def main(self) -> int: ...


script = cast(
    ScriptModule, cast(object, importlib.import_module("insert_transactions_by_date"))
)
classify_csv = script.classify_csv
col_index_to_letter = script.col_index_to_letter
col_letter_to_index = script.col_letter_to_index
normalize_date = script.normalize_date
tx_fingerprint = script.tx_fingerprint
parse_wechat_xlsx = script.parse_wechat_xlsx


@dataclass
class DummyTx:
    date: str
    amount: float
    payment_method: str
    platform: str
    merchant: str
    order_id: str
    source_file: str
    source_path: str
    source_row: str


class ColumnLetterIndexTests(unittest.TestCase):
    def test_col_index_to_letter_boundaries(self):
        self.assertEqual(col_index_to_letter(26), "Z")
        self.assertEqual(col_index_to_letter(27), "AA")
        self.assertEqual(col_index_to_letter(702), "ZZ")
        self.assertEqual(col_index_to_letter(703), "AAA")

    def test_col_letter_to_index_boundaries(self):
        self.assertEqual(col_letter_to_index("Z"), 26)
        self.assertEqual(col_letter_to_index("AA"), 27)
        self.assertEqual(col_letter_to_index("ZZ"), 702)
        self.assertEqual(col_letter_to_index("AAA"), 703)


class NormalizeDateTests(unittest.TestCase):
    def test_normalize_date_yyyy_mm_dd(self):
        self.assertEqual(normalize_date("2026-02-07"), "2026-02-07")

    def test_normalize_date_with_timestamp(self):
        self.assertEqual(normalize_date("2026-02-07 19:26:08"), "2026-02-07")

    def test_normalize_date_compact_yyyymmdd(self):
        self.assertEqual(normalize_date("20260207"), "2026-02-07")


class FingerprintTests(unittest.TestCase):
    def _build_tx(self, source_row: str):
        return DummyTx(
            date="2026-02-07",
            amount=-12.34,
            payment_method="微信",
            platform="wechat",
            merchant="Test Merchant",
            order_id="OID-001",
            source_file="sample.csv",
            source_path="/tmp/sample.csv",
            source_row=source_row,
        )

    def test_tx_fingerprint_is_stable_for_same_input(self):
        tx = self._build_tx("10")
        fp1 = tx_fingerprint(tx)
        fp2 = tx_fingerprint(tx)
        self.assertEqual(fp1, fp2)

    def test_tx_fingerprint_changes_when_source_row_changes(self):
        fp1 = tx_fingerprint(self._build_tx("10"))
        fp2 = tx_fingerprint(self._build_tx("11"))
        self.assertNotEqual(fp1, fp2)


class DateBlockDetectionTests(unittest.TestCase):
    def test_build_date_blocks_keeps_single_block_for_consecutive_same_date_rows(self):
        rows_data = {
            1: {"A": "日期", "D": "金额", "F": "支付方式"},
            2: {"A": "2026-02-07", "D": "-10.00", "F": "微信"},
            3: {"A": "2026-02-07", "D": "-20.00", "F": "微信"},
            4: {"A": "2026-02-07", "D": "-30.00", "F": "微信"},
            5: {"A": "2026-02-08", "D": "-40.00", "F": "微信"},
        }

        blocks = script.build_date_blocks(rows_data)

        self.assertEqual(blocks["2026-02-07"], {"start": 2, "end": 4})


class CsvDispatchSignatureTests(unittest.TestCase):
    def _write_temp_csv(self, content: str) -> str:
        fd, path = tempfile.mkstemp(suffix=".csv", text=True)
        os.close(fd)
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            _ = f.write(content)
        self.addCleanup(lambda: os.path.exists(path) and os.remove(path))
        return path

    def test_classify_csv_identifies_jd_by_header_signature(self):
        path = self._write_temp_csv("京东交易流水,京东账号名,导出交易类型\n")
        self.assertEqual(classify_csv(path), "jd_csv")

    def test_classify_csv_identifies_alipay_by_header_signature(self):
        path = self._write_temp_csv("交易创建时间,交易订单号,收/支,交易对方\n")
        self.assertEqual(classify_csv(path), "alipay_csv")


class MainOutputAccountingTests(unittest.TestCase):
    def _tx(self, source_file: str, source_row: str, amount: float, merchant: str):
        return script.Tx(
            date="2026-02-07",
            amount=amount,
            payment_method="微信",
            platform="wechat",
            merchant=merchant,
            order_id="",
            source_file=source_file,
            source_path=f"/tmp/{source_file}",
            source_row=source_row,
            extra_fields=[],
        )

    def test_main_reports_per_file_inserted_and_dup_counts(self):
        file_a = "/tmp/a.csv"
        file_b = "/tmp/b.csv"
        tx_a_1 = self._tx("a.csv", "1", -10.0, "dup-merchant")
        tx_a_2 = self._tx("a.csv", "1", -10.0, "dup-merchant")
        tx_b_1 = self._tx("b.csv", "1", -20.0, "unique-merchant")

        def fake_dispatch(file_path: str, _has_pdf: bool, _has_xlrd: bool):
            if file_path == file_a:
                return (
                    "jd_csv",
                    [tx_a_1, tx_a_2],
                    [],
                    {"rows_parsed": 2, "rows_skipped": 0},
                )
            if file_path == file_b:
                return "jd_csv", [tx_b_1], [], {"rows_parsed": 1, "rows_skipped": 0}
            return "unknown_csv", [], [], {"rows_parsed": 0, "rows_skipped": 0}

        args = script.argparse.Namespace(
            source_dir="/tmp/src",
            workbook="/tmp/book.xlsx",
            sheet="Sheet1",
            in_place=True,
            dry_run=True,
            deps_check=False,
            selftest_roundtrip_inline_str=False,
            selftest_dynamic_cols=False,
        )

        out = io.StringIO()
        with (
            patch.object(script, "setup_cli", return_value=args),
            patch.object(script, "dep_available", return_value=False),
            patch.object(
                script,
                "resolve_sheet_path",
                return_value=("xl/worksheets/sheet1.xml", []),
            ),
            patch.object(
                script, "discover_source_files", return_value=([file_a, file_b], [], 2)
            ),
            patch.object(script, "parser_dispatch", side_effect=fake_dispatch),
            patch.object(script, "read_sheet_structure", return_value={}),
            redirect_stdout(out),
        ):
            rc = script.main()

        stdout = out.getvalue()
        self.assertEqual(rc, 0)
        self.assertIn(
            "FILE_RESULT:a.csv type=csv parser=jd_csv parsed=2 inserted=1 dup_skipped=1 skipped=0 warnings=0",
            stdout,
        )
        self.assertIn(
            "FILE_RESULT:b.csv type=csv parser=jd_csv parsed=1 inserted=1 dup_skipped=0 skipped=0 warnings=0",
            stdout,
        )
        self.assertIn(
            "TOTAL: files=2 parsed=3 inserted=2 dup_skipped=1 skipped=0 warnings=0",
            stdout,
        )

    def test_main_truncates_warning_output_after_50_distinct(self):
        file_a = "/tmp/a.csv"
        warns = [f"W{i:03d}" for i in range(60)]

        def fake_dispatch(file_path: str, _has_pdf: bool, _has_xlrd: bool):
            if file_path == file_a:
                return "jd_csv", [], warns, {"rows_parsed": 0, "rows_skipped": 0}
            return "unknown_csv", [], [], {"rows_parsed": 0, "rows_skipped": 0}

        args = script.argparse.Namespace(
            source_dir="/tmp/src",
            workbook="/tmp/book.xlsx",
            sheet="Sheet1",
            in_place=True,
            dry_run=True,
            deps_check=False,
            selftest_roundtrip_inline_str=False,
            selftest_dynamic_cols=False,
        )

        out = io.StringIO()
        with (
            patch.object(script, "setup_cli", return_value=args),
            patch.object(script, "dep_available", return_value=False),
            patch.object(
                script,
                "resolve_sheet_path",
                return_value=("xl/worksheets/sheet1.xml", []),
            ),
            patch.object(
                script, "discover_source_files", return_value=([file_a], [], 1)
            ),
            patch.object(script, "parser_dispatch", side_effect=fake_dispatch),
            patch.object(script, "read_sheet_structure", return_value={}),
            redirect_stdout(out),
        ):
            rc = script.main()

        stdout_lines = out.getvalue().splitlines()
        self.assertEqual(rc, 0)
        warning_lines = [
            line
            for line in stdout_lines
            if line.startswith("W") and line != "WARNINGS_TRUNCATED:true"
        ]
        self.assertEqual(len(warning_lines), 50)
        self.assertEqual(warning_lines[0], "W000")
        self.assertEqual(warning_lines[-1], "W049")
        self.assertIn("WARNINGS_TRUNCATED:true", stdout_lines)


class ParserSkipWarningTests(unittest.TestCase):
    def _write_temp_csv(self, content: str) -> str:
        fd, path = tempfile.mkstemp(suffix=".csv", text=True)
        os.close(fd)
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            _ = f.write(content)
        self.addCleanup(lambda: os.path.exists(path) and os.remove(path))
        return path

    def test_parse_jd_csv_warns_for_nonempty_row_missing_amount(self):
        path = self._write_temp_csv(
            "交易时间,金额,收支,支付方式,交易对方,订单号,交易类型\n"
            "2026-02-07,,支出,微信,测试商户,OID-1,消费\n"
        )

        txs, warnings, stats = script.parse_jd_csv(path)

        self.assertEqual(txs, [])
        self.assertEqual(stats["rows_skipped"], 1)
        self.assertTrue(any(w.startswith("SKIP_JD_INVALID_ROW:") for w in warnings))


class PdfParserSkipWarningTests(unittest.TestCase):
    def test_parse_pdf_warns_for_date_line_without_amount(self):
        class FakePage:
            _text: str

            def __init__(self, text: str):
                self._text = text

            def extract_text(self, layout: bool = False) -> str:
                _ = layout
                return self._text

        class FakePdf:
            pages: list[FakePage]

            def __init__(self, pages: list[FakePage]):
                self.pages = pages

            def __enter__(self) -> "FakePdf":
                return self

            def __exit__(self, exc_type: object, exc: object, tb: object) -> bool:
                _ = (exc_type, exc, tb)
                return False

        def fake_open(_path: str) -> FakePdf:
            return FakePdf(
                [
                    FakePage("2026-02-07 SOME TEXT WITHOUT AMOUNT"),
                ]
            )

        fake_module = types.SimpleNamespace(open=fake_open)

        with patch.dict(sys.modules, {"pdfplumber": fake_module}):
            txs, warnings, stats = script.parse_pdf_transactions("/tmp/fake.pdf")

        self.assertEqual(txs, [])
        self.assertEqual(stats["rows_skipped"], 1)
        self.assertTrue(
            any(
                w.startswith("SKIP_PDF_AMBIGUOUS_AMOUNT:/tmp/fake.pdf:")
                and "2026-02-07 SOME TEXT WITHOUT AMOUNT" in w
                for w in warnings
            )
        )


class WechatXlsxKeyErrorTests(unittest.TestCase):
    def test_parse_wechat_xlsx_handles_missing_sheet1_xml_keyerror(self):
        """KeyError on xl/worksheets/sheet1.xml should return SKIP_XLSX_READ_ERROR warning."""
        fd, path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        self.addCleanup(lambda: os.path.exists(path) and os.remove(path))

        class FakeZipFile:
            def __init__(self, *args: Any, **kwargs: Any) -> None:
                pass

            def read(self, name: str) -> bytes:
                if name == "xl/worksheets/sheet1.xml":
                    raise KeyError("xl/worksheets/sheet1.xml")
                # Raise KeyError for shared strings too (simulates missing file)
                if name == "xl/sharedStrings.xml":
                    raise KeyError("xl/sharedStrings.xml")
                return b""

            def __enter__(self) -> "FakeZipFile":
                return self

            def __exit__(self, *args: Any) -> bool:
                return False

        with patch.object(zipfile, "ZipFile", FakeZipFile):
            txs, warnings, stats = parse_wechat_xlsx(path)

        self.assertEqual(txs, [])
        self.assertEqual(stats["rows_parsed"], 0)
        self.assertTrue(any("SKIP_XLSX_READ_ERROR" in w for w in warnings))


class NumericCellEmissionTests(unittest.TestCase):
    def test_create_worksheet_xml_emits_numeric_cells_without_inlineStr(self):
        """Column D (amount) should emit numeric cells without t=\"inlineStr\" - use <v> directly."""
        import xml.etree.ElementTree as ET

        rows_data = {
            1: {"A": "日期", "D": "金额", "F": "支付方式", "G": "指纹"},
            2: {"A": "2026-02-07", "D": "-12.30", "F": "微信", "G": "hello_inline"},
        }

        xml_bytes = script.create_worksheet_xml(rows_data)
        root = ET.fromstring(xml_bytes)

        ns = {"m": script.NS_MAIN}

        # D2 should be numeric without t="inlineStr"
        d2 = root.find(".//m:c[@r='D2']", ns)
        self.assertIsNotNone(d2, "Cell D2 should exist")
        if d2 is None:
            return
        # Numeric cells should NOT have t="inlineStr"
        self.assertIsNone(d2.get("t"), 'D2 should NOT have t="inlineStr" for numeric')
        # Should have <v> element with the numeric value
        v_elem = d2.find("m:v", ns)
        self.assertIsNotNone(v_elem, "D2 should have <v> element for numeric value")
        if v_elem is None:
            return
        self.assertEqual(v_elem.text, "-12.30", "D2 value should be -12.30")

        # G2 should remain inlineStr (text cell)
        g2 = root.find(".//m:c[@r='G2']", ns)
        self.assertIsNotNone(g2, "Cell G2 should exist")
        if g2 is None:
            return
        self.assertEqual(g2.get("t"), "inlineStr", 'G2 should have t="inlineStr"')
        is_elem = g2.find("m:is/m:t", ns)
        self.assertIsNotNone(is_elem, "G2 should have <is>/<t> for inline string")
        if is_elem is None:
            return
        self.assertEqual(is_elem.text, "hello_inline", "G2 text should be hello_inline")


class JdCsvSignNormalizationTests(unittest.TestCase):
    def _write_temp_csv(self, content: str) -> str:
        fd, path = tempfile.mkstemp(suffix=".csv", text=True)
        os.close(fd)
        with open(path, "w", encoding="utf-8-sig", newline="") as f:
            _ = f.write(content)
        self.addCleanup(lambda: os.path.exists(path) and os.remove(path))
        return path

    def test_parse_jd_csv_with_slash_sign_column_normalizes_expense_to_negative(self):
        # Test the bug case: real-world JD CSV has header '收/支' (not '收支')
        # and '支出' value in that column should make amount negative
        content = (
            "交易时间,金额,收/支,支付方式,交易对方,订单号,交易类型\n"
            + "2026-02-08,102.20,支出,微信,京东商城,JD123,消费\n"
        )
        path = self._write_temp_csv(content)

        txs, warnings, stats = script.parse_jd_csv(path)

        self.assertEqual(len(txs), 1)
        tx = cast(Any, txs[0])
        # Expecting this test to fail first - the bug is that amount remains positive
        self.assertEqual(
            tx.amount, -102.20, "Amount should be negative for '支出' in '收/支' column"
        )
        self.assertEqual(warnings, [])
        self.assertEqual(list(stats.values()), [1, 0])  # rows_parsed=1, rows_skipped=0


class RepairLegacyJdSignTests(unittest.TestCase):
    def test_insert_transactions_opt_in_repair_updates_legacy_positive_row_in_place(
        self,
    ):
        legacy_tx = script.Tx(
            date="2026-02-08",
            amount=102.20,
            payment_method="微信",
            platform="jd",
            merchant="京东商城",
            order_id="JD123",
            source_file="jd.csv",
            source_path="/tmp/jd.csv",
            source_row="2",
            extra_fields=[],
        )
        legacy_fp = script.tx_fingerprint(legacy_tx)
        canonical_tx = replace(legacy_tx, amount=-102.20)
        canonical_fp = script.tx_fingerprint(canonical_tx)

        rows_data = {
            1: {"A": "日期", "D": "金额", "F": "支付方式", "G": "tx_fingerprint"},
            2: {
                "A": "2026-02-08",
                "D": "102.20",
                "F": "微信",
                "G": legacy_fp,
            },
        }

        inserted, dup_skipped, repaired = script.insert_transactions(
            rows_data,
            [canonical_tx],
            repair_jd_legacy_sign=True,
        )

        self.assertEqual(inserted, 0)
        self.assertEqual(dup_skipped, 0)
        self.assertEqual(repaired, 1)
        self.assertEqual(rows_data[2]["D"], "-102.20")
        self.assertEqual(rows_data[2]["G"], canonical_fp)

        inserted2, dup2, repaired2 = script.insert_transactions(
            rows_data,
            [canonical_tx],
            repair_jd_legacy_sign=True,
        )
        self.assertEqual(inserted2, 0)
        self.assertEqual(dup2, 1)
        self.assertEqual(repaired2, 0)
        self.assertEqual(rows_data[2]["D"], "-102.20")
        self.assertEqual(rows_data[2]["G"], canonical_fp)
