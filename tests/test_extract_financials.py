import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace

import pandas as pd
from openpyxl import load_workbook

from extract_financials import (
    MissingApiKeyError,
    MissingStatementError,
    MissingXbrlError,
    STATEMENT_NAMES,
    export_statements,
    extract_statements,
    flatten_dataframe_columns,
    load_xbrl,
    require_api_key,
)


class FakeTable:
    def __init__(self, value):
        self.value = value

    def to_DataFrame(self, **kwargs):
        if kwargs.get("separator") is not False:
            raise AssertionError("numeric extraction must disable string separators")
        return pd.DataFrame(
            [[self.value, 1_234_567.0]],
            columns=pd.MultiIndex.from_tuples(
                [
                    ("계정", "과목"),
                    ("2025-12-31", "연결재무제표"),
                ]
            ),
        )


class FakeXbrl:
    def __init__(self, consolidated=True, missing_consolidated=()):
        self.consolidated = consolidated
        self.missing_consolidated = set(missing_consolidated)
        self.calls = []

    def exist_consolidated(self):
        return self.consolidated

    def _result(self, name, separate):
        self.calls.append((name, separate))
        if not separate and name in self.missing_consolidated:
            return None
        return [FakeTable(f"{name}-{'별도' if separate else '연결'}")]

    def get_financial_statement(self, separate=False):
        return self._result("재무상태표", separate)

    def get_income_statement(self, separate=False):
        return self._result("손익계산서", separate)

    def get_cash_flows(self, separate=False):
        return self._result("현금흐름표", separate)


class ApiKeyTests(unittest.TestCase):
    def test_missing_api_key_has_clear_error(self):
        with self.assertRaisesRegex(MissingApiKeyError, "DART_API_KEY"):
            require_api_key({})

    def test_blank_api_key_has_clear_error(self):
        with self.assertRaisesRegex(MissingApiKeyError, "DART_API_KEY"):
            require_api_key({"DART_API_KEY": "  "})

    def test_reads_api_key_from_named_environment_variable(self):
        self.assertEqual(
            require_api_key({"DART_API_KEY": " secret-value "}),
            "secret-value",
        )


class XbrlLoadingTests(unittest.TestCase):
    def test_missing_xbrl_has_clear_error(self):
        finance = SimpleNamespace(
            download_xbrl=lambda **kwargs: (_ for _ in ()).throw(
                FileNotFoundError("XBRL File Not Found")
            )
        )
        dart = SimpleNamespace(api=SimpleNamespace(finance=finance))

        with tempfile.TemporaryDirectory() as temp_dir:
            with self.assertRaisesRegex(MissingXbrlError, "XBRL"):
                load_xbrl("20240319000709", Path(temp_dir), dart)

    def test_empty_xbrl_has_clear_error(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            xbrl_file = Path(temp_dir) / "filing.xbrl"
            xbrl_file.touch()
            finance = SimpleNamespace(download_xbrl=lambda **kwargs: str(xbrl_file))
            dart = SimpleNamespace(
                api=SimpleNamespace(finance=finance),
                xbrl=SimpleNamespace(
                    get_xbrl_from_file=lambda path: SimpleNamespace(is_empty=True)
                ),
            )

            with self.assertRaisesRegex(MissingXbrlError, "XBRL"):
                load_xbrl("20240319000709", Path(temp_dir), dart)


class StatementExtractionTests(unittest.TestCase):
    def test_auto_prefers_complete_consolidated_statements(self):
        xbrl = FakeXbrl(consolidated=True)

        statements, selected_scope = extract_statements(xbrl, "auto")

        self.assertEqual(selected_scope, "consolidated")
        self.assertEqual(tuple(statements), STATEMENT_NAMES)
        self.assertTrue(all(separate is False for _, separate in xbrl.calls))

    def test_auto_falls_back_to_complete_separate_statements(self):
        xbrl = FakeXbrl(
            consolidated=True,
            missing_consolidated={"현금흐름표"},
        )

        statements, selected_scope = extract_statements(xbrl, "auto")

        self.assertEqual(selected_scope, "separate")
        self.assertEqual(tuple(statements), STATEMENT_NAMES)
        self.assertIn(("현금흐름표", False), xbrl.calls)
        self.assertIn(("재무상태표", True), xbrl.calls)

    def test_auto_uses_separate_when_consolidated_does_not_exist(self):
        xbrl = FakeXbrl(consolidated=False)

        _, selected_scope = extract_statements(xbrl, "auto")

        self.assertEqual(selected_scope, "separate")
        self.assertTrue(all(separate is True for _, separate in xbrl.calls))

    def test_missing_statement_has_clear_error(self):
        xbrl = FakeXbrl(
            consolidated=True,
            missing_consolidated={"현금흐름표"},
        )
        xbrl.get_cash_flows = lambda separate=False: None

        with self.assertRaisesRegex(
            MissingStatementError,
            "현금흐름표",
        ):
            extract_statements(xbrl, "auto")


class ExcelOutputTests(unittest.TestCase):
    def setUp(self):
        columns = pd.MultiIndex.from_tuples(
            [
                ("계정", "과목"),
                ("금액", "2025-12-31"),
            ]
        )
        self.statements = {
            name: pd.DataFrame([["현금", 1_234_567.0]], columns=columns)
            for name in STATEMENT_NAMES
        }

    def test_flattens_multi_level_headers_and_deduplicates(self):
        frame = pd.DataFrame(
            [[1, 2]],
            columns=pd.MultiIndex.from_tuples([("금액", "당기"), ("금액", "당기")]),
        )

        flattened = flatten_dataframe_columns(frame)

        self.assertEqual(
            flattened.columns.to_list(),
            ["금액 | 당기", "금액 | 당기 (2)"],
        )

    def test_single_mode_has_exactly_three_formatted_sheets_and_numeric_cells(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            flattened = {
                name: flatten_dataframe_columns(frame)
                for name, frame in self.statements.items()
            }
            paths = export_statements(
                flattened,
                receipt_number="20240319000709",
                output_mode="single",
                output_dir=Path(temp_dir),
            )

            self.assertEqual(len(paths), 1)
            workbook = load_workbook(paths[0])
            self.assertEqual(workbook.sheetnames, list(STATEMENT_NAMES))
            for worksheet in workbook.worksheets:
                self.assertEqual(worksheet.freeze_panes, "A2")
                self.assertTrue(worksheet["A1"].font.bold)
                self.assertEqual(worksheet["B2"].data_type, "n")
                self.assertEqual(worksheet["B2"].value, 1_234_567)
                self.assertIn(",", worksheet["B2"].number_format)

    def test_separate_mode_creates_exactly_three_individual_files(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            flattened = {
                name: flatten_dataframe_columns(frame)
                for name, frame in self.statements.items()
            }
            paths = export_statements(
                flattened,
                receipt_number="20240319000709",
                output_mode="separate",
                output_dir=Path(temp_dir),
            )

            self.assertEqual(len(paths), 3)
            self.assertEqual(len(list(Path(temp_dir).glob("*.xlsx"))), 3)
            for path, statement_name in zip(paths, STATEMENT_NAMES):
                workbook = load_workbook(path)
                self.assertEqual(workbook.sheetnames, [statement_name])


if __name__ == "__main__":
    unittest.main()
