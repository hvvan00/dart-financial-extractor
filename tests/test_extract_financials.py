import tempfile
import unittest
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

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
    run,
    simplify_statement_dataframe,
)


class FakeTable:
    def __init__(self, value, separate):
        self.value = value
        self.separate = separate

    def to_DataFrame(self, **kwargs):
        expected_options = {
            "lang": "ko",
            "label": "Separate" if self.separate else "Consolidated",
            "show_abstract": True,
            "show_class": False,
            "show_concept": False,
            "separator": False,
        }
        if kwargs != expected_options:
            raise AssertionError(f"unexpected XBRL table options: {kwargs}")
        if kwargs.get("separator") is not False:
            raise AssertionError("numeric extraction must disable string separators")
        definition = "연결 재무제표 (Unit: KRW)"
        return pd.DataFrame(
            [[self.value, 1_234_567.0]],
            columns=pd.MultiIndex.from_tuples(
                [
                    (definition, "label_ko"),
                    (definition, "[2025-12-31]연결재무제표"),
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
        return [FakeTable(f"{name}-{'별도' if separate else '연결'}", separate)]

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

    def test_open_dart_no_data_message_is_treated_as_missing_xbrl(self):
        finance = SimpleNamespace(
            download_xbrl=lambda **kwargs: (_ for _ in ()).throw(
                RuntimeError("013 조회된 데이터가 없습니다.")
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
        self.assertEqual(
            list(statements["재무상태표"].columns),
            ["항목", "2025-12-31"],
        )

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
        definition = "연결 재무상태표 (Unit: KRW)"
        columns = pd.MultiIndex.from_tuples(
            [
                (definition, "concept_id"),
                (definition, "label_ko"),
                (definition, "label_en"),
                (definition, "class0"),
                (definition, "주석"),
                (definition, "[2025-12-31]연결재무제표"),
                (definition, "[2024-12-31]연결재무제표"),
            ]
        )
        self.statements = {
            name: pd.DataFrame(
                [
                    [
                        "dart_CashAndCashEquivalents",
                        "현금및현금성자산",
                        "Cash and cash equivalents",
                        "유동자산",
                        5,
                        1_234_567.0,
                        900_000.0,
                    ]
                ],
                columns=columns,
            )
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

    def test_simplifies_to_korean_item_and_period_amounts_only(self):
        definition = "연결 포괄손익계산서 (Unit: KRW)"
        frame = pd.DataFrame(
            [
                ["dart_AssetsAbstract", "자산 [abstract]", "Assets", "", None, None],
                [
                    "ifrs-full_CashAndCashEquivalents",
                    "현금및현금성자산",
                    "Cash and cash equivalents",
                    "5",
                    1_234_567.0,
                    900_000.0,
                ],
            ],
            columns=pd.MultiIndex.from_tuples(
                [
                    (definition, "concept_id"),
                    (definition, "label_ko"),
                    (definition, "label_en"),
                    (definition, "주석"),
                    (definition, "[2025-01-01,2025-12-31]연결재무제표"),
                    (definition, "[2024-01-01,2024-12-31]연결재무제표"),
                ]
            ),
        )

        simplified = simplify_statement_dataframe(frame)

        self.assertEqual(
            list(simplified.columns),
            [
                "항목",
                "2025-01-01 ~ 2025-12-31",
                "2024-01-01 ~ 2024-12-31",
            ],
        )
        self.assertEqual(simplified.iloc[0, 0], "자산")
        self.assertEqual(simplified.iloc[1, 0], "현금및현금성자산")
        self.assertEqual(simplified.iloc[1, 1], 1_234_567)
        self.assertNotIn("Cash and cash equivalents", simplified.to_string())

    def test_single_mode_has_exactly_three_formatted_sheets_and_numeric_cells(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            paths = export_statements(
                self.statements,
                receipt_number="20240319000709",
                output_mode="single",
                output_dir=Path(temp_dir),
            )

            self.assertEqual(len(paths), 1)
            workbook = load_workbook(paths[0])
            self.assertEqual(workbook.sheetnames, list(STATEMENT_NAMES))
            for worksheet in workbook.worksheets:
                self.assertEqual(worksheet.freeze_panes, "A2")
                self.assertFalse(worksheet.sheet_view.showGridLines)
                self.assertTrue(worksheet["A1"].font.bold)
                self.assertEqual(
                    [cell.value for cell in worksheet[1]],
                    ["항목", "2025-12-31", "2024-12-31"],
                )
                self.assertEqual(worksheet["B2"].data_type, "n")
                self.assertEqual(worksheet["B2"].value, 1_234_567)
                self.assertIn(",", worksheet["B2"].number_format)
                self.assertNotIn("concept", " ".join(str(cell.value) for cell in worksheet[1]).lower())

    def test_separate_mode_creates_exactly_three_individual_files(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            paths = export_statements(
                self.statements,
                receipt_number="20240319000709",
                output_mode="separate",
                output_dir=Path(temp_dir),
            )

            self.assertEqual(len(paths), 3)
            self.assertEqual(len(list(Path(temp_dir).glob("*.xlsx"))), 3)
            for path, statement_name in zip(paths, STATEMENT_NAMES):
                workbook = load_workbook(path)
                self.assertEqual(workbook.sheetnames, [statement_name])


class PipelineFallbackTests(unittest.TestCase):
    def test_run_uses_pdf_only_when_xbrl_is_missing(self):
        finance = SimpleNamespace(
            download_xbrl=lambda **kwargs: (_ for _ in ()).throw(
                FileNotFoundError("XBRL File Not Found")
            )
        )
        dart = SimpleNamespace(
            api=SimpleNamespace(finance=finance),
            set_api_key=lambda **kwargs: None,
        )
        statements = {
            name: pd.DataFrame([["현금", 1_000]], columns=["과목", "당기"])
            for name in STATEMENT_NAMES
        }

        with tempfile.TemporaryDirectory() as temp_dir:
            with (
                patch(
                    "extract_financials.resolve_document_number",
                    return_value="11134296",
                ),
                patch(
                    "extract_financials.download_filing_pdf",
                    return_value=Path(temp_dir) / "filing.pdf",
                ),
                patch(
                    "extract_financials.extract_pdf_statements",
                    return_value=(statements, "consolidated"),
                ),
            ):
                paths, selected_scope, receipt_number, source = run(
                    "20260317801285",
                    output_dir=Path(temp_dir) / "output",
                    environ={"DART_API_KEY": "secret"},
                    dart_module=dart,
                )

        self.assertEqual(len(paths), 1)
        self.assertEqual(selected_scope, "consolidated")
        self.assertEqual(receipt_number, "20260317801285")
        self.assertEqual(source, "PDF")


if __name__ == "__main__":
    unittest.main()
