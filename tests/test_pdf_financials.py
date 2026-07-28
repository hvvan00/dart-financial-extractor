import tempfile
import unittest
from numbers import Integral
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from pdf_financials import (
    PdfExtractionError,
    dataframe_from_pdf_matrices,
    download_filing_pdf,
    extract_pdf_statements,
    parse_pdf_number,
    resolve_document_number,
)


def statement_matrix(label):
    return [
        ["과 목", "주석", "당기", "전기"],
        [label, "1", "1,234,567", "(200,000)"],
        ["합계", "", "2,000,000", "1,500,000"],
    ]


class FakeRegion:
    def __init__(self, matrix):
        self.matrix = matrix

    def extract_tables(self, **kwargs):
        return [self.matrix] if self.matrix else []


class FakePage:
    width = 595
    height = 842

    def __init__(self, title=None, matrix=None):
        self.title = title
        self.matrix = matrix

    def extract_words(self, **kwargs):
        if self.title is None:
            return []
        return [
            {
                "text": self.title,
                "top": 50.0,
                "bottom": 65.0,
                "x0": 100.0,
            }
        ]

    def crop(self, bbox):
        return FakeRegion(self.matrix)


class FakePdf:
    def __init__(self, pages):
        self.pages = pages

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        return False


class PdfNumberTests(unittest.TestCase):
    def test_converts_korean_financial_amount_strings(self):
        cases = {
            "1,234,567": 1_234_567,
            "(1,234,567)": -1_234_567,
            "△ 1,234": -1_234,
            "1,234.50": 1234.5,
            "-": None,
            "주석 12": None,
        }
        for value, expected in cases.items():
            with self.subTest(value=value):
                self.assertEqual(parse_pdf_number(value), expected)

    def test_dataframe_preserves_pdf_amounts_as_numbers(self):
        frame = dataframe_from_pdf_matrices([statement_matrix("현금")])

        self.assertIsNotNone(frame)
        self.assertEqual(frame.iloc[0, 2], 1_234_567)
        self.assertEqual(frame.iloc[0, 3], -200_000)
        self.assertIsInstance(frame.iloc[0, 2], Integral)

    def test_section_label_is_data_not_an_extra_header(self):
        matrix = [
            ["과목", "주석", "당기", "전기"],
            ["자산", "", "", ""],
            ["현금및현금성자산", "5", "1,234", "900"],
        ]

        frame = dataframe_from_pdf_matrices([matrix])

        self.assertIsNotNone(frame)
        self.assertEqual(list(frame.columns), ["과목", "주석", "당기", "전기"])
        self.assertEqual(frame.iloc[0, 0], "자산")
        self.assertEqual(frame.iloc[1, 2], 1_234)


class DartPdfDownloadTests(unittest.TestCase):
    def test_resolves_document_number_from_dart_viewdoc_html(self):
        response = SimpleNamespace(
            text=(
                "javascript: viewDoc('20260317801285', '11134296', "
                "null, null, null, 'dart3.xsd')"
            ),
            raise_for_status=lambda: None,
        )
        request = SimpleNamespace(get=lambda **kwargs: response)
        dart = SimpleNamespace(utils=SimpleNamespace(request=request))

        self.assertEqual(
            resolve_document_number("20260317801285", dart),
            "11134296",
        )

    def test_missing_document_number_has_clear_error(self):
        response = SimpleNamespace(text="<html>no document id</html>", raise_for_status=lambda: None)
        request = SimpleNamespace(get=lambda **kwargs: response)
        dart = SimpleNamespace(utils=SimpleNamespace(request=request))

        with self.assertRaisesRegex(PdfExtractionError, "dcmNo"):
            resolve_document_number("20260317801285", dart)

    def test_download_rejects_non_pdf_response(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            output = Path(temp_dir) / "response.pdf"

            def fake_download(**kwargs):
                output.write_text("<html>error</html>", encoding="utf-8")
                return {"full_path": str(output)}

            response = SimpleNamespace(raise_for_status=lambda: None)
            request = SimpleNamespace(
                get=lambda **kwargs: response,
                download=fake_download,
            )
            dart = SimpleNamespace(utils=SimpleNamespace(request=request))

            with self.assertRaisesRegex(PdfExtractionError, "PDF 형식"):
                download_filing_pdf(
                    "20260317801285",
                    "11134296",
                    Path(temp_dir),
                    dart,
                )


class PdfStatementExtractionTests(unittest.TestCase):
    def test_auto_extracts_complete_consolidated_set(self):
        pages = [
            FakePage("연 결 재 무 상 태 표", statement_matrix("자산")),
            FakePage("연 결 포 괄 손 익 계 산 서", statement_matrix("매출액")),
            FakePage("연 결 현 금 흐 름 표", statement_matrix("영업활동현금흐름")),
        ]

        with patch("pdfplumber.open", return_value=FakePdf(pages)):
            statements, selected_scope = extract_pdf_statements(
                Path("filing.pdf"),
                "auto",
            )

        self.assertEqual(selected_scope, "consolidated")
        self.assertEqual(
            tuple(statements),
            ("재무상태표", "손익계산서", "현금흐름표"),
        )
        self.assertEqual(statements["재무상태표"].iloc[0, 2], 1_234_567)

    def test_auto_falls_back_to_complete_separate_pdf_set(self):
        pages = [
            FakePage("재 무 상 태 표", statement_matrix("자산")),
            FakePage("포 괄 손 익 계 산 서", statement_matrix("매출액")),
            FakePage("현 금 흐 름 표", statement_matrix("영업활동현금흐름")),
        ]

        with patch("pdfplumber.open", return_value=FakePdf(pages)):
            _, selected_scope = extract_pdf_statements(Path("filing.pdf"), "auto")

        self.assertEqual(selected_scope, "separate")

    def test_missing_pdf_statement_has_clear_error(self):
        pages = [
            FakePage("재 무 상 태 표", statement_matrix("자산")),
            FakePage("포 괄 손 익 계 산 서", statement_matrix("매출액")),
        ]

        with patch("pdfplumber.open", return_value=FakePdf(pages)):
            with self.assertRaisesRegex(PdfExtractionError, "현금흐름표"):
                extract_pdf_statements(Path("filing.pdf"), "auto")

    def test_scanned_pdf_has_clear_error(self):
        with patch("pdfplumber.open", return_value=FakePdf([FakePage()])):
            with self.assertRaisesRegex(PdfExtractionError, "스캔 이미지 PDF"):
                extract_pdf_statements(Path("filing.pdf"), "auto")


if __name__ == "__main__":
    unittest.main()
