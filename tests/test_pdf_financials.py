import tempfile
import unittest
from numbers import Integral
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from pdf_financials import (
    PDF_DOWNLOAD_MAIN_URL,
    PdfExtractionError,
    _merge_wrapped_matrix_rows,
    _region_word_matrix,
    dataframe_from_pdf_matrices,
    download_filing_pdf,
    extract_pdf_statements,
    find_pdf_title_events,
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


class WordRegion:
    width = 595

    def __init__(self, words):
        self.words = words

    def extract_words(self, **kwargs):
        return self.words


def pdf_word(text, x0, x1, top):
    return {
        "text": text,
        "x0": float(x0),
        "x1": float(x1),
        "top": float(top),
        "bottom": float(top + 10),
    }


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
    def test_merges_vertically_wrapped_account_row_with_its_amounts(self):
        matrix = [
            ["과목", "당기", "전기"],
            ["5. 단기차입금 (주석", "", ""],
            ["", "5,928,000,000", "6,770,000,000"],
            ["10,16,18,23)", "", ""],
            ["부채총계", "9,697,775,679", "8,019,646,790"],
        ]

        merged = _merge_wrapped_matrix_rows(matrix)

        self.assertEqual(
            merged[1],
            [
                "5. 단기차입금 (주석 10,16,18,23)",
                "5,928,000,000",
                "6,770,000,000",
            ],
        )
        self.assertEqual(len(merged), 3)

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

    def test_merged_amount_subcolumns_stay_out_of_item_labels(self):
        words = [
            pdf_word("과", 52, 61, 10),
            pdf_word("목", 133, 142, 10),
            pdf_word("제", 225, 234, 10),
            pdf_word("20(당)", 237, 263, 10),
            pdf_word("기", 266, 275, 10),
            pdf_word("제", 386, 395, 10),
            pdf_word("19(전)", 398, 424, 10),
            pdf_word("기", 427, 436, 10),
            pdf_word("Ⅰ.", 52, 64, 30),
            pdf_word("유동자산", 67, 103, 30),
            pdf_word("3,934,983,786", 322, 382, 30),
            pdf_word("9,922,086,349", 482, 543, 30),
            pdf_word("1.", 84, 92, 50),
            pdf_word("현금및현금성자산", 95, 167, 50),
            pdf_word("927,476,242", 249, 302, 50),
            pdf_word("7,744,956,586", 402, 463, 50),
        ]

        matrix = _region_word_matrix(WordRegion(words))
        frame = dataframe_from_pdf_matrices([matrix])

        self.assertEqual(frame.iloc[0, 0], "Ⅰ. 유동자산")
        self.assertEqual(frame.iloc[0, 1], 3_934_983_786)
        self.assertEqual(frame.iloc[0, 2], 9_922_086_349)

    def test_headerless_continuation_page_uses_same_two_period_layout(self):
        words = [
            pdf_word("자", 52, 61, 10),
            pdf_word("산", 91, 100, 10),
            pdf_word("총", 110, 119, 10),
            pdf_word("계", 128, 137, 10),
            pdf_word("12,571,333,026", 316, 382, 10),
            pdf_word("16,981,078,507", 477, 543, 10),
            pdf_word("Ⅰ.", 52, 64, 30),
            pdf_word("유동부채", 67, 103, 30),
            pdf_word("9,697,775,679", 322, 382, 30),
            pdf_word("8,019,646,790", 482, 543, 30),
        ]

        matrix = _region_word_matrix(WordRegion(words))
        frame = dataframe_from_pdf_matrices([matrix])

        self.assertEqual(frame.iloc[0, 0], "자 산 총 계")
        self.assertEqual(frame.iloc[0, 1], 12_571_333_026)
        self.assertEqual(frame.iloc[0, 2], 16_981_078_507)


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
            get_calls = []
            download_calls = []

            def fake_download(**kwargs):
                download_calls.append(kwargs)
                output.write_text("<html>error</html>", encoding="utf-8")
                return {"full_path": str(output)}

            response = SimpleNamespace(raise_for_status=lambda: None)
            request = SimpleNamespace(
                get=lambda **kwargs: get_calls.append(kwargs) or response,
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

            self.assertEqual(get_calls[0]["url"], PDF_DOWNLOAD_MAIN_URL)
            self.assertEqual(
                get_calls[0]["payload"],
                {
                    "rcp_no": "20260317801285",
                    "dcm_no": "11134296",
                },
            )
            self.assertIn("/pdf/download/main.do?", download_calls[0]["referer"])


class PdfStatementExtractionTests(unittest.TestCase):
    def test_numbered_financial_statement_notes_are_a_boundary(self):
        events = find_pdf_title_events(
            FakePdf([FakePage("5. 재무제표 주석")])
        )

        self.assertEqual(len(events), 1)
        self.assertIsNone(events[0].statement_name)

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
