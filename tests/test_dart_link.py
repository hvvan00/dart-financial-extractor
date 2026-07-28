import unittest

from dart_link import (
    InvalidDisclosureReference,
    parse_disclosure_reference,
    parse_receipt_number,
)


class ParseReceiptNumberTests(unittest.TestCase):
    def test_accepts_raw_14_digit_receipt_number(self):
        self.assertEqual(
            parse_receipt_number(" 20240319000709 "),
            "20240319000709",
        )

    def test_extracts_receipt_number_from_main_disclosure_url(self):
        self.assertEqual(
            parse_receipt_number(
                "https://dart.fss.or.kr/dsaf001/main.do?rcpNo=20240319000709"
            ),
            "20240319000709",
        )

    def test_accepts_official_dart_subdomain_and_case_insensitive_parameter(self):
        self.assertEqual(
            parse_receipt_number(
                "https://m.dart.fss.or.kr/viewer?foo=1&RCPNO=20240319000709"
            ),
            "20240319000709",
        )

    def test_preserves_optional_document_number_for_pdf_fallback(self):
        reference = parse_disclosure_reference(
            "https://dart.fss.or.kr/dsaf001/main.do"
            "?rcpNo=20260317801285&dcmNo=11134296"
        )

        self.assertEqual(reference.receipt_number, "20260317801285")
        self.assertEqual(reference.document_number, "11134296")

    def test_rejects_non_dart_url(self):
        with self.assertRaisesRegex(InvalidDisclosureReference, "공식 DART"):
            parse_receipt_number(
                "https://example.com/dsaf001/main.do?rcpNo=20240319000709"
            )

    def test_rejects_url_without_receipt_number(self):
        with self.assertRaisesRegex(InvalidDisclosureReference, "rcpNo"):
            parse_receipt_number("https://dart.fss.or.kr/dsaf001/main.do")

    def test_rejects_wrong_length_or_non_numeric_receipt_number(self):
        invalid_values = (
            "2024031900070",
            "202403190007090",
            "2024031900070A",
            "not-a-link",
        )
        for value in invalid_values:
            with self.subTest(value=value):
                with self.assertRaises(InvalidDisclosureReference):
                    parse_receipt_number(value)

    def test_rejects_duplicate_receipt_parameters(self):
        with self.assertRaisesRegex(InvalidDisclosureReference, "rcpNo"):
            parse_receipt_number(
                "https://dart.fss.or.kr/main.do"
                "?rcpNo=20240319000709&rcpNo=20240319000710"
            )

    def test_rejects_invalid_document_number(self):
        with self.assertRaisesRegex(InvalidDisclosureReference, "dcmNo"):
            parse_disclosure_reference(
                "https://dart.fss.or.kr/main.do"
                "?rcpNo=20240319000709&dcmNo=not-numeric"
            )


if __name__ == "__main__":
    unittest.main()
