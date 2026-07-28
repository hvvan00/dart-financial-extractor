"""DART disclosure URL and receipt-number parsing."""

from __future__ import annotations

import re
from urllib.parse import parse_qs, urlparse


RECEIPT_NUMBER_PATTERN = re.compile(r"^\d{14}$")


class InvalidDisclosureReference(ValueError):
    """Raised when a DART URL or receipt number cannot be parsed."""


def parse_receipt_number(value: str) -> str:
    """Return a 14-digit DART receipt number from a number or disclosure URL.

    Only official DART hosts are accepted for URLs.  A URL must contain the
    receipt number in its ``rcpNo`` query parameter.
    """

    candidate = (value or "").strip()
    if RECEIPT_NUMBER_PATTERN.fullmatch(candidate):
        return candidate

    try:
        parsed = urlparse(candidate)
    except ValueError as exc:
        raise InvalidDisclosureReference(
            "올바른 DART 공시 URL 또는 14자리 접수번호를 입력하세요."
        ) from exc

    if parsed.scheme.lower() not in {"http", "https"} or not parsed.netloc:
        raise InvalidDisclosureReference(
            "입력값은 14자리 접수번호 또는 http(s) DART 공시 URL이어야 합니다."
        )

    try:
        hostname = (parsed.hostname or "").lower().rstrip(".")
    except ValueError as exc:
        raise InvalidDisclosureReference("DART 공시 URL의 호스트 형식이 올바르지 않습니다.") from exc

    if hostname != "dart.fss.or.kr" and not hostname.endswith(".dart.fss.or.kr"):
        raise InvalidDisclosureReference(
            "dart.fss.or.kr 도메인의 공식 DART 공시 URL만 사용할 수 있습니다."
        )

    query = parse_qs(parsed.query, keep_blank_values=True)
    receipt_values = [
        receipt
        for key, values in query.items()
        if key.lower() == "rcpno"
        for receipt in values
    ]

    if len(receipt_values) != 1 or not RECEIPT_NUMBER_PATTERN.fullmatch(
        receipt_values[0]
    ):
        raise InvalidDisclosureReference(
            "DART 공시 URL에서 유효한 14자리 rcpNo 접수번호를 찾을 수 없습니다."
        )

    return receipt_values[0]


__all__ = [
    "InvalidDisclosureReference",
    "parse_receipt_number",
]
