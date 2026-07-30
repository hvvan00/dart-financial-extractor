"""Extract exactly three clean financial statements from a DART filing."""

from __future__ import annotations

import argparse
import html
import math
import numbers
import os
import re
import sys
import tempfile
import unicodedata
from collections.abc import Mapping, Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

from dart_link import InvalidDisclosureReference, parse_disclosure_reference
from pdf_financials import (
    DART_MAIN_URL,
    PdfExtractionError,
    download_filing_pdf,
    extract_pdf_statements,
    resolve_document_number,
)


STATEMENT_SPECS = (
    ("재무상태표", "get_financial_statement"),
    ("손익계산서", "get_income_statement"),
    ("현금흐름표", "get_cash_flows"),
)
STATEMENT_NAMES = tuple(name for name, _ in STATEMENT_SPECS)
SCOPE_LABELS = {
    "consolidated": "연결",
    "separate": "별도",
}
REPORT_TYPES = ("사업보고서", "반기보고서", "분기보고서", "감사보고서")
_INVALID_FILENAME_CHARACTERS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_HTML_TAG_PATTERN = re.compile(r"<[^>]+>")


class FinancialExtractionError(RuntimeError):
    """Base error for expected extraction failures."""


class MissingApiKeyError(FinancialExtractionError):
    """Raised when DART_API_KEY is absent."""


class MissingXbrlError(FinancialExtractionError):
    """Raised when the filing does not provide a usable XBRL document."""


class MissingStatementError(FinancialExtractionError):
    """Raised when one or more required statements cannot be extracted."""


@dataclass(frozen=True)
class FilingMetadata:
    """The disclosure metadata used only to create a readable output filename."""

    company_name: str
    report_type: str
    year_month: str

    @property
    def filename_stem(self) -> str:
        company = _safe_filename_part(self.company_name)
        return f"{company}_{self.report_type}_{self.year_month}"


def _safe_filename_part(value: str) -> str:
    cleaned = _INVALID_FILENAME_CHARACTERS.sub("_", value)
    cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
    cleaned = re.sub(r"_+", "_", cleaned)
    if not cleaned:
        raise FinancialExtractionError("회사명을 안전한 파일명으로 만들 수 없습니다.")
    return cleaned


def parse_filing_metadata_html(html_text: str) -> FilingMetadata:
    """Read company, report type, and filing year/month from a DART page title."""

    title_match = re.search(
        r"<title[^>]*>(?P<title>.*?)</title>",
        html_text,
        flags=re.IGNORECASE | re.DOTALL,
    )
    title = title_match.group("title") if title_match else html_text
    title = html.unescape(_HTML_TAG_PATTERN.sub("", title))
    title = " ".join(title.replace("\u00a0", " ").split())

    report_pattern = "|".join(REPORT_TYPES)
    date_pattern = (
        r"(?P<year>(?:19|20)\d{2})[.\-/년]\s*"
        r"(?P<month>\d{1,2})"
        r"(?:[.\-/월]\s*(?:\d{1,2}일?)?)?"
    )
    patterns = (
        re.compile(
            rf"(?P<company>.+?)\s*/\s*"
            rf"(?P<report>{report_pattern})\s*/\s*{date_pattern}"
        ),
        re.compile(
            rf"\[(?P<company>[^\]]+)\]\s*"
            rf"(?P<report>{report_pattern})\s*\(\s*{date_pattern}\s*\)"
        ),
    )
    for pattern in patterns:
        match = pattern.search(title)
        if not match:
            continue
        company_name = match.group("company").strip()
        report_type = match.group("report")
        month = int(match.group("month"))
        if not 1 <= month <= 12:
            break
        return FilingMetadata(
            company_name=company_name,
            report_type=report_type,
            year_month=f"{match.group('year')}.{month:02d}",
        )

    raise FinancialExtractionError(
        "공시 페이지에서 회사명, 보고서 종류, 연월을 찾을 수 없습니다."
    )


def resolve_filing_metadata(
    receipt_number: str,
    dart_module: Any,
) -> FilingMetadata:
    """Download the DART disclosure page and resolve output filename metadata."""

    try:
        response = dart_module.utils.request.get(
            url=DART_MAIN_URL,
            payload={"rcpNo": receipt_number},
            timeout=120,
        )
        response.raise_for_status()
    except Exception as exc:
        raise FinancialExtractionError(
            f"공시 파일명 정보를 확인하지 못했습니다: {exc}"
        ) from exc
    return parse_filing_metadata_html(response.text)


def require_api_key(environ: Mapping[str, str] | None = None) -> str:
    """Read and validate the API key from DART_API_KEY only."""

    source = os.environ if environ is None else environ
    api_key = source.get("DART_API_KEY", "").strip()
    if not api_key:
        raise MissingApiKeyError(
            "DART_API_KEY가 없습니다. GitHub 저장소의 Actions secret에 "
            "DART_API_KEY를 등록하세요."
        )
    return api_key


def _import_dart_fss() -> Any:
    try:
        import dart_fss as dart
    except ImportError as exc:  # pragma: no cover - dependency exists in Actions
        raise FinancialExtractionError(
            "dart-fss가 설치되지 않았습니다. requirements.txt 의존성을 설치하세요."
        ) from exc
    return dart


def _looks_like_missing_xbrl(exc: Exception) -> bool:
    message = str(exc).lower()
    return (
        isinstance(exc, FileNotFoundError)
        or exc.__class__.__name__ == "NoDataReceived"
        or ("xbrl" in message and "not found" in message)
        or "no data received" in message
        or "조회된 데이타가 없습니다" in message
        or "조회된 데이터가 없습니다" in message
        or "파일이 존재하지 않습니다" in message
        or "target does not exist" in message
    )


def load_xbrl(receipt_number: str, download_dir: Path, dart_module: Any) -> Any:
    """Download and load one filing's XBRL through dart-fss."""

    try:
        xbrl_path = dart_module.api.finance.download_xbrl(
            path=str(download_dir),
            rcept_no=receipt_number,
        )
    except Exception as exc:
        if _looks_like_missing_xbrl(exc):
            raise MissingXbrlError(
                f"접수번호 {receipt_number} 공시에서 XBRL 파일을 찾을 수 없습니다."
            ) from exc
        raise FinancialExtractionError(f"XBRL 다운로드 중 오류가 발생했습니다: {exc}") from exc

    if not xbrl_path or not Path(xbrl_path).is_file():
        raise MissingXbrlError(
            f"접수번호 {receipt_number} 공시에서 XBRL 파일을 찾을 수 없습니다."
        )

    try:
        xbrl = dart_module.xbrl.get_xbrl_from_file(str(xbrl_path))
    except Exception as exc:
        if _looks_like_missing_xbrl(exc):
            raise MissingXbrlError(
                f"접수번호 {receipt_number} 공시의 XBRL 파일을 불러올 수 없습니다."
            ) from exc
        raise FinancialExtractionError(f"XBRL 로드 중 오류가 발생했습니다: {exc}") from exc

    is_empty = getattr(xbrl, "is_empty", False) if xbrl is not None else True
    if callable(is_empty):
        is_empty = is_empty()
    if xbrl is None or is_empty:
        raise MissingXbrlError(
            f"접수번호 {receipt_number} 공시에 사용 가능한 XBRL 데이터가 없습니다."
        )
    return xbrl


def _first_table(value: Any) -> Any | None:
    if value is None:
        return None
    if isinstance(value, Sequence) and not isinstance(value, (str, bytes)):
        return next((item for item in value if item is not None), None)
    return value


def _table_to_dataframe(table: Any, *, separate: bool) -> pd.DataFrame | None:
    if table is None:
        return None
    frame = table.to_DataFrame(
        lang="ko",
        label="Separate" if separate else "Consolidated",
        show_abstract=True,
        show_class=False,
        show_concept=False,
        separator=False,
    )
    if frame is None or frame.empty:
        return None
    simplified = simplify_statement_dataframe(frame)
    return simplified if not simplified.empty else None


def _extract_for_scope(
    xbrl: Any,
    *,
    separate: bool,
) -> tuple[dict[str, pd.DataFrame], list[str]]:
    statements: dict[str, pd.DataFrame] = {}
    missing: list[str] = []

    for statement_name, method_name in STATEMENT_SPECS:
        method = getattr(xbrl, method_name)
        table = _first_table(method(separate=separate))
        frame = _table_to_dataframe(table, separate=separate)
        if frame is None:
            missing.append(statement_name)
        else:
            statements[statement_name] = frame

    return statements, missing


def _scope_attempts(xbrl: Any, scope: str) -> list[tuple[str, bool]]:
    if scope == "consolidated":
        return [("consolidated", False)]
    if scope == "separate":
        return [("separate", True)]

    has_consolidated = bool(xbrl.exist_consolidated())
    if has_consolidated:
        return [("consolidated", False), ("separate", True)]
    return [("separate", True)]


def extract_statements(
    xbrl: Any,
    scope: str = "auto",
) -> tuple[dict[str, pd.DataFrame], str]:
    """Extract the required three statements and return their selected scope."""

    if scope not in {"auto", "consolidated", "separate"}:
        raise ValueError(f"지원하지 않는 재무제표 범위입니다: {scope}")

    failures: list[str] = []
    for scope_name, separate in _scope_attempts(xbrl, scope):
        statements, missing = _extract_for_scope(xbrl, separate=separate)
        if not missing:
            return statements, scope_name
        failures.append(f"{SCOPE_LABELS[scope_name]}: {', '.join(missing)}")

    details = "; ".join(failures)
    raise MissingStatementError(
        "필수 재무제표를 모두 찾을 수 없습니다. "
        f"누락된 재무제표 ({details}). "
        "재무상태표, 손익계산서(또는 포괄손익계산서), 현금흐름표가 모두 필요합니다."
    )


def _column_part_text(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass

    text = " ".join(str(value).split())
    if not text or text.lower().startswith("unnamed:"):
        return ""
    return text


def _flatten_column(column: Any) -> str:
    parts = column if isinstance(column, tuple) else (column,)
    cleaned: list[str] = []
    for part in parts:
        nested_parts = part if isinstance(part, tuple) else (part,)
        for nested_part in nested_parts:
            text = _column_part_text(nested_part)
            if text and (not cleaned or cleaned[-1] != text):
                cleaned.append(text)
    return " | ".join(cleaned) or "열"


def _deduplicate_headers(headers: Sequence[str]) -> list[str]:
    counts: dict[str, int] = {}
    unique: list[str] = []
    for header in headers:
        counts[header] = counts.get(header, 0) + 1
        occurrence = counts[header]
        unique.append(header if occurrence == 1 else f"{header} ({occurrence})")
    return unique


def flatten_dataframe_columns(frame: pd.DataFrame) -> pd.DataFrame:
    """Copy a DataFrame and replace any multi-level columns with one header row."""

    flattened = frame.copy()
    headers = [_flatten_column(column) for column in flattened.columns.to_list()]
    flattened.columns = _deduplicate_headers(headers)
    return flattened


_DATE_PATTERN = re.compile(
    r"(?<!\d)((?:19|20)\d{2})[-./년]\s*(\d{1,2})[-./월]\s*(\d{1,2})(?:일)?(?!\d)"
)
_COMPACT_DATE_PATTERN = re.compile(r"(?<!\d)((?:19|20)\d{6})(?!\d)")
_ITEM_COLUMN_NAMES = {
    "labelko",
    "계정",
    "계정과목",
    "계정명",
    "과목",
    "항목",
}
_METADATA_COLUMN_NAMES = {
    "concept",
    "conceptid",
    "labelen",
    "class",
    "주석",
    "note",
    "notes",
}


def _column_parts(column: Any) -> list[str]:
    values = column if isinstance(column, tuple) else (column,)
    parts: list[str] = []
    for value in values:
        nested_values = value if isinstance(value, tuple) else (value,)
        for nested_value in nested_values:
            text = _column_part_text(nested_value)
            if text and text not in parts:
                parts.append(text)
    return parts


def _normalized_header_part(value: str) -> str:
    return re.sub(r"[\s_\-]+", "", value).lower()


def _item_column_score(column: Any) -> int:
    normalized = [_normalized_header_part(part) for part in _column_parts(column)]
    if "labelko" in normalized:
        return 100
    if any(part in {"과목", "항목", "계정과목", "계정명"} for part in normalized):
        return 90
    if "계정" in normalized:
        return 80
    return 0


def _is_metadata_column(column: Any) -> bool:
    normalized = [_normalized_header_part(part) for part in _column_parts(column)]
    return any(
        part in _METADATA_COLUMN_NAMES
        or part.startswith("class")
        or part.startswith("conceptid")
        or part.startswith("주석")
        or part.startswith("note")
        for part in normalized
    )


def _as_excel_number(value: Any) -> int | float | Any:
    if value is None or isinstance(value, bool):
        return "" if value is None else value
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    if isinstance(value, numbers.Number):
        if isinstance(value, float) and value.is_integer():
            return int(value)
        return value

    text = " ".join(str(value).replace("\u00a0", " ").split())
    if not text or text in {"-", "–", "—"}:
        return ""

    negative = text.startswith("(") and text.endswith(")")
    numeric_text = text[1:-1] if negative else text
    numeric_text = numeric_text.replace(",", "")
    if re.fullmatch(r"[+-]?\d+(?:\.\d+)?", numeric_text):
        number: int | float
        number = float(numeric_text) if "." in numeric_text else int(numeric_text)
        return -number if negative else number
    return text


def _series_has_numeric_value(series: pd.Series) -> bool:
    return any(
        isinstance(converted := _as_excel_number(value), numbers.Number)
        and not isinstance(converted, bool)
        for value in series
    )


def _period_header(column: Any) -> str:
    parts = _column_parts(column)
    joined = " ".join(parts)
    dates = [
        f"{year}-{int(month):02d}-{int(day):02d}"
        for year, month, day in _DATE_PATTERN.findall(joined)
    ]
    dates.extend(
        f"{date[:4]}-{date[4:6]}-{date[6:]}"
        for date in _COMPACT_DATE_PATTERN.findall(joined)
    )
    dates = list(dict.fromkeys(dates))
    if dates:
        return dates[0] if len(dates) == 1 else f"{dates[0]} ~ {dates[-1]}"

    preferred_markers = ("당기", "전기", "분기", "반기", "누적", "기말", "연도")
    for part in reversed(parts):
        compact = _normalized_header_part(part)
        if any(marker in compact for marker in preferred_markers):
            return part

    for part in reversed(parts):
        compact = _normalized_header_part(part)
        if (
            compact not in _METADATA_COLUMN_NAMES
            and compact not in _ITEM_COLUMN_NAMES
            and "unit:" not in part.lower()
            and "재무상태표" not in compact
            and "손익계산서" not in compact
            and "현금흐름표" not in compact
        ):
            return part
    return "금액"


def _clean_item_value(value: Any) -> str:
    if value is None:
        return ""
    try:
        if pd.isna(value):
            return ""
    except (TypeError, ValueError):
        pass
    text = " ".join(str(value).replace("\u00a0", " ").split())
    return re.sub(r"\s*\[abstract\]\s*$", "", text, flags=re.IGNORECASE)


def simplify_statement_dataframe(frame: pd.DataFrame) -> pd.DataFrame:
    """Keep one Korean item column and only period columns containing amounts."""

    if frame is None or frame.empty:
        return pd.DataFrame()

    columns = frame.columns.to_list()
    item_index = max(
        range(len(columns)),
        key=lambda index: _item_column_score(columns[index]),
        default=0,
    )
    if _item_column_score(columns[item_index]) == 0:
        item_index = 0

    amount_indices = [
        index
        for index, column in enumerate(columns)
        if index != item_index
        and not _is_metadata_column(column)
        and _series_has_numeric_value(frame.iloc[:, index])
    ]
    if not amount_indices:
        return pd.DataFrame()

    output_columns = ["항목"] + [
        _period_header(columns[index]) for index in amount_indices
    ]
    output_columns = _deduplicate_headers(output_columns)

    series = [frame.iloc[:, item_index].map(_clean_item_value)]
    series.extend(
        frame.iloc[:, index].map(_as_excel_number) for index in amount_indices
    )
    simplified = pd.concat(series, axis=1)
    simplified.columns = output_columns

    populated_rows = simplified.apply(
        lambda row: any(value not in {"", None} for value in row),
        axis=1,
    )
    return simplified.loc[populated_rows].reset_index(drop=True)


def _display_width(value: Any) -> int:
    text = "" if value is None else str(value)
    return sum(
        2 if unicodedata.east_asian_width(character) in {"W", "F"} else 1
        for character in text
    )


def format_worksheet(worksheet: Any) -> None:
    """Apply the small, deterministic formatting set required for outputs."""

    worksheet.freeze_panes = "A2"
    worksheet.sheet_view.showGridLines = False

    for cell in worksheet[1]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(fill_type="solid", fgColor="D9EAF7")
        cell.alignment = Alignment(
            horizontal="center",
            vertical="center",
            wrap_text=True,
        )

    for row in worksheet.iter_rows(min_row=2):
        has_number = False
        for cell in row:
            value = cell.value
            if isinstance(value, numbers.Number) and not isinstance(value, bool):
                if isinstance(value, float) and not math.isfinite(value):
                    continue
                cell.number_format = "#,##0;[Red](#,##0);-"
                cell.alignment = Alignment(horizontal="right")
                has_number = True
            elif cell.column == 1:
                cell.alignment = Alignment(horizontal="left")
        if not has_number and row[0].value:
            row[0].font = Font(bold=True)

    for column_index, column_cells in enumerate(worksheet.columns, start=1):
        content_width = max((_display_width(cell.value) for cell in column_cells), default=0)
        minimum_width = 24 if column_index == 1 else 16
        maximum_width = 45 if column_index == 1 else 24
        worksheet.column_dimensions[get_column_letter(column_index)].width = min(
            max(content_width + 2, minimum_width),
            maximum_width,
        )


def _write_workbook(path: Path, sheets: Mapping[str, pd.DataFrame]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        for sheet_name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=sheet_name, index=False)
            format_worksheet(writer.book[sheet_name])


def export_statements(
    statements: Mapping[str, pd.DataFrame],
    *,
    metadata: FilingMetadata,
    output_mode: str,
    output_dir: Path,
) -> list[Path]:
    """Write one three-sheet workbook or three one-sheet workbooks."""

    missing = [name for name in STATEMENT_NAMES if name not in statements]
    extras = [name for name in statements if name not in STATEMENT_NAMES]
    if missing or extras:
        raise ValueError(
            "내보낼 재무제표는 정확히 재무상태표, 손익계산서, 현금흐름표여야 합니다."
        )

    ordered_statements = {
        name: simplify_statement_dataframe(statements[name])
        for name in STATEMENT_NAMES
    }
    empty = [name for name, frame in ordered_statements.items() if frame.empty]
    if empty:
        raise ValueError(
            "항목과 기간별 금액으로 정리할 수 없는 재무제표가 있습니다: "
            + ", ".join(empty)
        )
    output_dir.mkdir(parents=True, exist_ok=True)
    filename_stem = metadata.filename_stem

    if output_mode == "single":
        workbook_path = output_dir / f"{filename_stem}.xlsx"
        _write_workbook(workbook_path, ordered_statements)
        return [workbook_path]

    if output_mode == "separate":
        paths: list[Path] = []
        for statement_name, frame in ordered_statements.items():
            workbook_path = output_dir / f"{filename_stem}_{statement_name}.xlsx"
            _write_workbook(workbook_path, {statement_name: frame})
            paths.append(workbook_path)
        return paths

    raise ValueError(f"지원하지 않는 출력 방식입니다: {output_mode}")


def run(
    disclosure: str,
    *,
    scope: str = "auto",
    output_mode: str = "single",
    output_dir: Path = Path("output"),
    environ: Mapping[str, str] | None = None,
    dart_module: Any | None = None,
) -> tuple[list[Path], str, str, str]:
    """Run extraction and return paths, selected scope, receipt, and source."""

    reference = parse_disclosure_reference(disclosure)
    receipt_number = reference.receipt_number
    api_key = require_api_key(environ)
    dart = _import_dart_fss() if dart_module is None else dart_module
    dart.set_api_key(api_key=api_key)
    metadata = resolve_filing_metadata(receipt_number, dart)

    with tempfile.TemporaryDirectory(prefix="dart-financials-") as temp_dir:
        temp_path = Path(temp_dir)
        try:
            xbrl = load_xbrl(receipt_number, temp_path, dart)
        except MissingXbrlError as xbrl_error:
            try:
                document_number = reference.document_number or resolve_document_number(
                    receipt_number,
                    dart,
                )
                pdf_path = download_filing_pdf(
                    receipt_number,
                    document_number,
                    temp_path,
                    dart,
                )
                statements, selected_scope = extract_pdf_statements(pdf_path, scope)
            except PdfExtractionError as pdf_error:
                raise FinancialExtractionError(
                    f"{xbrl_error} PDF 대체 추출도 실패했습니다: {pdf_error}"
                ) from pdf_error
            source = "PDF"
        else:
            statements, selected_scope = extract_statements(xbrl, scope)
            source = "XBRL"

        paths = export_statements(
            statements,
            metadata=metadata,
            output_mode=output_mode,
            output_dir=output_dir,
        )

    return paths, selected_scope, receipt_number, source


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description=(
            "DART 공시 XBRL에서 재무제표 3종을 추출하고, "
            "XBRL이 없으면 공시 PDF를 사용합니다."
        )
    )
    parser.add_argument("disclosure", help="DART 공시 URL 또는 14자리 접수번호")
    parser.add_argument(
        "--scope",
        choices=("auto", "consolidated", "separate"),
        default="auto",
        help="auto는 연결 3종을 우선하고 불가능하면 별도 3종으로 재시도합니다.",
    )
    parser.add_argument(
        "--output-mode",
        choices=("single", "separate"),
        default="single",
        help="single은 3개 시트 통합 파일, separate는 개별 파일 3개입니다.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("output"),
        help="Excel 파일을 저장할 폴더(기본값: output)",
    )
    return parser


def main(argv: Sequence[str] | None = None) -> int:
    args = build_argument_parser().parse_args(argv)
    try:
        paths, selected_scope, receipt_number, source = run(
            args.disclosure,
            scope=args.scope,
            output_mode=args.output_mode,
            output_dir=args.output_dir,
        )
    except (
        InvalidDisclosureReference,
        FinancialExtractionError,
        ValueError,
    ) as exc:
        print(f"오류: {exc}", file=sys.stderr)
        return 1

    print(f"접수번호: {receipt_number}")
    print(f"추출 원본: {source}")
    print(f"재무제표 범위: {SCOPE_LABELS[selected_scope]}")
    for path in paths:
        print(f"생성 완료: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
