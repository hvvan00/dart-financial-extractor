"""Extract exactly three financial statements from a DART filing XBRL."""

from __future__ import annotations

import argparse
import math
import numbers
import os
import sys
import tempfile
import unicodedata
from collections.abc import Mapping, Sequence
from pathlib import Path
from typing import Any

import pandas as pd
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from dart_link import InvalidDisclosureReference, parse_receipt_number


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


class FinancialExtractionError(RuntimeError):
    """Base error for expected extraction failures."""


class MissingApiKeyError(FinancialExtractionError):
    """Raised when DART_API_KEY is absent."""


class MissingXbrlError(FinancialExtractionError):
    """Raised when the filing does not provide a usable XBRL document."""


class MissingStatementError(FinancialExtractionError):
    """Raised when one or more required statements cannot be extracted."""


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
    return isinstance(exc, FileNotFoundError) or exc.__class__.__name__ == "NoDataReceived" or (
        ("xbrl" in message and "not found" in message)
        or "no data received" in message
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


def _table_to_dataframe(table: Any) -> pd.DataFrame | None:
    if table is None:
        return None
    frame = table.to_DataFrame(
        lang="ko",
        show_abstract=False,
        show_class=True,
        show_concept=True,
        separator=False,
    )
    if frame is None or frame.empty:
        return None
    return flatten_dataframe_columns(frame)


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
        frame = _table_to_dataframe(table)
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


def _display_width(value: Any) -> int:
    text = "" if value is None else str(value)
    return sum(
        2 if unicodedata.east_asian_width(character) in {"W", "F"} else 1
        for character in text
    )


def format_worksheet(worksheet: Any) -> None:
    """Apply the small, deterministic formatting set required for outputs."""

    worksheet.freeze_panes = "A2"
    worksheet.auto_filter.ref = worksheet.dimensions

    for cell in worksheet[1]:
        cell.font = Font(bold=True)

    for row in worksheet.iter_rows(min_row=2):
        for cell in row:
            value = cell.value
            if isinstance(value, numbers.Number) and not isinstance(value, bool):
                if isinstance(value, float) and not math.isfinite(value):
                    continue
                cell.number_format = "#,##0.00;[Red]-#,##0.00"

    for column_index, column_cells in enumerate(worksheet.columns, start=1):
        content_width = max((_display_width(cell.value) for cell in column_cells), default=0)
        worksheet.column_dimensions[get_column_letter(column_index)].width = min(
            max(content_width + 2, 10),
            45,
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
    receipt_number: str,
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

    ordered_statements = {name: statements[name] for name in STATEMENT_NAMES}
    output_dir.mkdir(parents=True, exist_ok=True)

    if output_mode == "single":
        workbook_path = output_dir / f"DART_{receipt_number}_재무제표.xlsx"
        _write_workbook(workbook_path, ordered_statements)
        return [workbook_path]

    if output_mode == "separate":
        paths: list[Path] = []
        for statement_name, frame in ordered_statements.items():
            workbook_path = output_dir / f"DART_{receipt_number}_{statement_name}.xlsx"
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
) -> tuple[list[Path], str, str]:
    """Run the extraction pipeline and return paths, selected scope, and receipt."""

    receipt_number = parse_receipt_number(disclosure)
    api_key = require_api_key(environ)
    dart = _import_dart_fss() if dart_module is None else dart_module
    dart.set_api_key(api_key=api_key)

    with tempfile.TemporaryDirectory(prefix="dart-xbrl-") as temp_dir:
        xbrl = load_xbrl(receipt_number, Path(temp_dir), dart)
        statements, selected_scope = extract_statements(xbrl, scope)
        paths = export_statements(
            statements,
            receipt_number=receipt_number,
            output_mode=output_mode,
            output_dir=output_dir,
        )

    return paths, selected_scope, receipt_number


def build_argument_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="DART 공시 XBRL에서 재무제표 3종을 Excel로 추출합니다."
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
        paths, selected_scope, receipt_number = run(
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
    print(f"재무제표 범위: {SCOPE_LABELS[selected_scope]}")
    for path in paths:
        print(f"생성 완료: {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
