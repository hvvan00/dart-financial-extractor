"""Strict PDF fallback for DART filings that do not provide XBRL."""

from __future__ import annotations

import re
from collections import Counter
from collections.abc import Sequence
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import pandas as pd


PDF_DOWNLOAD_URL = "https://dart.fss.or.kr/pdf/download/pdf.do"
PDF_DOWNLOAD_MAIN_URL = "https://dart.fss.or.kr/pdf/download/main.do"
DART_MAIN_URL = "https://dart.fss.or.kr/dsaf001/main.do"
STATEMENT_NAMES = ("재무상태표", "손익계산서", "현금흐름표")
SCOPE_LABELS = {"consolidated": "연결", "separate": "별도"}
MAX_STATEMENT_PAGES = 6

_TITLE_PATTERNS = (
    (
        "재무상태표",
        re.compile(
            r"^(?P<scope>연결|별도)?(?:재무상태표|대차대조표)"
            r"(?P<continued>\(?계속\)?)?$"
        ),
    ),
    (
        "손익계산서",
        re.compile(
            r"^(?P<scope>연결|별도)?(?:포괄손익계산서|손익계산서)"
            r"(?P<continued>\(?계속\)?)?$"
        ),
    ),
    (
        "현금흐름표",
        re.compile(
            r"^(?P<scope>연결|별도)?현금흐름표"
            r"(?P<continued>\(?계속\)?)?$"
        ),
    ),
)
_BOUNDARY_PATTERNS = (
    re.compile(r"^(?:연결|별도)?자본변동표(?:\(?계속\)?)?$"),
    re.compile(r"^(?:연결|별도)?재무제표에대한주석$"),
    re.compile(r"^\d+\.?(?:재무제표)?주석$"),
    re.compile(r"^주석$"),
)
_NUMBER_PATTERN = re.compile(r"^[+-]?(?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?$")

_LINE_TABLE_SETTINGS = {
    "vertical_strategy": "lines",
    "horizontal_strategy": "lines",
    "snap_tolerance": 3,
    "join_tolerance": 3,
    "intersection_tolerance": 5,
}
_TEXT_TABLE_SETTINGS = {
    "vertical_strategy": "text",
    "horizontal_strategy": "text",
    "min_words_vertical": 2,
    "min_words_horizontal": 1,
    "text_x_tolerance": 2,
    "text_y_tolerance": 3,
}


class PdfExtractionError(RuntimeError):
    """Raised when a DART PDF cannot produce a complete statement set."""


@dataclass(frozen=True)
class PdfTitleEvent:
    page_index: int
    top: float
    bottom: float
    statement_name: str | None
    scope: str | None
    continued: bool = False


def _compact_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", "", str(value)).strip()


def _clean_cell(value: Any) -> str:
    if value is None:
        return ""
    return " ".join(str(value).replace("\u00a0", " ").split())


def parse_pdf_number(value: Any) -> int | float | None:
    """Convert one complete Korean financial amount string to a number."""

    if isinstance(value, bool) or value is None:
        return None
    if isinstance(value, (int, float)):
        return value

    text = _compact_text(value)
    if not text or text in {"-", "–", "—"}:
        return None

    negative = False
    if text.startswith("(") and text.endswith(")"):
        negative = True
        text = text[1:-1]
    elif text.startswith("△"):
        negative = True
        text = text[1:]

    if not _NUMBER_PATTERN.fullmatch(text):
        return None

    normalized = text.replace(",", "")
    number: int | float
    if "." in normalized:
        number = float(normalized)
    else:
        number = int(normalized)
    return -number if negative else number


def resolve_document_number(receipt_number: str, dart_module: Any) -> str:
    """Resolve DART's internal document number from a receipt-number page."""

    try:
        response = dart_module.utils.request.get(
            url=DART_MAIN_URL,
            payload={"rcpNo": receipt_number},
            timeout=120,
        )
        response.raise_for_status()
        html = response.text
    except Exception as exc:
        raise PdfExtractionError(
            f"PDF 문서번호(dcmNo)를 확인하지 못했습니다: {exc}"
        ) from exc

    patterns = (
        re.compile(
            rf"viewDoc\(\s*['\"]{re.escape(receipt_number)}['\"]\s*,"
            r"\s*['\"](?P<dcm>\d{5,12})['\"]"
        ),
        re.compile(r"(?:dcmNo|dcm_no)\s*[=:]\s*['\"]?(?P<dcm>\d{5,12})"),
    )
    for pattern in patterns:
        match = pattern.search(html)
        if match:
            return match.group("dcm")

    raise PdfExtractionError(
        "공시 페이지에서 PDF 문서번호(dcmNo)를 찾을 수 없습니다. "
        "가능하면 dcmNo가 포함된 DART URL을 입력하세요."
    )


def download_filing_pdf(
    receipt_number: str,
    document_number: str,
    download_dir: Path,
    dart_module: Any,
) -> Path:
    """Download the full disclosure PDF through dart-fss's request session."""

    filename = f"DART_{receipt_number}_{document_number}.pdf"
    try:
        # DART's download popup initializes the PDF download session.  Calling
        # pdf.do directly can return HTTP 200 with an empty response body.
        landing_response = dart_module.utils.request.get(
            url=PDF_DOWNLOAD_MAIN_URL,
            payload={
                "rcp_no": receipt_number,
                "dcm_no": document_number,
            },
            referer=f"{DART_MAIN_URL}?rcpNo={receipt_number}",
            timeout=120,
        )
        landing_response.raise_for_status()
        download_referer = (
            f"{PDF_DOWNLOAD_MAIN_URL}?rcp_no={receipt_number}"
            f"&dcm_no={document_number}"
        )
        result = dart_module.utils.request.download(
            url=PDF_DOWNLOAD_URL,
            path=str(download_dir),
            filename=filename,
            payload={
                "rcp_no": receipt_number,
                "dcm_no": document_number,
            },
            referer=download_referer,
            timeout=120,
        )
    except Exception as exc:
        raise PdfExtractionError(f"DART 공시 PDF 다운로드에 실패했습니다: {exc}") from exc

    pdf_path = Path(result["full_path"] if isinstance(result, dict) else result)
    if not pdf_path.is_file() or pdf_path.stat().st_size == 0:
        raise PdfExtractionError("DART에서 PDF 파일을 내려받지 못했습니다.")

    with pdf_path.open("rb") as stream:
        header = stream.read(1024)
    if b"%PDF-" not in header:
        raise PdfExtractionError(
            "DART 응답이 PDF 형식이 아닙니다. 공시 URL과 dcmNo를 확인하세요."
        )
    return pdf_path


def _page_lines(page: Any) -> list[tuple[str, float, float]]:
    return [
        (
            " ".join(str(item["text"]) for item in group),
            min(float(item["top"]) for item in group),
            max(float(item["bottom"]) for item in group),
        )
        for group in _group_page_words(page)
    ]


def _group_page_words(page: Any) -> list[list[dict[str, Any]]]:
    words = page.extract_words(
        x_tolerance=2,
        y_tolerance=3,
        keep_blank_chars=False,
        use_text_flow=True,
    )
    if not words:
        return []

    grouped: list[list[dict[str, Any]]] = []
    for word in sorted(words, key=lambda item: (float(item["top"]), float(item["x0"]))):
        if not grouped or abs(float(word["top"]) - float(grouped[-1][0]["top"])) > 3:
            grouped.append([word])
        else:
            grouped[-1].append(word)

    for group in grouped:
        group.sort(key=lambda item: float(item["x0"]))
    return grouped


def _match_title(text: str) -> tuple[str, str, bool] | None:
    compact = _compact_text(text)
    for statement_name, pattern in _TITLE_PATTERNS:
        match = pattern.fullmatch(compact)
        if match:
            scope = "consolidated" if match.group("scope") == "연결" else "separate"
            return statement_name, scope, bool(match.group("continued"))
    return None


def _is_boundary_title(text: str) -> bool:
    compact = _compact_text(text)
    return any(pattern.fullmatch(compact) for pattern in _BOUNDARY_PATTERNS)


def find_pdf_title_events(pdf: Any) -> list[PdfTitleEvent]:
    """Locate target statement titles and excluded-statement boundaries."""

    events: list[PdfTitleEvent] = []
    for page_index, page in enumerate(pdf.pages):
        for text, top, bottom in _page_lines(page):
            matched = _match_title(text)
            if matched:
                statement_name, scope, continued = matched
                events.append(
                    PdfTitleEvent(
                        page_index=page_index,
                        top=top,
                        bottom=bottom,
                        statement_name=statement_name,
                        scope=scope,
                        continued=continued,
                    )
                )
            elif _is_boundary_title(text):
                events.append(
                    PdfTitleEvent(
                        page_index=page_index,
                        top=top,
                        bottom=bottom,
                        statement_name=None,
                        scope=None,
                    )
                )
    return sorted(events, key=lambda event: (event.page_index, event.top))


def _event_is_after(candidate: PdfTitleEvent, start: PdfTitleEvent) -> bool:
    return (candidate.page_index, candidate.top) > (start.page_index, start.top)


def _next_boundary(
    start: PdfTitleEvent,
    events: Sequence[PdfTitleEvent],
) -> PdfTitleEvent | None:
    for event in events:
        if not _event_is_after(event, start):
            continue
        is_same_continuation = (
            event.continued
            and event.statement_name == start.statement_name
            and event.scope == start.scope
        )
        if not is_same_continuation:
            return event
    return None


def _statement_regions(
    pdf: Any,
    start: PdfTitleEvent,
    events: Sequence[PdfTitleEvent],
) -> list[Any]:
    boundary = _next_boundary(start, events)
    last_page = min(start.page_index + MAX_STATEMENT_PAGES - 1, len(pdf.pages) - 1)
    if boundary is not None:
        last_page = min(last_page, boundary.page_index)

    continuations = {
        event.page_index: event
        for event in events
        if event.continued
        and event.statement_name == start.statement_name
        and event.scope == start.scope
        and _event_is_after(event, start)
    }

    regions: list[Any] = []
    for page_index in range(start.page_index, last_page + 1):
        page = pdf.pages[page_index]
        top = 0.0
        bottom = float(page.height)

        if page_index == start.page_index:
            top = start.bottom
        elif page_index in continuations:
            top = continuations[page_index].bottom

        if boundary is not None and page_index == boundary.page_index:
            bottom = max(top, boundary.top - 0.5)

        if bottom - top > 5:
            regions.append(page.crop((0, top, float(page.width), bottom)))
    return regions


def _clean_matrix(table: Sequence[Sequence[Any]]) -> list[list[str]]:
    rows = [[_clean_cell(cell) for cell in row] for row in table if row]
    rows = [row for row in rows if any(row)]
    if not rows:
        return []

    max_columns = max(len(row) for row in rows)
    padded = [row + [""] * (max_columns - len(row)) for row in rows]
    keep_columns = [
        column_index
        for column_index in range(max_columns)
        if any(row[column_index] for row in padded)
    ]
    return [[row[column_index] for column_index in keep_columns] for row in padded]


def _matrix_score(matrix: Sequence[Sequence[str]]) -> int:
    if len(matrix) < 3 or max((len(row) for row in matrix), default=0) < 2:
        return -1
    numeric_cells = sum(
        parse_pdf_number(cell) is not None for row in matrix for cell in row
    )
    if numeric_cells < 3:
        return -1
    nonempty_cells = sum(bool(cell) for row in matrix for cell in row)
    return numeric_cells * 5 + nonempty_cells + len(matrix) * 2


def _word_center(word: dict[str, Any]) -> float:
    return (float(word["x0"]) + float(word["x1"])) / 2


def _is_amount_word(word: dict[str, Any]) -> bool:
    text = _clean_cell(word["text"])
    return parse_pdf_number(text) is not None or _compact_text(text) in {"-", "–", "—"}


def _select_period_amount(
    candidates: Sequence[dict[str, Any]],
) -> dict[str, Any] | None:
    """Prefer a real amount over a farther-right blank placeholder."""

    numeric_candidates = [
        word
        for word in candidates
        if parse_pdf_number(_clean_cell(word["text"])) is not None
    ]
    selectable = numeric_candidates or list(candidates)
    return max(selectable, key=_word_center, default=None)


def _merge_wrapped_matrix_rows(matrix: list[list[str]]) -> list[list[str]]:
    """Join one visual table row that PDF text extraction split vertically."""

    if len(matrix) < 3:
        return matrix

    merged = [matrix[0]]
    index = 1
    while index < len(matrix):
        row = matrix[index].copy()
        values = row[-2:]
        has_values = any(_compact_text(value) for value in values)
        next_row = matrix[index + 1] if index + 1 < len(matrix) else None
        next_has_values = bool(
            next_row
            and any(_compact_text(value) for value in next_row[-2:])
        )
        next_has_account = bool(next_row and _compact_text(next_row[0]))

        if (
            not has_values
            and next_row is not None
            and next_has_values
            and not next_has_account
        ):
            row[-2:] = next_row[-2:]
            index += 2

            open_parentheses = row[0].count("(") - row[0].count(")")
            if open_parentheses > 0 and index < len(matrix):
                continuation = matrix[index]
                continuation_has_values = any(
                    _compact_text(value) for value in continuation[-2:]
                )
                if not continuation_has_values and _compact_text(continuation[0]):
                    row[0] = f"{row[0]} {continuation[0]}".strip()
                    index += 1
            merged.append(row)
            continue

        merged.append(row)
        index += 1
    return merged


def _line_text(line: Sequence[dict[str, Any]]) -> str:
    return " ".join(_clean_cell(word["text"]) for word in line)


def _described_periods(
    lines: Sequence[Sequence[dict[str, Any]]],
) -> list[str]:
    periods: list[str] = []
    for line in lines:
        text = _line_text(line)
        compact = _compact_text(text)
        if (
            compact.startswith("제")
            and re.search(r"(?:19|20)\d{2}년", compact)
            and any(marker in compact for marker in ("현재", "부터", "까지"))
        ):
            periods.append(text)
    return periods


def _region_word_matrix(region: Any) -> list[list[str]] | None:
    """Recover a clean table by assigning words to visible period columns."""

    try:
        lines = _group_page_words(region)
    except Exception:
        return None

    header_index = next(
        (
            index
            for index, line in enumerate(lines)
            if "과목" in _compact_text(" ".join(str(word["text"]) for word in line))
            and sum(
                _compact_text(word["text"]).startswith("제")
                for word in line
            )
            >= 2
        ),
        None,
    )
    if header_index is None:
        loose_header_index = next(
            (
                index
                for index, line in enumerate(lines)
                if "과목" in _compact_text(_line_text(line))
            ),
            None,
        )
        numeric_lines = [
            [
                word
                for word in line
                if _word_center(word) >= float(region.width) * 0.35
                and _is_amount_word(word)
            ]
            for line in lines
        ]
        if sum(len(words) >= 2 for words in numeric_lines) < 2:
            return None

        current_floor = float(region.width) * 0.35
        period_boundary = float(region.width) * 0.65
        note_word = None
        note_anchor = None
        account_boundary = None
        described_periods = _described_periods(
            lines[:loose_header_index]
            if loose_header_index is not None
            else lines
        )
        if len(described_periods) >= 2:
            header = ["과목", *described_periods[-2:]]
        else:
            header = ["과목", "당기", "전기"]
        data_lines = (
            lines[loose_header_index + 1 :]
            if loose_header_index is not None
            else lines
        )
    else:
        header_words = lines[header_index]
        period_starts = [
            word
            for word in header_words
            if _compact_text(word["text"]).startswith("제")
        ]
        if len(period_starts) < 2:
            return None
        current_floor = float(period_starts[-2]["x0"])
        period_boundary = float(period_starts[-1]["x0"])

        note_word = next(
            (
                word
                for word in header_words
                if _compact_text(word["text"]) == "주석"
            ),
            None,
        )
        account_words = [
            word
            for word in header_words
            if _compact_text(word["text"]) in {"과", "목", "과목"}
        ]
        if not account_words:
            return None
        account_anchor = sum(
            _word_center(word) for word in account_words
        ) / len(account_words)

        if note_word is not None:
            note_anchor = _word_center(note_word)
            account_boundary = (account_anchor + note_anchor) / 2
        else:
            note_anchor = None
            account_boundary = None

        current_header = " ".join(
            _clean_cell(word["text"])
            for word in header_words
            if current_floor <= _word_center(word) < period_boundary
        )
        prior_header = " ".join(
            _clean_cell(word["text"])
            for word in header_words
            if _word_center(word) >= period_boundary
        )
        described_periods = _described_periods(lines[:header_index])
        if len(described_periods) >= 2:
            current_header, prior_header = described_periods[-2:]

        header = ["과목"]
        if note_word is not None:
            header.append("주석")
        header.extend([current_header or "당기", prior_header or "전기"])
        data_lines = lines[header_index + 1 :]

    matrix: list[list[str]] = [header]
    for line in data_lines:
        amount_words = [
            word
            for word in line
            if _word_center(word) >= current_floor and _is_amount_word(word)
        ]
        current_candidates = [
            word for word in amount_words if _word_center(word) < period_boundary
        ]
        prior_candidates = [
            word for word in amount_words if _word_center(word) >= period_boundary
        ]
        current_word = _select_period_amount(current_candidates)
        prior_word = _select_period_amount(prior_candidates)
        selected_amount_ids = {id(word) for word in amount_words}

        note_parts: list[str] = []
        account_parts: list[str] = []
        for word in line:
            if id(word) in selected_amount_ids:
                continue
            center = _word_center(word)
            text = _clean_cell(word["text"])
            if (
                note_anchor is not None
                and account_boundary is not None
                and account_boundary <= center < current_floor
            ):
                note_parts.append(text)
            else:
                account_parts.append(text)

        account = " ".join(part for part in account_parts if part)
        note = " ".join(part for part in note_parts if part)
        current = _clean_cell(current_word["text"]) if current_word else ""
        prior = _clean_cell(prior_word["text"]) if prior_word else ""

        compact_row = _compact_text(account + note + current + prior)
        if not compact_row:
            continue
        if any(
            marker in compact_row
            for marker in ("dart.fss.or.kr", "감사보고서제출", "첨부된주석은")
        ):
            continue

        row = [account]
        if note_word is not None:
            row.append(note)
        row.extend([current, prior])
        matrix.append(row)

    matrix = _merge_wrapped_matrix_rows(matrix)
    return matrix if _matrix_score(matrix) >= 0 else None


def _best_region_matrix(region: Any) -> list[list[str]] | None:
    word_matrix = _region_word_matrix(region)
    if word_matrix is not None:
        return word_matrix

    candidates: list[tuple[int, list[list[str]]]] = []
    for settings in (_LINE_TABLE_SETTINGS, _TEXT_TABLE_SETTINGS):
        try:
            tables = region.extract_tables(table_settings=settings) or []
        except Exception:
            continue
        for table in tables:
            matrix = _clean_matrix(table)
            score = _matrix_score(matrix)
            if score >= 0:
                candidates.append((score, matrix))
    if not candidates:
        return None
    return max(candidates, key=lambda item: item[0])[1]


def _unique_headers(headers: Sequence[str]) -> list[str]:
    counts: dict[str, int] = {}
    result: list[str] = []
    for index, header in enumerate(headers, start=1):
        name = header or f"열 {index}"
        counts[name] = counts.get(name, 0) + 1
        result.append(name if counts[name] == 1 else f"{name} ({counts[name]})")
    return result


def dataframe_from_pdf_matrices(
    matrices: Sequence[Sequence[Sequence[Any]]],
) -> pd.DataFrame | None:
    """Turn compatible page tables into one strict, numeric DataFrame."""

    cleaned = [_clean_matrix(matrix) for matrix in matrices]
    cleaned = [matrix for matrix in cleaned if matrix]
    if not cleaned:
        return None

    column_counts = Counter(len(matrix[0]) for matrix in cleaned)
    target_columns = column_counts.most_common(1)[0][0]
    compatible = [
        matrix
        for matrix in cleaned
        if matrix and all(len(row) == target_columns for row in matrix)
    ]
    if not compatible:
        return None

    rows = [row for matrix in compatible for row in matrix]
    explicit_header_index = next(
        (
            index
            for index, row in enumerate(rows)
            if any(_compact_text(cell) in {"과목", "구분"} for cell in row)
        ),
        None,
    )
    header_index = explicit_header_index if explicit_header_index is not None else 0
    if explicit_header_index is not None:
        first_data_index = header_index + 1
    else:
        first_data_index = next(
            (
                index
                for index in range(header_index + 1, len(rows))
                if sum(parse_pdf_number(cell) is not None for cell in rows[index]) >= 1
            ),
            None,
        )
    if first_data_index is None:
        return None

    header_rows = rows[header_index:first_data_index]
    headers: list[str] = []
    for column_index in range(target_columns):
        parts: list[str] = []
        for row in header_rows:
            text = _clean_cell(row[column_index])
            if text and (not parts or parts[-1] != text):
                parts.append(text)
        headers.append(" | ".join(parts))
    headers = _unique_headers(headers)

    data_rows: list[list[Any]] = []
    for row in rows[first_data_index:]:
        compact_row = "".join(_compact_text(cell) for cell in row)
        if not compact_row:
            continue
        if any(
            marker in compact_row
            for marker in ("첨부된주석은", "dart.fss.or.kr", "감사보고서제출")
        ):
            continue
        if (
            sum(parse_pdf_number(cell) is not None for cell in row) == 0
            and any(_compact_text(cell) in {"과목", "구분"} for cell in row)
        ):
            continue

        converted: list[Any] = []
        for cell in row:
            numeric = parse_pdf_number(cell)
            converted.append(numeric if numeric is not None else _clean_cell(cell))
        if any(value != "" for value in converted):
            data_rows.append(converted)

    if sum(
        isinstance(value, (int, float)) and not isinstance(value, bool)
        for row in data_rows
        for value in row
    ) < 3:
        return None

    frame = pd.DataFrame(data_rows, columns=headers)
    frame = frame.dropna(axis=1, how="all")
    return frame if not frame.empty else None


def _extract_candidate(
    pdf: Any,
    event: PdfTitleEvent,
    events: Sequence[PdfTitleEvent],
) -> pd.DataFrame | None:
    matrices: list[list[list[str]]] = []
    for region in _statement_regions(pdf, event, events):
        matrix = _best_region_matrix(region)
        if matrix is not None:
            matrices.append(matrix)
    return dataframe_from_pdf_matrices(matrices)


def _scope_order(scope: str) -> tuple[str, ...]:
    if scope == "auto":
        return ("consolidated", "separate")
    if scope in {"consolidated", "separate"}:
        return (scope,)
    raise ValueError(f"지원하지 않는 재무제표 범위입니다: {scope}")


def extract_pdf_statements(
    pdf_path: Path,
    scope: str = "auto",
) -> tuple[dict[str, pd.DataFrame], str]:
    """Extract a complete three-statement family from a text-based PDF."""

    try:
        import pdfplumber
    except ImportError as exc:  # pragma: no cover - dependency exists in Actions
        raise PdfExtractionError(
            "PDF 추출에 필요한 pdfplumber가 설치되지 않았습니다."
        ) from exc

    try:
        pdf_context = pdfplumber.open(pdf_path)
    except Exception as exc:
        raise PdfExtractionError(f"PDF 파일을 열 수 없습니다: {exc}") from exc

    with pdf_context as pdf:
        events = find_pdf_title_events(pdf)
        if not events:
            raise PdfExtractionError(
                "PDF에서 재무제표 제목을 찾지 못했습니다. "
                "스캔 이미지 PDF는 현재 자동 인식할 수 없습니다."
            )

        failures: list[str] = []
        for scope_name in _scope_order(scope):
            statements: dict[str, pd.DataFrame] = {}
            missing: list[str] = []
            for statement_name in STATEMENT_NAMES:
                candidates = [
                    event
                    for event in events
                    if event.statement_name == statement_name
                    and event.scope == scope_name
                    and not event.continued
                ]
                frame = next(
                    (
                        extracted
                        for event in candidates
                        if (extracted := _extract_candidate(pdf, event, events))
                        is not None
                    ),
                    None,
                )
                if frame is None:
                    missing.append(statement_name)
                else:
                    statements[statement_name] = frame

            if not missing:
                return statements, scope_name
            failures.append(f"{SCOPE_LABELS[scope_name]}: {', '.join(missing)}")

    raise PdfExtractionError(
        "PDF에서 필수 재무제표 3종을 모두 추출하지 못했습니다 "
        f"({'; '.join(failures)}). "
        "텍스트 선택이 가능한 PDF인지, 표 선과 열이 명확한지 확인하세요."
    )


__all__ = [
    "PdfExtractionError",
    "dataframe_from_pdf_matrices",
    "download_filing_pdf",
    "extract_pdf_statements",
    "find_pdf_title_events",
    "parse_pdf_number",
    "resolve_document_number",
]
