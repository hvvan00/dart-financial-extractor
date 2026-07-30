# Repository constraints

- Keep this repository as a GitHub Actions-driven command-line utility. Do not
  add a web server, database, desktop application, or other user interface.
- Accept only an official DART disclosure URL with `rcpNo` or a raw 14-digit
  receipt number.
- Read the Open DART key only from the `DART_API_KEY` environment variable. In
  GitHub Actions it must come only from the `DART_API_KEY` repository secret;
  never add it as a workflow input, command-line argument, file, log value, or
  source-code constant.
- Use `dart-fss` to download and load the filing XBRL first. Only when XBRL is
  unavailable, use the DART full-report PDF as a strict fallback.
- The PDF fallback may extract text-based tables with `pdfplumber`. It must
  fail clearly for scanned/image-only PDFs or incomplete statement sets rather
  than guessing or mixing tables.
- Export exactly these statements: statement of financial position, income or
  comprehensive income statement, and cash flow statement. Never add statement
  of changes in equity or unrelated DART data.
- Auto scope must prefer a complete consolidated set and fall back to a
  complete separate set without mixing scopes.
- Single output must contain exactly `재무상태표`, `손익계산서`, and
  `현금흐름표` sheets. Separate output must contain exactly three files.
- Name one-filing single outputs `회사명_보고서종류_YYYY.MM.xlsx`. Name
  multi-filing single outputs `회사명_보고서종류_YYYY-YYYY.xlsx`, using the
  earliest and latest financial-statement period years. For separate mode,
  append the Korean statement name before `.xlsx`. Report type must be one of
  `사업보고서`, `반기보고서`, `분기보고서`, or `감사보고서`.
- Multiple inputs must be the same company, report type, and selected scope.
  Keep only the primary current period from each filing, label annual columns
  by year and interim columns by quarter or half-year, and do not repeat each
  filing's comparative-period columns. Merge rows conservatively by normalized
  Korean item labels and their statement hierarchy. Merge known aliases only
  within the same section and when their populated periods do not overlap.
  Infer the `매출액` or `매출원가` parent for unambiguous revenue/cost detail
  labels when an older PDF omits the visible parent row, while keeping the
  total row separate from its details. Fill a blank parent-period amount from
  unique numeric direct details, but never overwrite a disclosed total or
  calculate when duplicate detail keys could cause double counting.
  Repeated contra/subsidy labels must also have the same immediate parent.
  Keep genuinely renamed items separate.
- Each exported statement must contain only one Korean `항목` column followed
  by period amount columns. Exclude concept IDs, English labels, class/category
  columns, and note-reference columns from user-facing Excel files.
- Keep XBRL numeric values as numeric Excel cells. Flatten pandas multi-level
  columns before export, convert complete PDF amount strings to numeric cells,
  and retain the required basic workbook formatting. Visually distinguish
  major sections, subsections, detail rows, and totals with restrained fills,
  borders, bold text, and indentation. Use `#006074`, `#008187`, and
  `#06A39F` as the highlight palette, with white text on dark highlight fills.
  Remove source-specific Roman, parenthesized, Arabic, and Korean numbering
  from user-facing item labels. Apply consistent semantic subgroup styling and
  indentation across all three statements even when filings use different
  numbering styles. Prefix only true balance-sheet child/contra accounts with
  `└`.
- Keep the manual `workflow_dispatch` workflow on Python 3.12 and upload outputs
  with `actions/upload-artifact@v4`.

# Test commands

Run both commands after code changes:

```text
python -m unittest discover -s tests -v
python -m py_compile dart_link.py extract_financials.py
```
