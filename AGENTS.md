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
- Each exported statement must contain only one Korean `항목` column followed
  by period amount columns. Exclude concept IDs, English labels, class/category
  columns, and note-reference columns from user-facing Excel files.
- Keep XBRL numeric values as numeric Excel cells. Flatten pandas multi-level
  columns before export, convert complete PDF amount strings to numeric cells,
  and retain the required basic workbook formatting.
- Keep the manual `workflow_dispatch` workflow on Python 3.12 and upload outputs
  with `actions/upload-artifact@v4`.

# Test commands

Run both commands after code changes:

```text
python -m unittest discover -s tests -v
python -m py_compile dart_link.py extract_financials.py
```
