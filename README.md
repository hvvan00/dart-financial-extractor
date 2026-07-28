# DART 재무제표 Excel 추출기

DART 공시 URL 또는 14자리 접수번호를 입력하면 아래 재무제표 3종만 추출하여
Excel 파일로 내려받을 수 있는 GitHub Actions 저장소입니다. XBRL을 우선 사용하고,
XBRL이 없는 공시는 DART의 전체 보고서 PDF에서 표를 찾아 다시 시도합니다.

- 재무상태표
- 손익계산서 또는 포괄손익계산서
- 현금흐름표

웹 서버, 데이터베이스, 데스크톱 프로그램은 없습니다. 사용자의 PC에 Python을
설치할 필요도 없습니다. GitHub Actions가 `dart-fss`로 XBRL과 공시 PDF를
내려받고 결과 파일을 아티팩트로 제공합니다.

## 0. GitHub 웹사이트에 처음 올리기

GitHub Desktop이나 로컬 Python은 필요하지 않습니다.

1. GitHub 웹사이트에서 빈 저장소를 만든 뒤 **Add file → Upload files**를
   누릅니다.
2. 이 프로젝트의 파일과 `tests` 폴더를 업로드하고 **Commit changes**를
   누릅니다.
3. `.github` 폴더가 숨김 파일로 표시되어 선택되지 않으면
   **Add file → Create new file**을 누릅니다.
4. 파일 이름 칸에 아래 경로 전체를 입력합니다. `/`를 입력하면 GitHub가 폴더를
   자동으로 만듭니다.

```text
.github/workflows/extract-financials.yml
```

5. 이 프로젝트의 같은 경로에 있는 파일 내용을 붙여넣고
   **Commit changes**를 누릅니다.

`.github`는 점(`.`)으로 시작해서 일부 파일 선택 화면에서 숨겨질 수 있지만,
GitHub 웹사이트에서 위 경로로 직접 만들면 정상적인 Actions 워크플로 파일이
됩니다.

## 1. 사전 준비

1. [Open DART](https://opendart.fss.or.kr/)에서 API 인증키를 발급받습니다.
2. 이 저장소의 **Settings → Secrets and variables → Actions**로 이동합니다.
3. **New repository secret**을 누르고 아래와 같이 등록합니다.
   - Name: `DART_API_KEY`
   - Secret: 발급받은 Open DART API 인증키
4. 저장소의 **Actions** 탭에서 워크플로 실행이 허용되어 있는지 확인합니다.

인증키는 워크플로 입력값이나 코드에 넣지 않습니다. 워크플로는 오직
`secrets.DART_API_KEY`를 `DART_API_KEY` 환경 변수로 전달합니다.

## 2. 실행 방법

1. GitHub 저장소의 **Actions** 탭을 엽니다.
2. 왼쪽에서 **DART 재무제표 Excel 추출**을 선택합니다.
3. **Run workflow**를 누릅니다.
4. 다음 입력값을 지정하고 실행합니다.

| 입력값 | 설명 |
| --- | --- |
| `disclosure` | DART 공시 URL 또는 14자리 접수번호 |
| `statement_scope` | `auto`, `consolidated`, `separate` 중 하나 |
| `output_mode` | `single` 또는 `separate` |

입력 예:

```text
https://dart.fss.or.kr/dsaf001/main.do?rcpNo=20240319000709
```

또는:

```text
20240319000709
```

작업이 끝나면 실행 상세 화면 아래의 **Artifacts**에서
`dart-financial-statements-실행번호` 파일을 내려받아 압축을 풉니다.

## 재무제표 범위

- `auto`(기본값): 연결 재무제표 3종을 우선 추출합니다. 연결 3종 중 하나라도
  없으면 별도 재무제표 3종 전체로 다시 시도합니다.
- `consolidated`: 연결 재무제표 3종만 시도합니다.
- `separate`: 별도 재무제표 3종만 시도합니다.

`auto`는 연결/별도 자료를 한 결과 안에서 섞지 않습니다.

## XBRL이 없는 PDF 공시

XBRL 다운로드가 불가능하면 자동으로 다음 순서가 실행됩니다.

1. 입력 URL의 `dcmNo`를 사용합니다. URL에 `dcmNo`가 없거나 접수번호만 입력한
   경우 DART 공시 페이지에서 문서번호를 찾습니다.
2. DART의 전체 보고서 PDF를 다운로드합니다.
3. PDF 안에서 재무상태표, 손익계산서 또는 포괄손익계산서, 현금흐름표 제목을
   찾습니다.
4. 세 표를 모두 찾고 표 구조를 읽을 수 있을 때만 Excel을 만듭니다.

PDF 대체 추출은 글자를 마우스로 선택할 수 있는 텍스트형 PDF를 대상으로 합니다.
종이를 촬영하거나 스캔한 이미지 PDF, 암호화된 PDF, 표의 열 경계가 무너진 PDF는
자동 추출하지 못할 수 있습니다. 이 경우 잘못된 숫자를 임의로 만들지 않고 오류로
중단합니다. PDF에서도 `auto`는 연결 3종을 먼저 찾고, 완전한 연결 세트가 없으면
별도 3종 전체로 다시 시도합니다.

## 출력 방식

### `single`

Excel 파일 하나를 만들며 시트는 정확히 아래 3개입니다.

1. `재무상태표`
2. `손익계산서`
3. `현금흐름표`

파일명은 DART 공시 페이지의 회사명, 보고서 종류, 공시 연월을 사용합니다.

```text
펀진_사업보고서_2026.03.xlsx
```

### `separate`

재무제표별 Excel 파일을 하나씩, 총 3개 만듭니다.

```text
펀진_사업보고서_2026.03_재무상태표.xlsx
펀진_사업보고서_2026.03_손익계산서.xlsx
펀진_사업보고서_2026.03_현금흐름표.xlsx
```

보고서 종류에는 `사업보고서`, `반기보고서`, `분기보고서` 중 공시에 표시된
종류가 들어갑니다.

모든 출력은 다단계 pandas 열 이름을 읽기 쉬운 한 줄 머리글로 바꿉니다. 숫자
셀은 문자열이 아닌 Excel 숫자 값으로 유지하며, 굵은 머리글, 첫 행 고정,
천 단위 표시 및 열 너비 조정을 적용합니다.

각 시트에는 보고서 표처럼 다음 열만 남습니다.

```text
항목 | 2025-12-31 | 2024-12-31
```

손익계산서나 현금흐름표처럼 기간 누적값인 경우에는 다음처럼 기간 범위를
표시합니다.

```text
항목 | 2025-01-01 ~ 2025-12-31 | 2024-01-01 ~ 2024-12-31
```

XBRL의 `concept_id`, 영문 계정명, `class` 분류 열과 PDF의 주석 열은 Excel에
포함하지 않습니다. 연결/별도 범위는 선택한 재무제표 세트 자체에 적용되므로
각 행마다 같은 분류를 반복해서 표시하지 않습니다.

## 오류 안내

- `DART_API_KEY가 없습니다`: 저장소 Actions secret의 이름과 값을 확인하세요.
- `PDF 문서번호(dcmNo)를 찾을 수 없습니다`: 가능하면 주소에 `dcmNo`가 포함된
  DART 공시 URL을 사용하세요.
- `PDF에서 필수 재무제표 3종을 모두 추출하지 못했습니다`: PDF에 글자 선택이
  가능한지, 세 재무제표가 모두 있는지 확인하세요.
- `스캔 이미지 PDF`: 현재 OCR을 사용하지 않으므로 이미지로만 된 표는 추출할 수
  없습니다.
- `필수 재무제표를 모두 찾을 수 없습니다`: 선택한 범위에 재무상태표,
  손익계산서(또는 포괄손익계산서), 현금흐름표 중 하나 이상이 없습니다.
- URL 오류: `dart.fss.or.kr` 공식 URL의 `rcpNo` 값 또는 14자리 접수번호를
  확인하세요.

이 저장소는 자본변동표나 다른 주석/명세서를 추출하지 않습니다.

## 개발 및 테스트

개발 환경에 Python과 의존성이 있는 경우 아래 명령으로 검증할 수 있습니다.
일반 사용자는 이 절차가 필요하지 않습니다.

```bash
python -m pip install -r requirements.txt
python -m unittest discover -s tests -v
python -m py_compile dart_link.py extract_financials.py
python -m py_compile pdf_financials.py
```
