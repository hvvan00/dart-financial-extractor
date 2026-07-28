# DART 재무제표 Excel 추출기

DART 공시 URL 또는 14자리 접수번호를 입력하면 공시의 XBRL에서 아래 재무제표
3종만 추출하여 Excel 파일로 내려받을 수 있는 GitHub Actions 저장소입니다.

- 재무상태표
- 손익계산서 또는 포괄손익계산서
- 현금흐름표

웹 서버, 데이터베이스, 데스크톱 프로그램은 없습니다. 사용자의 PC에 Python을
설치할 필요도 없습니다. GitHub Actions가 `dart-fss`로 XBRL을 읽고 결과 파일을
아티팩트로 제공합니다.

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

## 출력 방식

### `single`

Excel 파일 하나를 만들며 시트는 정확히 아래 3개입니다.

1. `재무상태표`
2. `손익계산서`
3. `현금흐름표`

### `separate`

재무제표별 Excel 파일을 하나씩, 총 3개 만듭니다.

모든 출력은 다단계 pandas 열 이름을 읽기 쉬운 한 줄 머리글로 바꿉니다. 숫자
셀은 문자열이 아닌 Excel 숫자 값으로 유지하며, 굵은 머리글, 첫 행 고정,
천 단위 표시 및 열 너비 조정을 적용합니다.

## 오류 안내

- `DART_API_KEY가 없습니다`: 저장소 Actions secret의 이름과 값을 확인하세요.
- `XBRL 파일을 찾을 수 없습니다`: 해당 공시가 정기보고서 XBRL을 제공하는지
  확인하세요. 모든 DART 공시가 재무제표 XBRL을 포함하지는 않습니다.
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
```
