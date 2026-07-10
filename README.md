# NPL 투심자료 자동화

은행/매각자문사로부터 받은 NPL(부실채권) Data Disk를 읽어, 회사 표준 양식의 투자심의위원회(투심) 엑셀 보고서와 워드 보고서(자산 소개, 1~9페이지)를 자동으로 생성하는 도구입니다.

## 주요 기능

- Data Disk 엑셀 파일(복수 선택 가능)에서 차주시트/담보시트/신용보증서시트/회생시트를 자동 판별
- 은행·매각자문사별로 다른 컬럼명을 표준 스키마 컬럼에 자동 매칭 (사용자 확인/수정 가능)
- 표준 4개 시트(차주/담보/신용보증서/회생) + `투심` 집계 시트가 포함된 엑셀 산출물 생성
  - `투심` 시트는 정적 값이 아니라 실제 Excel 수식(COUNTIFS/SUMIFS 등)으로 작성되어, 원본 데이터가 수정되면 자동 재계산됩니다.
- 워드 보고서(총괄표, 금액별/담보물종류별/담보물지역별/경매절차/회생절차/신용보증서 분류, 자동 생성 나레이션 문장)를 원본 DRM 템플릿에 값만 채워서 생성
  - LSPA/평가결과/시뮬레이션/시장동향/투자금액 등 Data Disk만으로 산출 불가능한 10페이지 이후 내용은 건드리지 않습니다.
- DRM(IRM/DocumentSAFER 등)이 걸린 엑셀·워드 파일도 로컬에 설치된 Office를 통해 정상적으로 읽고 씁니다 (사용자 본인의 접근 권한 필요).

## 실행 방법

### 1) exe로 바로 실행 (권장)

[Releases](../../releases) 탭에서 최신 버전의 `NPL투심자료자동화.exe`와 `classification_config.json`을 **같은 폴더**에 다운로드한 뒤, exe를 더블클릭하면 됩니다.

**요구 사항**
- Windows + MS Office(Excel, Word) 설치
- 처리하려는 Data Disk/템플릿 파일에 대한 DRM 접근 권한으로 로그인되어 있어야 함

### 2) 소스코드로 실행

```bash
pip install -r requirements.txt
python main.py
```

## 사용 순서 (GUI)

1. **파일 선택**: Data Disk 엑셀 파일 선택 (여러 개 선택 가능)
2. **시트 매핑 확인**: 각 시트가 차주/담보/신용보증서/회생 중 무엇으로 자동 판별됐는지 확인·수정
3. **컬럼 매핑 확인**: 원본 컬럼명이 표준 컬럼에 올바르게 매칭됐는지 확인·수정
4. **보고서 정보 입력**: 매각은행명, Program명, 보고서 날짜 입력 + (선택) 워드 보고서 템플릿(DRM 원본) 지정
5. **생성**: 저장 경로 지정 후 엑셀(및 워드) 산출물 생성

## 프로젝트 구조

| 파일 | 역할 |
|---|---|
| `main.py` | 진입점 |
| `gui.py` | tkinter 기반 마법사 UI |
| `excel_io.py` | DRM 대응 엑셀 리더 (openpyxl → 실패 시 win32com 폴백) |
| `schema.py` | 표준 시트/컬럼 스키마 정의 |
| `sheet_matcher.py` / `column_matcher.py` | 원본 시트·컬럼 → 표준 스키마 자동 매칭 |
| `normalizer.py` | 매핑 결과를 표준 컬럼 기준 DataFrame으로 정규화 |
| `classify.py` / `classification_config.json` | 금액구간/담보유형/지역 분류 기준 (편집 가능) |
| `aggregator.py` | 투심 집계표 8종 계산 |
| `writer.py` | 엑셀 산출물 작성 (표준 시트 + 투심 시트, 실시간 수식) |
| `narrative.py` | 워드 보고서용 나레이션 문장 자동 생성 |
| `word_writer.py` | DRM 워드 템플릿을 열어 표/문장을 채워 넣는 COM 자동화 |
| `regression_test.py` | 실제 Data Disk 기준 회귀 테스트 |

## 분류 기준 수정

`classification_config.json`에서 금액 구간, 담보물 유형 그룹(유사유형), 지역 그룹 분류 기준을 코드 수정 없이 편집할 수 있습니다. exe로 실행하는 경우 exe와 같은 폴더의 `classification_config.json`을 읽습니다.

## 새 버전 빌드

```bash
pyinstaller "NPL투심자료자동화.spec" --clean -y
```

`dist/` 폴더에 exe가 생성되며, `classification_config.json`을 같은 폴더로 복사해야 합니다.
