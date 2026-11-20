---
name: discovery
description: RPA 자동화 후보 업무를 발굴한다. 현업 배포용 업무 조사 양식(Excel)을 만들고, 회수한 조사서·인터뷰 메모·업무 목록·회의록에서 후보 업무를 추출·정리하고 1차 스크리닝해 과제 목록(assessment.json + Excel 평가표)을 만든다. 사용자가 과제 발굴, 자동화 대상 찾기, 업무 조사, 현업 설문/조사 양식, 인터뷰 정리, 후보 업무 목록을 요청할 때 사용.
---

# RPA 과제 발굴

흩어진 자료에서 자동화 후보 업무를 찾아 **같은 기준의 과제 목록**으로 만든다.
이 단계의 목표는 빠짐없이 모으고, 규모(건수·시간)를 수치로 확보하고, 부족한 정보를 드러내는 것이다.
점수를 매기는 일은 다음 단계(`/rpa-predev:feasibility`)에서 한다.

스크립트 위치: `${CLAUDE_PLUGIN_ROOT}/scripts`
필요 패키지: `python -c "import openpyxl"` 실패 시 `pip install -r "${CLAUDE_PLUGIN_ROOT}/requirements.txt"`

## 작업 절차

### A. 조사 양식이 필요할 때 (발굴 시작 전)

```bash
python "${CLAUDE_PLUGIN_ROOT}/scripts/build_assessment.py" --intake-template -o <작업폴더>/RPA_현업작성양식.xlsx
```

드롭다운(빈도, 입력 형태, 판단 개입)과 예시 행이 포함된 양식이다. 현업 배포용으로 전달하고,
인터뷰를 할 경우 `references/interview-guide.md`의 질문 목록을 함께 안내한다.

### B. 자료 수집 → 후보 목록

1. **작성된 조사 양식** (여러 파일 가능):
   ```bash
   python "${CLAUDE_PLUGIN_ROOT}/scripts/read_intake.py" 조사서1.xlsx 조사서2.xlsx -o <작업폴더>/assessment.json
   # 기존 목록에 추가할 때: --append
   ```
2. **비정형 자료** (인터뷰 메모, 회의록, 업무 분장표, 메일, PPT): 직접 읽고 후보 업무를 추출해
   `assessment.json`의 `candidates`에 추가한다. PPT는 `extract_pptx.py`로 텍스트를 뽑는다.
3. **정리 규칙** (`references/interview-guide.md`의 "업무 단위 정하기"):
   - 1 과제 = 1 트리거 → 1 결과물. 중복·유사 업무는 합치고, 합친 사실을 `notes`에 남긴다.
   - 업무명은 "대상 + 행위 + 시스템" 명사구로 통일하고 ID는 `T-001`부터 부여한다.
   - `frequency`, `monthly_volume`, `minutes_per_case`는 **숫자로** 적는다. "월 수백 건"처럼 모호하면 비우고
     `open_questions`에 "월 처리건수 확인"을 추가한다. 자료에 없는 수치를 추측하지 않는다.
   - `input_type`, `human_judgment`는 스키마의 3개 선택지 중 하나로 표기한다.
   - `source`에 출처(파일명·행, 인터뷰 날짜)를 남긴다.

스키마: `${CLAUDE_PLUGIN_ROOT}/skills/feasibility/references/assessment-schema.md`

### C. 평가표 생성 (1차 스크리닝 포함)

```bash
python "${CLAUDE_PLUGIN_ROOT}/scripts/build_assessment.py" <작업폴더>/assessment.json
```

`1_과제목록` 시트에 스크리닝 결과(적합 후보 / 검토 필요 / 부적합 가능 / 정보 부족)가 규칙에 따라 표시되고,
`5_확인필요` 시트에 확인 사항이 모인다. 점수가 없으므로 평가·우선순위 시트는 "평가 미완료"로 나오는 것이 정상이다.

### D. 결과 보고

- 후보 수, 스크리닝 결과별 건수, 월 소요시간 합계
- 정보 부족 과제와 현업에 다시 물어볼 질문
- 다음 단계: `/rpa-predev:feasibility`로 타당성 평가 및 우선순위 산정
