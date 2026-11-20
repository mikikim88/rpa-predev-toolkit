---
name: feasibility
description: 발굴된 RPA 후보 업무들의 타당성을 평가하고(5개 항목 25점 척도, 등급, 권장 자동화 범위) ROI(절감시간, FTE, 절감액)와 우선순위를 산정해 Excel 과제 평가표로 만든다. 사용자가 타당성 검토, 자동화 우선순위, ROI 계산, 과제 선정, 어떤 업무부터 자동화할지, 평가표 작성을 요청할 때 사용.
---

# RPA 타당성 평가 · ROI · 우선순위

과제 발굴(`/rpa-predev:discovery`) 결과인 `assessment.json`에 **점수와 근거, 예상 자동화율**을 채워
`RPA_과제평가표.xlsx`를 갱신한다. 평가표의 총점·등급·ROI·우선순위는 Excel 수식이라 현업이 점수를 고쳐도 자동 반영된다.

스크립트 위치: `${CLAUDE_PLUGIN_ROOT}/scripts`
스키마: `references/assessment-schema.md` · 예시: `${CLAUDE_PLUGIN_ROOT}/examples/assessment.json`

## 작업 절차

### 1. 입력 확보

- `assessment.json`이 있으면 그대로 사용. 없고 평가표(xlsx)만 있으면 openpyxl로 `1_과제목록`, `2_타당성평가`,
  `3_ROI`의 입력 셀(노란 셀)을 읽어 JSON을 재구성한다.
- 후보가 하나도 없으면 먼저 과제 발굴을 하도록 안내한다.

### 2. 5개 항목 채점 (각 1–5점)

| 항목 | 5점 | 3점 | 1점 |
|------|-----|-----|-----|
| 업무 표준화 `standardization` | 절차가 명확히 정형화 | 일부 변동 | 매번 다름 |
| 데이터 정형성 `data_structure` | 정형 100% | 일부 비정형 | 비정형 위주 |
| 시스템 안정성 `system_stability` | 변경 적음 | 가끔 변경 | 자주 변경 / 차세대 예정 |
| 처리 빈도/볼륨 `volume` | 고빈도·고볼륨 | 중간 | 저빈도·소량 |
| 판단 개입 `judgment` | 거의 없음 | 룰 기반 판단 | 사람 판단 필수 |

채점 원칙:
- **모든 점수에 근거(`reason`)를 한 줄로** 적는다. 근거는 자료·인터뷰에서 확인된 사실이어야 한다.
- 판단할 정보가 없으면 점수를 **비워 두고** `open_questions`에 질문을 추가한다 (평가표에 "평가 미완료"로 표시됨).
  추정으로 채운 경우엔 reason에 "(추정)"을 명시한다.
- 볼륨 기준 참고: 월 40시간↑ 또는 월 500건↑ = 5, 월 10–40시간 = 3–4, 월 5시간 미만 = 1–2.

### 3. 자동화율 · 구현 난이도

- `expected_automation_rate`: 완전 자동화 0.8–0.95 / 부분 자동화 0.4–0.7 / RDA 0.2–0.4.
  예외 처리·사람 승인에 남는 시간을 빼고 잡는다.
- `complexity` (하/중/상): 웹·Excel 위주 = 하, 데스크톱 앱·SAP GUI·시스템 3개 이상 = 중,
  Citrix/VDI·레거시 단말·공동인증서·OCR 필요 = 상.

### 4. 평가표 생성

```bash
python "${CLAUDE_PLUGIN_ROOT}/scripts/build_assessment.py" <작업폴더>/assessment.json
# → <작업폴더>/RPA_과제평가표.xlsx
```

`meta.hourly_cost`(시간당 인건비)를 사용자가 알려주면 넣는다. 모르면 생략한다 (절감액 열만 빈칸).

### 5. 결과 보고

채팅에는 다음만 간결하게 정리한다:
- 우선순위 상위 과제 표 (업무명, 총점/등급, 월 절감시간, 우선순위 구분)
- 평가 미완료 · 정보 부족 과제와 필요한 확인 사항
- 다음 단계: 1순위 과제는 `/rpa-predev:process-analysis`로 PDD 작성 (현업 업무 설명 자료 요청)

종합점수 공식은 적합도 60% + 효과 40%이며, 회사 기준이 다르면 사용자에게 확인 후
`build_assessment.py`의 `sheet_priority` 수식을 조정한다.
