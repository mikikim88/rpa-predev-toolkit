---
name: process-analysis
description: 현업 담당자가 작성한 업무 설명 자료(PPT, 인터뷰 메모, 매뉴얼)를 바탕으로 RPA 프로세스 분석서(PDD, Process Design Document)를 PPT로 작성한다. 사용자가 PDD, 프로세스 분석서, 업무 분석, As-Is 분석을 요청하거나, 현업이 작성한 업무 설명 .pptx를 첨부하고 "자동화 가능한지 봐줘", "프로세스 정리해줘", "RPA 가능성 검토"라고 할 때 사용. PPT 첨부 + 업무 분석 요청이면 RPA를 명시하지 않아도 사용.
---

# RPA 프로세스 분석서(PDD) 작성

현업이 쓴 "사람이 하는 절차"를 **봇이 따라 할 수 있는 절차**로 재해석해 표준 PDD를 만든다.
자료를 그대로 옮기지 않는다. **누락 정보를 찾아 "확인 필요"로 드러내는 것**이 이 문서의 핵심 가치다.

스크립트 위치: `${CLAUDE_PLUGIN_ROOT}/scripts`
필요 패키지: `python -c "import pptx, openpyxl"` 실패 시 `pip install -r "${CLAUDE_PLUGIN_ROOT}/requirements.txt"`

## 작업 절차

### 1단계: 입력 자료 읽기

```bash
python "${CLAUDE_PLUGIN_ROOT}/scripts/extract_pptx.py" <업무설명.pptx> --images <작업폴더>/images
python "${CLAUDE_PLUGIN_ROOT}/scripts/preview_pptx.py" <업무설명.pptx> -o <작업폴더>/preview
```

- 텍스트와 **스크린샷 이미지를 모두** 확인한다 (Read 도구로 PNG 열람). 현업 PPT는 캡처만 붙이고 설명이 없는 경우가 많다.
- 화면 제목, 메뉴 트리, URL/트랜잭션 코드, 필드 라벨, 버튼명, 팝업 메시지를 이미지에서 읽어낸다.
- 자료 유형별 주의사항은 `references/ppt-extraction-tips.md` 참고.
- 같은 업무가 과제 평가표(`assessment.json`)에 있으면 거기 적힌 빈도·건수·점수를 출발점으로 쓴다.

### 2단계: 7가지 분석 항목 도출

| 항목 | 도출 기준 |
|------|----------|
| 1. 업무 개요/목적 | 목적, 수행 부서, 빈도, 트리거, 건당 시간, 월 건수, 결과물 |
| 2. 사용 시스템 | 시스템명, 접속 방식, 인증, 메뉴 경로, 자동화 난이도(하/중/상) |
| 3. 작업 단계 | 봇의 단일 행위 단위로 분해 (화면 전환·대기도 분리) |
| 4. 입력/출력 데이터 | 출처·형식·구조·건수·검증 규칙 / 저장 위치·후속 처리 |
| 5. 예외/분기 | 비즈니스 분기, 시스템 예외, 데이터 예외 |
| 6. 자동화 가능성 | 5개 항목 × 5점, 자동화 구간 / 수동 구간 |
| 7. 리스크/제약 | 기술, 보안·규제, 운영, 조직 |

항목별 상세 기준과 체크리스트는 `references/analysis-criteria.md`.

### 3단계: 자동화 관점 재해석

- **암묵 절차 → 명시 절차**: "확인 후 처리" → "A 필드가 X면 승인 버튼, Y면 반려 버튼 클릭"
- **덩어리 절차 분해**: "거래처 조회" → 메뉴 진입 → 조건 입력 → 조회 → 결과 더블클릭
- **부적합 신호 표시**: 사람 판단, 비정형 입력, 시스템 외부 작업(전화·대면)은 `human_judgment: true`로 두고
  부분 자동화 / RDA 범위를 제안
- **누락 정보**: 건수·시간, 예외 빈도, 대기 기준, 검증 규칙이 없으면 값을 비우고 `open_questions`에 질문으로 추가.
  **추측해서 채우지 않는다.**

### 4단계: pdd.json 작성 → PPT 생성

스키마는 `references/pdd-schema.md`, 완성 예시는 `${CLAUDE_PLUGIN_ROOT}/examples/pdd.json`.

```bash
python "${CLAUDE_PLUGIN_ROOT}/scripts/render_pdd.py" <작업폴더>/pdd.json
# → <작업폴더>/RPA분석서_<업무명>_v1.0.pptx
```

pdd.json은 산출물과 함께 보관한다 (인터뷰 후 v2.0 보완, SDD 작성의 입력이 된다).

### 5단계: 시각 QA

```bash
python "${CLAUDE_PLUGIN_ROOT}/scripts/preview_pptx.py" <작업폴더>/RPA분석서_<업무명>_v1.0.pptx
```

슬라이드 이미지를 확인해 글자 넘침, 겹침, 비어 보이는 칸, 잘못 강조된 "확인 필요"를 점검하고
JSON을 고쳐 다시 렌더링한다. 미리보기 도구가 없으면 `extract_pptx.py`로 텍스트만 점검하고 그 사실을 알린다.

## 사용자와의 대화

- 시작 전: 업무명이 불명확하거나 자료 외에 알고 있는 사실(권한, 처리 빈도 등)이 있는지 짧게 확인한다.
- 완료 후: 등급·예상 절감시간·권장 범위를 요약하고, **"현업 인터뷰 필요 항목"을 목록으로 안내**한다.
  인터뷰 답변을 받으면 pdd.json을 갱신해 `version: "v2.0"`으로 다시 생성한다.
- 다음 단계: 등급이 상/중이고 확인 필요 항목이 해소되면 `/rpa-predev:solution-design`으로 SDD를 작성할 수 있다고 안내한다.
