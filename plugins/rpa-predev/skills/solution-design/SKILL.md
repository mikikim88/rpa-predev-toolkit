---
name: solution-design
description: RPA 프로세스 분석서(PDD)를 바탕으로 솔루션 설계서(SDD, Solution Design Document)를 PPT로 작성한다. To-Be 프로세스, 봇 아키텍처(모듈/Dispatcher-Performer), 설정·큐 설계, 예외 처리, 로그·알림, 실행 환경, 보안, 테스트 시나리오, 개발 계획을 정의. 사용자가 SDD, 설계서, 봇 설계, To-Be 설계, 개발 착수 전 설계를 요청할 때 사용.
---

# RPA 솔루션 설계서(SDD) 작성

PDD가 "무엇을 자동화하는가"라면 SDD는 "봇을 어떻게 만들고 운영하는가"를 정의한다.
개발자가 이 문서만 보고 개발에 착수할 수 있어야 한다.

스크립트 위치: `${CLAUDE_PLUGIN_ROOT}/scripts`
스키마: `references/sdd-schema.md` · 설계 원칙: `references/design-guidelines.md` ·
예시: `${CLAUDE_PLUGIN_ROOT}/examples/sdd.json`

## 작업 절차

### 1. 입력 확인

- PDD의 `pdd.json`(또는 PDD PPT → `extract_pptx.py`)을 읽는다. PDD가 없으면 `/rpa-predev:process-analysis`를 먼저 안내한다.
- PDD의 **"현업 인터뷰 필요 항목"이 남아 있는지** 확인한다. 설계에 영향을 주는 항목(예외 처리 기준, 계정·권한,
  시스템 제약)이 미해결이면 사용자에게 알리고, 진행하더라도 SDD의 `open_questions`로 넘긴다.
- 확인할 것 (모르면 사용자에게 질문): **RPA 플랫폼**, 무인/유인 실행, 실행 서버·스케줄, 결과 알림 수신자.

### 2. 설계

`references/design-guidelines.md`를 따른다. 핵심:

1. **To-Be 흐름**: PDD Step을 Bot / 사람 / 시스템 주체로 재구성. PDD에서 `human_judgment`인 단계는 사람 또는 승인 연동으로.
2. **아키텍처**: 건수·재처리 필요성으로 Dispatcher–Performer vs 단일 프로세스 결정, 모듈별 역할·입출력 정의.
3. **설정 분리**: 경로·수신자·임계값은 Config/Asset, 계정은 Credential. 실제 비밀번호는 절대 기재하지 않는다.
4. **예외 처리**: PDD 예외를 Business / System으로 분류해 전부 옮기고, 감지 방법·재시도·알림을 구체화. 중복 처리 방지 방안 포함.
5. **로그·알림·리포트**, **실행 환경**, **보안**(최소 권한, 민감정보 마스킹, 감사 추적).
6. **테스트 시나리오**: 정상 + PDD 예외별 1건 이상 + 재시도/중복 방지 검증.
7. **개발 계획**과 **확인 필요 사항**.

### 3. sdd.json 작성 → PPT 생성

```bash
python "${CLAUDE_PLUGIN_ROOT}/scripts/render_sdd.py" <작업폴더>/sdd.json
# → <작업폴더>/RPA설계서_<업무명>_v1.0.pptx
```

### 4. 시각 QA

```bash
python "${CLAUDE_PLUGIN_ROOT}/scripts/preview_pptx.py" <작업폴더>/RPA설계서_<업무명>_v1.0.pptx
```

슬라이드 이미지를 확인해 넘침·겹침을 점검하고 JSON을 수정해 다시 렌더링한다.

### 5. 결과 보고

- 구조(패턴·모듈 수), 예외 처리 건수, 테스트 시나리오 수, 예상 개발 기간 요약
- 개발 착수 전 확정해야 할 확인 사항 (IT·보안·현업별)
- 이 단계가 개발 전 마지막 산출물임을 알리고, 개발 착수 체크리스트(계정·권한, 개발/운영 환경, 테스트 데이터)를 안내
