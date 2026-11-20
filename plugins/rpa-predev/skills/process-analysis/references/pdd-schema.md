# pdd.json 스키마

`render_pdd.py`가 읽는 입력 형식입니다. 전체 예시는 `${플러그인}/examples/pdd.json`.

- 모르는 값은 **비워 두거나(null / "") 생략**합니다. 렌더러가 빨간색 "확인 필요"로 표시합니다.
  추측한 값을 채우지 마세요.
- 문자열 안의 "확인 필요"는 자동으로 빨간색으로 강조됩니다.
- 총점·등급·권장 범위·월 소요시간·절감시간은 **렌더러가 계산**합니다.

```jsonc
{
  "meta": {
    "task_name": "월말 거래처 정산 데이터 ERP 등록",   // 필수. 파일명에도 사용
    "department": "재무팀",
    "author": "작성자",
    "date": "2026-11-20",
    "version": "v1.0",
    "work_hours_per_month": 160                      // FTE 환산 기준 (기본 160)
  },
  "summary": {
    "key_points": ["핵심 분석 결과 3~5줄"]           // 분석 요약 슬라이드 하단
  },
  "overview": {                                      // 1. 업무 개요
    "purpose": "업무 목적",
    "owner": "수행 부서/담당 (인원)",
    "frequency": "월 1회 (매월 말일)",
    "trigger": "시작 조건 (시간/이벤트)",
    "summary_flow": "A → B → C (한 줄 흐름)",
    "time_per_case_min": 6,                          // 숫자 (분)
    "monthly_volume": 450,                           // 숫자 (건)
    "output": "결과물 / 수혜자"
  },
  "systems": [                                       // 2. 사용 시스템
    {"name": "SAP S/4HANA", "access": "클라이언트", "auth": "ID/PW",
     "menu_path": "회계 → 매입채무 → FB60", "difficulty": "하|중|상", "note": ""}
  ],
  "steps": [                                         // 3. 작업 단계 (봇 단일 행위 단위)
    {"no": 1, "action": "행위", "system": "시스템/화면", "operation": "입력/조작",
     "check": "결과 확인", "human_judgment": false}  // true → 빨간 ⚠ 표시
  ],
  "inputs":  [{"name": "", "source": "", "format": "", "structure": "", "volume": "", "validation": ""}],
  "outputs": [{"name": "", "destination": "", "format": "", "follow_up": "", "notify": ""}],
  "exceptions": [                                    // 5. 예외 / 분기
    {"category": "비즈니스|시스템|데이터", "type": "", "condition": "", "handling": "", "notify": ""}
  ],
  "assessment": {                                    // 6. 자동화 가능성 (각 1~5점)
    "scores": {
      "standardization":  {"score": 5, "reason": "근거"},
      "data_structure":   {"score": 4, "reason": ""},
      "system_stability": {"score": 4, "reason": ""},
      "volume":           {"score": 4, "reason": ""},
      "judgment":         {"score": 4, "reason": "개입이 적을수록 고득점"}
    },
    "automation_rate": 0.85,                         // 0~1, 절감시간 계산용
    "scope": "",                                     // 비우면 등급으로 자동 결정
    "automated_part": "자동화 구간 (Step 번호)",
    "manual_part": "수동 구간 (Step 번호)"
  },
  "risks": [                                         // 7. 리스크 (카테고리별 4건까지 카드 표시)
    {"category": "기술|보안·규제|운영|조직", "item": "", "impact": "상|중|하", "mitigation": ""}
  ],
  "open_questions": [                                // 현업 인터뷰 필요 항목 (핵심!)
    {"area": "업무개요|시스템|Step|데이터|예외|리스크", "question": ""}
  ],
  "recommendations": {"short": [""], "mid": [""], "long": [""]},
  "source_map": [{"source": "원본 Slide 2", "content": "", "used_in": "1. 업무 개요"}],   // 선택
  "glossary": [{"term": "PDD", "desc": ""}]          // 선택 (생략 시 기본 용어집)
}
```

## 생성되는 슬라이드 (12~16장)

1. 표지 · 2. 분석 요약 · 3. 업무 개요 · 4. 사용 시스템 · 5~. 작업 단계 (자동 페이지 분할) ·
입력 데이터 · 출력 데이터 · 예외 처리 · 자동화 가능성 평가 · 리스크 · ⚠ 현업 인터뷰 필요 항목 ·
권고안 · 부록 (원본 매핑, 용어 정의)

빈도·월 처리건수·건당 처리시간이 비어 있으면 인터뷰 항목에 자동으로 추가됩니다.
