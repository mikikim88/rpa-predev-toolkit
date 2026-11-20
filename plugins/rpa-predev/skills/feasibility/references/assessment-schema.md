# assessment.json 스키마

`build_assessment.py`가 읽는 형식입니다. 과제 발굴(discovery)과 타당성 평가(feasibility)가
**같은 파일을 이어서** 채웁니다. 전체 예시: `${플러그인}/examples/assessment.json`.

```jsonc
{
  "meta": {
    "project": "재무본부 RPA 과제 발굴",
    "organization": "대상 조직",
    "date": "2026-11-20",
    "author": "작성자",
    "work_hours_per_month": 160,     // 1 FTE 월 근무시간 (ROI 시트 파라미터)
    "hourly_cost": 35000,            // 시간당 인건비(원). 없으면 생략 → 절감액 미산출
    "effect_threshold_hours": 20     // 우선순위 '효과 높음' 기준 (월 절감시간)
  },
  "candidates": [
    {
      // ---- 과제 발굴 단계에서 채움 ----
      "id": "T-001",
      "task_name": "업무명 (한 줄 명사구)",
      "department": "부서",
      "owner": "담당자/역할",
      "description": "무엇을, 왜, 결과물",
      "frequency": "일|주|월|분기|반기|연|수시",
      "monthly_volume": 450,           // 숫자. 모르면 생략
      "minutes_per_case": 6,           // 숫자(분). 모르면 생략
      "systems": "SAP, Excel, Outlook",
      "input_type": "정형 (Excel/CSV/DB) | 반정형 (양식 PDF/메일) | 비정형 (자유 텍스트/이미지)",
      "human_judgment": "없음 | 일부 (룰 기반) | 필수 (사람 판단)",
      "source": "출처 (인터뷰 날짜, 조사서 파일명)",
      "screening": {"result": "", "reason": ""},   // 선택: 생략하면 규칙으로 자동 판정
      "open_questions": [{"area": "데이터", "question": ""}],   // 또는 문자열 배열

      // ---- 타당성 평가 단계에서 채움 ----
      "scores": {                      // 각 1~5. 근거 없으면 생략 (→ 평가 미완료)
        "standardization":  {"score": 5, "reason": ""},
        "data_structure":   {"score": 4, "reason": ""},
        "system_stability": {"score": 4, "reason": ""},
        "volume":           {"score": 4, "reason": ""},
        "judgment":         {"score": 4, "reason": ""}
      },
      "expected_automation_rate": 0.85,   // 0~1
      "complexity": "하|중|상",            // 구현 난이도
      "scope": "",                        // 선택: 비우면 등급으로 자동
      "notes": ""
    }
  ]
}
```

## 자동 계산 (Excel 수식)

| 값 | 계산 |
|----|------|
| 월 소요시간 | 월 처리건수 × 건당 처리시간 ÷ 60 |
| 총점 / 등급 | 5개 점수 합 (5개 모두 있어야 계산) / 상 20↑, 중 12↑, 하 |
| 월 절감시간 | 월 소요시간 × 예상 자동화율 |
| FTE / 연간 절감액 | 월 절감시간 ÷ 월 근무시간 / 연간 절감시간 × 시간당 인건비 |
| 종합점수(100) | 총점/25 × 60 + MIN(월 절감시간 ÷ (효과기준 × 2), 1) × 40 |
| 우선순위 | 1순위 = 총점 20↑ & 절감 ≥ 기준 / 2순위 = 둘 중 하나 / 3순위 / 보류 = 총점 12 미만 |

## 1차 스크리닝 규칙 (scoring.py `screen`)

1. 판단 개입 "필수" → **부적합 가능**
2. 월 처리건수·건당 시간 누락 → **정보 부족**
3. 비정형 입력 또는 월 5시간 미만 → **검토 필요**
4. 그 외 → **적합 후보**
