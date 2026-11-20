# sdd.json 스키마

`render_sdd.py`가 읽는 형식입니다. 전체 예시: `${플러그인}/examples/sdd.json`.
값이 없으면 비워 두면 빨간색 "확인 필요"로 표시됩니다. `queue`가 없으면 큐 설계 슬라이드는 생략됩니다.

```jsonc
{
  "meta": {
    "task_name": "업무명 (PDD와 동일)",
    "author": "", "date": "2026-11-20", "version": "v1.0",
    "rpa_platform": "UiPath (REFramework) | Power Automate Desktop | Automation Anywhere | Brity RPA | 미정",
    "pdd_ref": "RPA분석서_<업무명>_v1.0"
  },
  "overview": {
    "goal": "자동화 목표 한 문장",
    "automation_type": "Unattended | Attended(RDA) | 혼합",
    "expected_effect": "월 N시간 절감 (N FTE) 등",
    "scope_in": ["자동화 범위"], "scope_out": ["제외 범위"]
  },
  "tobe_flow": [{"no": 1, "actor": "Bot|사람|시스템", "action": "", "note": ""}],
  "architecture": {
    "pattern": "Dispatcher – Performer | 단일 프로세스",
    "description": "구조 설명 한 줄",
    "modules": [
      {"id": "M01", "name": "Init", "role": "", "input": "", "output": "", "systems": "",
       "steps": ["모듈 상세 처리 단계 (선택, 있으면 '모듈 상세' 슬라이드 생성)"]}
    ]
  },
  "config": [{"key": "", "type": "Setting|Asset|Credential", "value_example": "", "description": ""}],
  "queue": {                                   // 선택
    "name": "", "unit": "트랜잭션 1건의 단위", "reference": "고유 참조키 규칙",
    "retry": "", "priority": "",
    "item_fields": [{"name": "", "type": "", "description": ""}]
  },
  "exception_policy": {"business": "업무 예외 공통 정책", "system": "시스템 예외 공통 정책"},
  "exceptions": [{"type": "Business|System", "name": "", "detection": "", "handling": "", "retry": "", "notify": ""}],
  "logging": {
    "policy": "", "report": "처리 결과 리포트 형태/위치",
    "items": [{"point": "기록 시점", "content": "기록 내용"}]
  },
  "notifications": [{"trigger": "", "recipient": "", "channel": "", "content": ""}],
  "environment": {
    "machine": "", "schedule": "", "runtime": "", "licenses": "",
    "credential_store": "", "network": "", "orchestrator": "", "dependencies": ""
  },
  "security": [{"item": "", "measure": ""}],
  "test_cases": [{"id": "TC-01", "scenario": "", "type": "정상|예외", "data": "", "expected": ""}],
  "plan": [{"phase": "", "task": "", "duration": "", "owner": ""}],
  "open_questions": [{"area": "현업|IT|보안", "question": ""}]
}
```

Credential 유형 설정값은 슬라이드에 값이 출력되지 않고 "(Credential Store에 보관)"으로 표시됩니다.
실제 비밀번호·키는 JSON에 절대 넣지 않습니다.
