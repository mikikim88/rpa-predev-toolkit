"""Shared RPA feasibility scoring rules (used by the Excel and PPT renderers).

5 criteria x 1-5 points = 25 points.
  상 20-25 : 완전 자동화 권장
  중 12-19 : 부분 자동화 / RDA 검토
  하  5-11 : 자동화 부적합, 프로세스 개선 우선
"""
from __future__ import annotations

from typing import Any

# (json key, 한글 라벨, 5점 / 3점 / 1점 기준)
CRITERIA = [
    ("standardization", "업무 표준화", ("절차가 명확히 정형화됨", "일부 변동 있음", "매번 다름")),
    ("data_structure", "데이터 정형성", ("정형 데이터 100%", "일부 비정형", "비정형 위주")),
    ("system_stability", "시스템 안정성", ("안정적, 변경 적음", "가끔 변경", "자주 변경 / 차세대 예정")),
    ("volume", "처리 빈도/볼륨", ("고빈도·고볼륨", "중간", "저빈도·소량")),
    ("judgment", "판단 개입", ("거의 없음", "일부 룰 기반 판단", "사람 판단 필수")),
]

SCOPES = {
    "상": "완전 자동화 (STP)",
    "중": "부분 자동화 / RDA",
    "하": "자동화 보류 (프로세스 개선 선행)",
}


def _score(v: Any) -> float | None:
    if isinstance(v, dict):
        v = v.get("score")
    try:
        f = float(v)
    except (TypeError, ValueError):
        return None
    return f if 1 <= f <= 5 else None


def total_score(scores: dict) -> int | None:
    """Sum of the 5 criteria, or None if any criterion is missing."""
    vals = [_score((scores or {}).get(k)) for k, _, _ in CRITERIA]
    if any(v is None for v in vals):
        return None
    return int(round(sum(vals)))


def grade_of(total: int | None) -> str:
    if total is None:
        return "확인 필요"
    return "상" if total >= 20 else "중" if total >= 12 else "하"


def recommended_scope(grade: str) -> str:
    return SCOPES.get(grade, "확인 필요")


def screen(c: dict) -> tuple[str, str]:
    """1차 스크리닝 (과제 발굴 단계). Returns (결과, 사유)."""
    if c.get("screening"):
        s = c["screening"]
        return s.get("result", ""), s.get("reason", "")
    reasons = []
    judg = str(c.get("human_judgment", ""))
    inp = str(c.get("input_type", ""))
    try:
        hours = float(c.get("monthly_volume")) * float(c.get("minutes_per_case")) / 60
    except (TypeError, ValueError):
        hours = None
    if "필수" in judg:
        return "부적합 가능", "사람 판단 필수 단계 존재"
    if "비정형" in inp:
        reasons.append("비정형 입력 (IDP 결합 검토)")
    if hours is None:
        return "정보 부족", "월 처리건수 / 건당 처리시간 확인 필요"
    if hours < 5:
        reasons.append(f"효과 작음 (월 {hours:.1f}시간)")
    if reasons:
        return "검토 필요", "; ".join(reasons)
    return "적합 후보", f"월 {hours:.1f}시간 소요, 정형 업무"
