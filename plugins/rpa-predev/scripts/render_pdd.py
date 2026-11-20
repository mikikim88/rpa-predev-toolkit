"""Render an RPA PDD (프로세스 분석서) PowerPoint from a JSON spec.

Usage:
    python render_pdd.py pdd.json [-o RPA분석서_<업무명>_v1.0.pptx]

JSON spec: see skills/rpa-process-analysis/references/pdd-schema.md
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pptx_kit import (C, CENTER, CONTENT_TOP, CONTENT_W, MARGIN_X, MIDDLE, TBD, Deck,  # noqa: E402
                      add_arrow_down, add_badge, add_card, add_chevrons, add_rect,
                      add_score_bars, add_table, add_text, blank, est_lines, fmt, grade_color,
                      impact_color, num)
from scoring import CRITERIA, grade_of, recommended_scope, total_score  # noqa: E402

DOC_LABEL = "RPA 프로세스 분석서"

DEFAULT_GLOSSARY = [
    ("PDD", "Process Design Document. 현행(As-Is) 업무를 자동화 관점으로 정리한 프로세스 분석서"),
    ("SDD", "Solution Design Document. 봇 구조·예외처리·운영환경을 정의한 솔루션 설계서"),
    ("STP", "Straight Through Processing. 사람 개입 없는 완전 자동 처리"),
    ("RDA", "Robotic Desktop Automation. 사람이 실행하고 봇이 보조하는 유인(Attended) 자동화"),
    ("IDP", "Intelligent Document Processing. OCR/AI로 비정형 문서에서 데이터를 추출하는 기술"),
    ("FTE", "Full Time Equivalent. 1인이 한 달 동안 일하는 시간(기본 160시간) 기준 인력 환산치"),
]


def monthly_hours(ov: dict) -> float | None:
    vol, mins = num(ov.get("monthly_volume")), num(ov.get("time_per_case_min"))
    if vol is None or mins is None:
        return None
    return vol * mins / 60


def build(data: dict, out: Path) -> Path:
    meta = data.get("meta", {})
    ov = data.get("overview", {})
    asm = data.get("assessment", {})
    task = meta.get("task_name") or TBD
    deck = Deck(DOC_LABEL, task)

    scores = asm.get("scores", {})
    total = total_score(scores)
    grade = asm.get("grade") or grade_of(total)
    scope = asm.get("scope") or recommended_scope(grade)
    m_hours = monthly_hours(ov)
    rate = num(asm.get("automation_rate"))
    saved = m_hours * rate if (m_hours is not None and rate is not None) else None
    work_h = num(meta.get("work_hours_per_month")) or 160

    # 1. cover ---------------------------------------------------------------
    deck.cover("RPA 프로세스 분석서", task, [
        ("작성일", meta.get("date")),
        ("분석 대상 부서", meta.get("department")),
        ("작성자", meta.get("author")),
        ("버전", meta.get("version") or "v1.0"),
    ], "PDD")

    # 2. executive summary ---------------------------------------------------
    s = deck.slide("분석 요약", "자동화 가능성, 예상 효과, 권장 범위")
    cw, ch, y = (CONTENT_W - 0.6) / 3, 2.7, CONTENT_TOP
    # card 1: grade
    x = MARGIN_X
    add_rect(s, x, y, cw, ch, C["light"])
    add_text(s, x + 0.25, y + 0.15, cw - 0.5, 0.4, "자동화 가능성", size=14, bold=True, color=C["muted"])
    add_text(s, x + 0.25, y + 0.6, 1.6, 1.4, grade if total is not None else "-", size=60, bold=True,
             color=grade_color(grade), anchor=MIDDLE)
    add_text(s, x + 1.9, y + 0.85, cw - 2.1, 0.5, f"{total} / 25점" if total is not None else TBD,
             size=20, bold=True)
    add_text(s, x + 1.9, y + 1.35, cw - 2.1, 0.5, "5개 항목 × 5점", size=11, color=C["muted"])
    add_text(s, x + 0.25, y + 2.05, cw - 0.5, 0.5, "상 20–25 · 중 12–19 · 하 5–11", size=10.5,
             color=C["muted"])
    # card 2: effect
    x += cw + 0.3
    add_rect(s, x, y, cw, ch, C["light"])
    add_text(s, x + 0.25, y + 0.15, cw - 0.5, 0.4, "예상 효과 (월)", size=14, bold=True, color=C["muted"])
    add_text(s, x + 0.25, y + 0.6, cw - 0.5, 0.9,
             f"{saved:,.1f} 시간 절감" if saved is not None else TBD, size=28, bold=True,
             color=C["success"] if saved is not None else C["highlight"], anchor=MIDDLE)
    add_text(s, x + 0.25, y + 1.5, cw - 0.5, 1.1, [
        f"월 처리 {fmt(ov.get('monthly_volume'))}건 × 건당 {fmt(ov.get('time_per_case_min'))}분"
        if m_hours is not None else f"월 처리건수 · 건당 처리시간 {TBD}",
        f"현재 월 {m_hours:,.1f}시간 소요" if m_hours is not None else f"현재 소요시간 {TBD}",
        f"자동화율 {rate:.0%} · 약 {saved / work_h:.2f} FTE" if saved is not None
        else f"자동화율 {fmt(asm.get('automation_rate'))}",
    ], size=11.5)
    # card 3: scope
    x += cw + 0.3
    add_rect(s, x, y, cw, ch, C["light"])
    add_text(s, x + 0.25, y + 0.15, cw - 0.5, 0.4, "권장 자동화 범위", size=14, bold=True, color=C["muted"])
    add_badge(s, x + 0.25, y + 0.7, cw - 0.5, 0.65, scope, C["primary"], size=18)
    add_text(s, x + 0.25, y + 1.5, cw - 0.5, 1.1, fmt(asm.get("automated_part"), ""), size=11.5)
    # key points
    pts = data.get("summary", {}).get("key_points") or []
    add_text(s, MARGIN_X, y + ch + 0.3, 3, 0.4, "핵심 분석 결과", size=14, bold=True, color=C["primary"])
    add_text(s, MARGIN_X, y + ch + 0.75, CONTENT_W, 1.9, pts or [TBD], size=12.5, bullets=True,
             line_spacing=1.15)

    # 3. overview -----------------------------------------------------------
    s = deck.slide("1. 업무 개요")
    rows = [
        ["업무명", task],
        ["업무 목적", fmt(ov.get("purpose"))],
        ["수행 부서 / 담당", fmt(ov.get("owner") or meta.get("department"))],
        ["수행 빈도", fmt(ov.get("frequency"))],
        ["시작 트리거", fmt(ov.get("trigger"))],
        ["건당 처리시간", f"{fmt(ov.get('time_per_case_min'))}분" if not blank(ov.get("time_per_case_min")) else TBD],
        ["월 처리건수", f"{fmt(ov.get('monthly_volume'))}건" if not blank(ov.get("monthly_volume")) else TBD],
        ["월 총 소요시간", f"{m_hours:,.1f}시간 (계산값)" if m_hours is not None else TBD],
        ["결과물 / 수혜자", fmt(ov.get("output"))],
    ]
    add_table(s, MARGIN_X, CONTENT_TOP, [2.0, 5.6], ["항목", "내용"], rows, size=11,
              bold_cols=[0], zebra=True)
    fx, fw = 8.9, 3.8
    flow = [("Trigger", ov.get("trigger")), ("수행", ov.get("summary_flow") or ov.get("purpose")),
            ("결과", ov.get("output"))]
    for i, (lab, txt) in enumerate(flow):
        yy = CONTENT_TOP + i * 1.85
        add_badge(s, fx, yy, fw, 0.42, lab, C["primary"] if i != 1 else C["success"], size=12)
        add_text(s, fx, yy + 0.42, fw, 0.95, fmt(txt), size=11, fill=C["light"], anchor=MIDDLE,
                 margin=0.12)
        if i < 2:
            add_arrow_down(s, fx + fw / 2 - 0.2, yy + 1.42, h=0.36)

    # 4. systems -----------------------------------------------------------
    systems = data.get("systems", [])
    sys_rows = [[fmt(x.get("name")), fmt(x.get("access")), fmt(x.get("auth")), fmt(x.get("menu_path")),
                 fmt(x.get("difficulty")), fmt(x.get("note"), "-")] for x in systems]

    def sys_legend(sl, _):
        lx = MARGIN_X + CONTENT_W - 5.3
        add_text(sl, lx, 1.02, 1.4, 0.32, "자동화 난이도:", size=10, color=C["muted"], anchor=MIDDLE)
        for i, (k, lab) in enumerate([("하", "하 (쉬움)"), ("중", "중"), ("상", "상 (어려움)")]):
            add_badge(sl, lx + 1.45 + i * 1.3, 1.04, 1.2, 0.28, lab, impact_color(k), size=9.5)

    deck.table_slides("2. 사용 프로그램 / 시스템", ["시스템명", "접속 방식", "인증 방식", "메뉴 경로", "난이도", "비고"],
                      sys_rows, [2.2, 1.9, 1.8, 3.4, 0.9, 1.9], size=10.5,
                      cell_fill_fn=lambda i: {4: impact_color(systems[i].get("difficulty"))} if systems else {},
                      align=[None, None, None, None, CENTER, None], bold_cols=[0],
                      before=sys_legend)

    # 5. steps ---------------------------------------------------------------
    steps = data.get("steps", [])
    st_rows, st_colors = [], []
    for i, x in enumerate(steps, start=1):
        judg = bool(x.get("human_judgment"))
        st_rows.append([fmt(x.get("no") or i), ("⚠ " if judg else "") + fmt(x.get("action")),
                        fmt(x.get("system")), fmt(x.get("operation")), fmt(x.get("check"))])
        st_colors.append(C["highlight"] if judg else None)
    n_j = sum(1 for x in steps if x.get("human_judgment"))
    deck.table_slides("3. 작업 단계 (Step별 절차)", ["Step", "행위", "시스템 / 화면", "입력 / 조작", "결과 확인"],
                      st_rows, [0.7, 2.6, 2.5, 3.4, 2.9], size=10.5,
                      caption=f"총 {len(steps)}단계 · ⚠ 빨간색 = 사람 판단 개입 단계 ({n_j}건, 자동화 부적합 신호)",
                      row_colors=st_colors, align=[CENTER, None, None, None, None])

    # 6. input / output ----------------------------------------------------
    ins = data.get("inputs", [])
    deck.table_slides("4-1. 입력 데이터 명세", ["데이터명", "출처", "형식", "컬럼 / 구조", "건수", "검증 규칙"],
                      [[fmt(x.get("name")), fmt(x.get("source")), fmt(x.get("format")), fmt(x.get("structure")),
                        fmt(x.get("volume")), fmt(x.get("validation"))] for x in ins],
                      [1.9, 1.9, 1.2, 3.1, 1.2, 2.8], caption="Source → Process", bold_cols=[0])
    outs = data.get("outputs", [])
    deck.table_slides("4-2. 출력 데이터 명세", ["데이터명", "저장 위치", "형식", "후속 처리", "알림"],
                      [[fmt(x.get("name")), fmt(x.get("destination")), fmt(x.get("format")),
                        fmt(x.get("follow_up")), fmt(x.get("notify"))] for x in outs],
                      [2.2, 2.8, 1.3, 3.5, 2.3], caption="Process → Destination", bold_cols=[0])

    # 7. exceptions ----------------------------------------------------------
    order = {"비즈니스": 0, "시스템": 1, "데이터": 2}
    exc = sorted(data.get("exceptions", []), key=lambda e: order.get(str(e.get("category")), 9))
    cat_col = {"비즈니스": C["primary"], "시스템": C["warn"], "데이터": C["success"]}
    deck.table_slides("5. 예외 처리 및 분기 조건", ["분류", "예외 유형", "발생 조건", "처리 방법", "알림 대상"],
                      [[fmt(e.get("category")), fmt(e.get("type")), fmt(e.get("condition")),
                        fmt(e.get("handling")), fmt(e.get("notify"))] for e in exc],
                      [1.1, 2.2, 3.2, 3.6, 2.0],
                      caption="비즈니스 분기 / 시스템 예외 / 데이터 예외",
                      cell_fill_fn=lambda i: {0: cat_col.get(str(exc[i].get("category")), C["muted"])} if exc else {},
                      align=[CENTER, None, None, None, None])

    # 8. assessment ----------------------------------------------------------
    s = deck.slide("6. 자동화 가능성 평가", "5개 항목 × 5점 (판단 개입은 개입이 적을수록 고득점)")
    items = [(lab, num((scores.get(key) or {}).get("score")), 5) for key, lab, _ in CRITERIA]
    add_score_bars(s, MARGIN_X, CONTENT_TOP, 6.6, items, row_h=0.56)
    reason_rows = [[lab, fmt((scores.get(key) or {}).get("reason"))] for key, lab, _ in CRITERIA]
    add_table(s, MARGIN_X, CONTENT_TOP + 2.95, [1.7, 4.9], ["항목", "평가 근거"], reason_rows, size=9.5,
              bold_cols=[0])
    rx = 7.75
    add_rect(s, rx, CONTENT_TOP, 4.98, 2.2, C["light"])
    add_text(s, rx + 0.3, CONTENT_TOP + 0.15, 2.6, 0.4, "종합 점수", size=13, bold=True, color=C["muted"])
    add_text(s, rx + 0.3, CONTENT_TOP + 0.55, 2.8, 1.3, f"{total}" if total is not None else "-", size=64,
             bold=True, color=C["primary"], anchor=MIDDLE)
    add_text(s, rx + 1.75, CONTENT_TOP + 1.05, 1.2, 0.6, "/ 25", size=20, color=C["muted"], anchor=MIDDLE)
    add_badge(s, rx + 3.2, CONTENT_TOP + 0.75, 1.5, 0.9, f"등급 {grade}" if total is not None else "미평가",
              grade_color(grade), size=20)
    add_text(s, rx, CONTENT_TOP + 2.4, 4.98, 0.35, "권장 자동화 범위", size=12, bold=True, color=C["primary"])
    add_badge(s, rx, CONTENT_TOP + 2.8, 4.98, 0.5, scope, C["primary"], size=14)
    add_text(s, rx, CONTENT_TOP + 3.45, 4.98, 1.95, [
        ("자동화 구간", {"bold": True, "color": C["success"]}),
        fmt(asm.get("automated_part")),
        ("수동 구간 (사람)", {"bold": True, "color": C["highlight"], "space_before": 6}),
        fmt(asm.get("manual_part")),
    ], size=11)

    # 9. risks ---------------------------------------------------------------
    s = deck.slide("7. 리스크 / 제약사항", "카테고리별 리스크와 영향도 (상/중/하) 및 대응 방안")
    cats = ["기술", "보안·규제", "운영", "조직"]
    risks = data.get("risks", [])
    cw, chh = (CONTENT_W - 0.3) / 2, (5.35 - 0.3) / 2
    for i, cat in enumerate(cats):
        x = MARGIN_X + (i % 2) * (cw + 0.3)
        y = CONTENT_TOP + (i // 2) * (chh + 0.3)
        mine = [r for r in risks if str(r.get("category", "")).replace("/", "·") == cat]
        add_rect(s, x, y, cw, chh, C["light"])
        add_rect(s, x, y, 0.08, chh, C["primary"])
        add_text(s, x + 0.2, y + 0.08, 3, 0.42, cat, size=14, bold=True, color=C["primary"], anchor=MIDDLE)
        if not mine:
            add_text(s, x + 0.2, y + 0.6, cw - 0.4, 0.4, "식별된 리스크 없음 (현업 확인 권장)", size=11,
                     color=C["muted"])
        for j, r in enumerate(mine[:4]):
            yy = y + 0.6 + j * 0.56
            add_badge(s, x + 0.2, yy + 0.04, 0.5, 0.3, fmt(r.get("impact"), "-"), impact_color(r.get("impact")),
                      size=10)
            add_text(s, x + 0.8, yy - 0.04, cw - 1.0, 0.56, [
                (fmt(r.get("item")), {"bold": True, "size": 11}),
                ("→ " + fmt(r.get("mitigation"), "대응 방안 확인 필요"), {"size": 9.5, "color": C["muted"]}),
            ], size=11)
        if len(mine) > 4:
            add_text(s, x + 0.2, y + chh - 0.35, cw - 0.4, 0.3, f"외 {len(mine) - 4}건 (부록 참조)", size=9.5,
                     color=C["muted"])

    # 10. open questions -------------------------------------------------------
    qs = list(data.get("open_questions", []))
    auto = [("업무개요", f"{lab} 수치 확인 (자동화 효과 산정 필수)")
            for k, lab in [("frequency", "수행 빈도"), ("monthly_volume", "월 처리건수"),
                           ("time_per_case_min", "건당 처리시간")] if blank(ov.get(k))]
    known = {q.get("question") for q in qs}
    qs += [{"area": a, "question": q} for a, q in auto if q not in known]
    areas = ["업무개요", "시스템", "Step", "데이터", "예외", "리스크"]
    grouped: dict[str, list[str]] = {}
    for q in qs:
        a = q.get("area") if q.get("area") in areas else "기타"
        grouped.setdefault(a, []).append(q.get("question", ""))
    q_rows = []
    for a in areas + ["기타"]:
        for i, q in enumerate(grouped.get(a, [])):
            q_rows.append([a if i == 0 else "", "□  " + q, ""])
    deck.table_slides("⚠ 현업 인터뷰 필요 항목", ["구분", "확인 사항", "답변 / 확인자"], q_rows or [["-", "없음", ""]],
                      [1.5, 7.6, 3.0], size=10.5,
                      caption=f"자료에서 도출하지 못한 항목 {len(qs)}건 — 인터뷰 후 분석서 v2.0으로 보완",
                      bold_cols=[0], zebra=False)

    # 11. recommendations ----------------------------------------------------
    rec = data.get("recommendations", {})
    s = deck.slide("권고안 및 다음 단계")
    add_chevrons(s, MARGIN_X, CONTENT_TOP, CONTENT_W, 0.75,
                 ["단기  |  누락 정보 보완", "중기  |  PoC / 부분 자동화", "장기  |  확대 / 프로세스 개선"],
                 [C["muted"], C["primary"], C["success"]], size=13)
    cw = (CONTENT_W - 0.6) / 3
    items = {k: rec.get(k) or [TBD] for k in ["short", "mid", "long"]}
    lines = max(sum(est_lines("• " + t, cw - 0.4, 12) for t in v) for v in items.values())
    box_h = min(4.2, max(1.8, lines * 12 * 1.45 / 72 + 0.5))
    for i, key in enumerate(["short", "mid", "long"]):
        x = MARGIN_X + i * (cw + 0.3)
        add_text(s, x, CONTENT_TOP + 1.0, cw, box_h, items[key], size=12, bullets=True,
                 fill=C["light"], margin=0.2, line_spacing=1.15)

    # 12. appendix -----------------------------------------------------------
    gl = data.get("glossary") or [{"term": t, "desc": d} for t, d in DEFAULT_GLOSSARY]
    src = data.get("source_map") or []
    if src:
        deck.table_slides("부록 A. 원본 자료 매핑", ["원본 위치", "내용", "반영 위치"],
                          [[fmt(x.get("source")), fmt(x.get("content")), fmt(x.get("used_in"))] for x in src],
                          [2.4, 6.4, 3.3])
    if len(risks) > 16 or any(len([r for r in risks if r.get("category") == c]) > 4 for c in cats):
        deck.table_slides("부록. 전체 리스크 목록", ["분류", "리스크", "영향도", "대응 방안"],
                          [[fmt(r.get("category")), fmt(r.get("item")), fmt(r.get("impact")),
                            fmt(r.get("mitigation"))] for r in risks], [1.5, 4.6, 1.0, 5.0])
    deck.table_slides("부록. 용어 정의", ["용어", "설명"], [[g["term"], g["desc"]] for g in gl], [1.8, 10.3],
                      bold_cols=[0])

    out.parent.mkdir(parents=True, exist_ok=True)
    deck.save(str(out))
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec", type=Path)
    ap.add_argument("-o", "--out", type=Path)
    a = ap.parse_args()
    data = json.loads(a.spec.read_text(encoding="utf-8-sig"))
    meta = data.get("meta", {})
    out = a.out or a.spec.with_name(f"RPA분석서_{meta.get('task_name', '업무')}_{meta.get('version', 'v1.0')}.pptx")
    path = build(data, out)
    print(f"[OK] {path}  ({len(__import__('pptx').Presentation(str(path)).slides)} slides)")


if __name__ == "__main__":
    main()
