"""Render an RPA SDD (솔루션 설계서) PowerPoint from a JSON spec.

Usage:
    python render_sdd.py sdd.json [-o RPA설계서_<업무명>_v1.0.pptx]

JSON spec: see skills/rpa-solution-design/references/sdd-schema.md
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from pptx_kit import (C, CENTER, CONTENT_TOP, CONTENT_W, MARGIN_X, TBD, Deck, add_card,  # noqa: E402
                      add_chevrons, add_table, add_text, blank, fmt, paginate)

DOC_LABEL = "RPA 솔루션 설계서"
ACTOR = {"bot": C["success"], "봇": C["success"], "사람": C["highlight"], "human": C["highlight"],
         "시스템": C["primary"], "system": C["primary"]}


def actor_color(a) -> str:
    return ACTOR.get(str(a).strip().lower(), C["muted"])


def build(data: dict, out: Path) -> Path:
    meta = data.get("meta", {})
    ov = data.get("overview", {})
    task = meta.get("task_name") or TBD
    deck = Deck(DOC_LABEL, task)

    # 1. cover ---------------------------------------------------------------
    deck.cover("RPA 솔루션 설계서", task, [
        ("작성일", meta.get("date")),
        ("RPA 플랫폼", meta.get("rpa_platform")),
        ("작성자", meta.get("author")),
        ("버전", meta.get("version") or "v1.0"),
        ("참조 PDD", meta.get("pdd_ref")),
    ], "SDD")

    # 2. overview ------------------------------------------------------------
    s = deck.slide("1. 설계 개요")
    rows = [["자동화 목표", fmt(ov.get("goal"))],
            ["자동화 유형", fmt(ov.get("automation_type"))],
            ["RPA 플랫폼", fmt(meta.get("rpa_platform"))],
            ["봇 구조", fmt(data.get("architecture", {}).get("pattern"))],
            ["예상 효과", fmt(ov.get("expected_effect"))]]
    add_table(s, MARGIN_X, CONTENT_TOP, [2.0, 10.1], ["항목", "내용"], rows, size=11, bold_cols=[0])
    cw = (CONTENT_W - 0.3) / 2
    y = CONTENT_TOP + 2.75
    add_card(s, MARGIN_X, y, cw, 2.6, "자동화 범위 (In-Scope)", ov.get("scope_in") or [TBD],
             accent=C["success"], bullets=True)
    add_card(s, MARGIN_X + cw + 0.3, y, cw, 2.6, "제외 범위 (Out-of-Scope)", ov.get("scope_out") or ["없음"],
             accent=C["highlight"], bullets=True)

    # 3. to-be flow ----------------------------------------------------------
    flow = data.get("tobe_flow", [])

    def flow_legend(sl, _):
        lx = MARGIN_X + CONTENT_W - 4.0
        for i, (lab, col) in enumerate([("Bot", C["success"]), ("사람", C["highlight"]), ("시스템", C["primary"])]):
            add_text(sl, lx + i * 1.35, 1.04, 1.25, 0.28, lab, size=9.5, bold=True, color=C["white"],
                     align=CENTER, fill=col, auto_tbd=False)

    deck.table_slides("2. To-Be 프로세스", ["No", "수행 주체", "처리 내용", "비고"],
                      [[fmt(x.get("no") or i), fmt(x.get("actor")), fmt(x.get("action")), fmt(x.get("note"), "-")]
                       for i, x in enumerate(flow, start=1)],
                      [0.7, 1.4, 6.6, 3.4], size=10.5, caption="자동화 후 업무 흐름 (사람 개입 구간 강조)",
                      cell_fill_fn=lambda i: {1: actor_color(flow[i].get("actor"))} if flow else {},
                      align=[CENTER, CENTER, None, None], before=flow_legend)

    # 4. architecture --------------------------------------------------------
    arch = data.get("architecture", {})
    mods = arch.get("modules", [])
    m_rows = [[fmt(m.get("id")), fmt(m.get("name")), fmt(m.get("role")), fmt(m.get("input"), "-"),
               fmt(m.get("output"), "-"), fmt(m.get("systems"), "-")] for m in mods]
    widths = [0.8, 2.2, 3.5, 1.9, 1.9, 1.8]
    pages = paginate(m_rows, widths, 10, 3.9)
    for p_i, page in enumerate(pages):
        t = "3. 봇 아키텍처" + ("" if len(pages) == 1 else f" ({p_i + 1}/{len(pages)})")
        s = deck.slide(t, fmt(arch.get("description"), "모듈 구성 및 호출 흐름"))
        labels = [f"{m.get('id', '')} {m.get('name', '')}".strip() for m in mods[:7]] or [TBD]
        add_chevrons(s, MARGIN_X, CONTENT_TOP, CONTENT_W, 0.8, labels, size=10.5)
        add_table(s, MARGIN_X, CONTENT_TOP + 1.1, widths, ["ID", "모듈명", "역할", "입력", "출력", "대상 시스템"],
                  page or [[TBD, "", "", "", "", ""]], size=10, align=[CENTER] + [None] * 5, bold_cols=[1])

    # 5. module detail -------------------------------------------------------
    det = [[fmt(m.get("id")) + "\n" + fmt(m.get("name")),
            "\n".join(f"{i}. {st}" for i, st in enumerate(m.get("steps", []), start=1))]
           for m in mods if m.get("steps")]
    if det:
        deck.table_slides("4. 모듈 상세 처리 로직", ["모듈", "처리 단계"], det, [2.4, 9.7], size=10, bold_cols=[0])

    # 6. config --------------------------------------------------------------
    cfg = data.get("config", [])
    cfg_rows = []
    for x in cfg:
        is_cred = "cred" in str(x.get("type", "")).lower() or "자격" in str(x.get("type", ""))
        cfg_rows.append([fmt(x.get("key")), fmt(x.get("type")),
                         "(Credential Store에 보관)" if is_cred else fmt(x.get("value_example"), "-"),
                         fmt(x.get("description"))])
    deck.table_slides("5. 설정 항목 (Config / Asset)", ["Key", "구분", "예시 값", "설명"], cfg_rows,
                      [3.0, 1.6, 3.2, 4.3], size=10.5,
                      caption="환경별(개발/운영)로 달라지는 값은 코드에 하드코딩하지 않고 설정으로 분리",
                      bold_cols=[0])

    # 7. queue ---------------------------------------------------------------
    q = data.get("queue")
    if q:
        s = deck.slide("6. 큐 / 트랜잭션 설계", "Dispatcher가 적재하고 Performer가 건별로 처리")
        info = [["큐 이름", fmt(q.get("name"))], ["트랜잭션 단위", fmt(q.get("unit"))],
                ["고유 참조키 (Reference)", fmt(q.get("reference"))], ["재시도 정책", fmt(q.get("retry"))],
                ["처리 순서 / 우선순위", fmt(q.get("priority"), "-")]]
        add_table(s, MARGIN_X, CONTENT_TOP, [2.6, 3.3], ["항목", "내용"], info, size=10.5, bold_cols=[0])
        fields = [[fmt(fl.get("name")), fmt(fl.get("type"), "-"), fmt(fl.get("description"), "-")]
                  for fl in q.get("item_fields", [])]
        add_table(s, MARGIN_X + 6.2, CONTENT_TOP, [2.0, 1.1, 2.8], ["Item 필드", "타입", "설명"],
                  fields or [[TBD, "", ""]], size=10.5, bold_cols=[0])

    # 8. exceptions ----------------------------------------------------------
    pol = data.get("exception_policy", {})
    exc = data.get("exceptions", [])
    s = None

    def exc_header(sl, p_i):
        if p_i:
            return
        cw2 = (CONTENT_W - 0.3) / 2
        add_card(sl, MARGIN_X, CONTENT_TOP, cw2, 1.35, "Business Exception (업무 예외)",
                 fmt(pol.get("business"), "재시도 없이 해당 건을 예외 처리하고 담당자에게 통보"),
                 accent=C["primary"], body_size=10.5)
        add_card(sl, MARGIN_X + cw2 + 0.3, CONTENT_TOP, cw2, 1.35, "System Exception (시스템 예외)",
                 fmt(pol.get("system"), "스크린샷 저장 → 앱 재기동 → 최대 N회 재시도 후 실패 처리"),
                 accent=C["warn"], body_size=10.5)

    deck.table_slides("7. 예외 처리 설계", ["구분", "예외명", "감지 방법", "처리 방법", "재시도", "알림"],
                      [[fmt(e.get("type")), fmt(e.get("name")), fmt(e.get("detection")), fmt(e.get("handling")),
                        fmt(e.get("retry"), "-"), fmt(e.get("notify"), "-")] for e in exc],
                      [1.2, 2.1, 2.8, 3.4, 0.9, 1.7], size=10, top=CONTENT_TOP + 1.55,
                      cell_fill_fn=lambda i: {0: C["primary"] if "bus" in str(exc[i].get("type", "")).lower()
                                              or "업무" in str(exc[i].get("type", "")) else C["warn"]} if exc else {},
                      align=[CENTER, None, None, None, CENTER, None], before=exc_header)

    # 9. logging & notification ---------------------------------------------
    lg = data.get("logging", {})
    noti = data.get("notifications", [])
    s = deck.slide("8. 로그 및 알림 설계", fmt(lg.get("policy"), "처리 이력은 감사 추적이 가능하도록 건별로 기록"))
    add_text(s, MARGIN_X, CONTENT_TOP, 5.9, 0.35, "로그 기록 지점", size=13, bold=True, color=C["primary"])
    add_table(s, MARGIN_X, CONTENT_TOP + 0.4, [1.9, 4.0], ["시점", "기록 내용"],
              [[fmt(x.get("point")), fmt(x.get("content"))] for x in lg.get("items", [])] or [[TBD, ""]],
              size=10, bold_cols=[0])
    rx = MARGIN_X + 6.2
    add_text(s, rx, CONTENT_TOP, 5.9, 0.35, "알림", size=13, bold=True, color=C["primary"])
    add_table(s, rx, CONTENT_TOP + 0.4, [1.5, 1.4, 1.0, 2.0], ["발생 조건", "수신자", "채널", "내용"],
              [[fmt(x.get("trigger")), fmt(x.get("recipient")), fmt(x.get("channel")), fmt(x.get("content"))]
               for x in noti] or [[TBD, "", "", ""]], size=10)
    if lg.get("report"):
        add_text(s, MARGIN_X, 6.25, CONTENT_W, 0.5, "처리 결과 리포트: " + fmt(lg.get("report")), size=11,
                 fill=C["light"], margin=0.12)

    # 10. environment ---------------------------------------------------------
    env = data.get("environment", {})
    labels = [("machine", "실행 머신"), ("schedule", "실행 스케줄"), ("runtime", "예상 수행시간"),
              ("licenses", "필요 라이선스"), ("credential_store", "자격증명 보관"), ("network", "네트워크 / 망분리"),
              ("orchestrator", "관리 콘솔 (Orchestrator 등)"), ("dependencies", "사전 설치 / 의존성")]
    env_rows = [[lab, fmt(env.get(k))] for k, lab in labels if k in env or k in ("machine", "schedule")]
    env_rows += [[k, fmt(v)] for k, v in env.items() if k not in dict(labels)]
    deck.table_slides("9. 실행 환경", ["항목", "내용"], env_rows, [3.0, 9.1], size=11, bold_cols=[0])

    # 11. security -----------------------------------------------------------
    sec = data.get("security", [])
    deck.table_slides("10. 보안 고려사항", ["항목", "적용 방안"],
                      [[fmt(x.get("item")), fmt(x.get("measure"))] for x in sec], [3.0, 9.1], size=11,
                      bold_cols=[0])

    # 12. test ---------------------------------------------------------------
    tc = data.get("test_cases", [])
    deck.table_slides("11. 테스트 시나리오 (UAT)", ["ID", "시나리오", "유형", "테스트 데이터", "기대 결과"],
                      [[fmt(t.get("id")), fmt(t.get("scenario")), fmt(t.get("type")), fmt(t.get("data"), "-"),
                        fmt(t.get("expected"))] for t in tc],
                      [0.9, 3.8, 0.9, 3.0, 3.5], size=10,
                      cell_fill_fn=lambda i: {2: C["success"] if "정상" in str(tc[i].get("type")) else C["warn"]}
                      if tc else {}, align=[CENTER, None, CENTER, None, None])

    # 13. plan ---------------------------------------------------------------
    plan = data.get("plan", [])
    deck.table_slides("12. 개발 계획", ["단계", "주요 작업", "기간", "담당"],
                      [[fmt(p.get("phase")), fmt(p.get("task")), fmt(p.get("duration")), fmt(p.get("owner"), "-")]
                       for p in plan], [2.0, 6.6, 1.8, 1.7], size=10.5, bold_cols=[0])

    # 14. open questions -------------------------------------------------------
    qs = data.get("open_questions", [])
    deck.table_slides("⚠ 설계 확정 전 확인 필요 사항", ["구분", "확인 사항", "답변 / 확인자"],
                      [[fmt(q.get("area"), "-") if isinstance(q, dict) else "-",
                        "□  " + (q.get("question", "") if isinstance(q, dict) else str(q)), ""] for q in qs]
                      or [["-", "없음", ""]], [1.6, 7.5, 3.0], size=10.5, bold_cols=[0], zebra=False,
                      caption=f"{len(qs)}건 — 개발 착수 전 현업 / IT / 보안 담당과 확정")

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
    out = a.out or a.spec.with_name(f"RPA설계서_{meta.get('task_name', '업무')}_{meta.get('version', 'v1.0')}.pptx")
    path = build(data, out)
    print(f"[OK] {path}  ({len(__import__('pptx').Presentation(str(path)).slides)} slides)")


if __name__ == "__main__":
    main()
