"""Build the RPA 과제 평가표 (discovery + feasibility workbook) from a JSON spec.

Usage:
    python build_assessment.py assessment.json [-o RPA_과제평가표.xlsx]
    python build_assessment.py --intake-template [-o RPA_현업작성양식.xlsx]

Sheets: 0_안내 / 1_과제목록 / 2_타당성평가 / 3_ROI / 4_우선순위 / 5_확인필요 / 6_현업작성양식
Scores, 자동화율 and ROI parameters are input cells; totals, grades, ROI and
priority are live Excel formulas, so the workbook stays correct when edited.

JSON spec: see skills/rpa-feasibility/references/assessment-schema.md
"""
from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

sys.path.insert(0, str(Path(__file__).resolve().parent))
from scoring import CRITERIA, SCOPES, screen, total_score  # noqa: E402

FONT = "Malgun Gothic"
PRIMARY, LIGHT, RED, GREEN, AMBER, GREY = "36454F", "F2F2F2", "E63946", "2A9D8F", "E9A23B", "7A858C"
THIN = Side(style="thin", color="D0D5D9")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
INPUT_FILL = PatternFill("solid", fgColor="FFF8E1")  # 입력 셀 (연노랑)

FREQ = ["일", "주", "월", "분기", "반기", "연", "수시"]
INPUT_TYPES = ["정형 (Excel/CSV/DB)", "반정형 (양식 PDF/메일)", "비정형 (자유 텍스트/이미지)"]
JUDGMENT = ["없음", "일부 (룰 기반)", "필수 (사람 판단)"]


def f(size=10, bold=False, color="212121", italic=False):
    return Font(name=FONT, size=size, bold=bold, color=color, italic=italic)


def fill(c):
    return PatternFill("solid", fgColor=c)


def title(ws, text, sub=None):
    ws["A1"] = text
    ws["A1"].font = f(16, True, PRIMARY)
    if sub:
        ws["A2"] = sub
        ws["A2"].font = f(10, color=GREY, italic=True)
    ws.sheet_view.showGridLines = False


def header(ws, row, cols, widths=None):
    for i, c in enumerate(cols, start=1):
        cell = ws.cell(row=row, column=i, value=c)
        cell.font = f(10, True, "FFFFFF")
        cell.fill = fill(PRIMARY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = BORDER
    ws.row_dimensions[row].height = 30
    if widths:
        for i, w in enumerate(widths, start=1):
            ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = ws.cell(row=row + 1, column=3)


def body(ws, r1, r2, ncols, wrap_cols=(), center_cols=(), input_cols=(), num_fmt=None):
    for r in range(r1, r2 + 1):
        ws.row_dimensions[r].height = None
        for c in range(1, ncols + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = f(10)
            cell.border = BORDER
            cell.alignment = Alignment(vertical="center", wrap_text=c in wrap_cols,
                                       horizontal="center" if c in center_cols else None)
            if c in input_cols:
                cell.fill = INPUT_FILL
            if num_fmt and c in num_fmt:
                cell.number_format = num_fmt[c]


def color_rule(ws, rng, formula, color, font_color="FFFFFF"):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill(color),
                                                   font=Font(name=FONT, bold=True, color=font_color)))


def add_list_validation(ws, rng, items):
    dv = DataValidation(type="list", formula1='"' + ",".join(items) + '"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)


# --------------------------------------------------------------------------- #
def sheet_guide(wb, meta):
    ws = wb.active
    ws.title = "0_안내"
    title(ws, f"RPA 과제 평가표 — {meta.get('project', '')}".rstrip(" —"),
          "과제 발굴 → 1차 스크리닝 → 타당성 평가(5개 항목) → ROI → 우선순위")
    rows = [("작성일", meta.get("date", "")), ("작성자", meta.get("author", "")),
            ("대상 조직", meta.get("organization", ""))]
    for i, (k, v) in enumerate(rows, start=4):
        ws.cell(row=i, column=1, value=k).font = f(10, True, GREY)
        ws.cell(row=i, column=2, value=v).font = f(10)

    r = 8
    ws.cell(row=r, column=1, value="■ 시트 구성").font = f(12, True, PRIMARY)
    sheets = [
        ("1_과제목록", "발굴된 자동화 후보 업무와 1차 스크리닝 결과"),
        ("2_타당성평가", "5개 항목 × 5점 평가 → 총점·등급·권장 범위 (노란 셀 = 입력)"),
        ("3_ROI", "월/연 절감시간, FTE, 절감액 (상단 파라미터 수정 가능)"),
        ("4_우선순위", "적합도(60%) + 효과(40%) 종합점수와 우선순위 구분"),
        ("5_확인필요", "현업 인터뷰로 확인해야 할 항목"),
        ("6_현업작성양식", "현업 담당자 배포용 업무 조사 양식"),
    ]
    for i, (a, b) in enumerate(sheets, start=r + 1):
        ws.cell(row=i, column=1, value=a).font = f(10, True)
        ws.cell(row=i, column=2, value=b).font = f(10)

    r = r + len(sheets) + 2
    ws.cell(row=r, column=1, value="■ 타당성 평가 기준 (각 1–5점)").font = f(12, True, PRIMARY)
    header_row = r + 1
    for i, h in enumerate(["평가 항목", "5점", "3점", "1점"], start=1):
        c = ws.cell(row=header_row, column=i, value=h)
        c.font = f(10, True, "FFFFFF")
        c.fill = fill(PRIMARY)
        c.border = BORDER
    for j, (_, lab, (s5, s3, s1)) in enumerate(CRITERIA, start=header_row + 1):
        for i, v in enumerate([lab, s5, s3, s1], start=1):
            c = ws.cell(row=j, column=i, value=v)
            c.font = f(10, i == 1)
            c.border = BORDER
    r = header_row + len(CRITERIA) + 2
    notes = [
        "■ 등급: 상 20–25점 → " + SCOPES["상"] + " / 중 12–19점 → " + SCOPES["중"] + " / 하 5–11점 → " + SCOPES["하"],
        "■ 1차 스크리닝: 판단 필수 → 부적합 가능 / 비정형 입력·월 5시간 미만 → 검토 필요 / 수치 누락 → 정보 부족",
        "■ ROI: 월 절감시간 = 월 처리건수 × 건당 처리시간(분) ÷ 60 × 예상 자동화율",
        "■ 종합점수(100) = 총점/25 × 60 + MIN(월 절감시간 ÷ (효과기준 × 2), 1) × 40",
        "■ 우선순위: 1순위 = 총점 20↑ 이고 월 절감시간 ≥ 효과기준 / 2순위 = 둘 중 하나 / 3순위 = 그 외 / 보류 = 총점 12 미만",
        "※ 노란색 셀은 입력 셀입니다. 나머지 계산 셀은 수식이므로 직접 수정하지 마세요.",
    ]
    for i, n in enumerate(notes):
        ws.cell(row=r + i, column=1, value=n).font = f(10, color=RED if n.startswith("※") else "212121")
    ws.column_dimensions["A"].width = 22
    for col in "BCD":
        ws.column_dimensions[col].width = 34


def sheet_candidates(wb, cands):
    ws = wb.create_sheet("1_과제목록")
    title(ws, "1. 과제 목록 (발굴)", "현업 인터뷰·업무 조사로 수집한 자동화 후보 업무")
    cols = ["ID", "업무명", "부서", "담당자", "업무 설명", "수행 빈도", "월 처리건수", "건당 처리시간(분)",
            "월 소요시간(h)", "사용 시스템", "입력 데이터 유형", "판단 개입", "출처", "1차 스크리닝", "스크리닝 사유"]
    header(ws, 3, cols, [8, 24, 12, 10, 36, 9, 10, 10, 10, 22, 18, 14, 16, 12, 30])
    for i, c in enumerate(cands):
        r = 4 + i
        res, why = screen(c)
        vals = [c.get("id"), c.get("task_name"), c.get("department"), c.get("owner"), c.get("description"),
                c.get("frequency"), c.get("monthly_volume"), c.get("minutes_per_case"),
                f"=IF(AND(ISNUMBER(G{r}),ISNUMBER(H{r})),G{r}*H{r}/60,\"\")",
                c.get("systems"), c.get("input_type"), c.get("human_judgment"), c.get("source"), res, why]
        for j, v in enumerate(vals, start=1):
            ws.cell(row=r, column=j, value=v)
    last = 3 + max(len(cands), 1)
    body(ws, 4, last, len(cols), wrap_cols={2, 5, 10, 13, 15}, center_cols={1, 6, 12, 14},
         input_cols={7, 8}, num_fmt={7: "#,##0", 8: "#,##0.0", 9: "#,##0.0"})
    rng = f"N4:N{last}"
    color_rule(ws, rng, 'N4="적합 후보"', GREEN)
    color_rule(ws, rng, 'N4="검토 필요"', AMBER)
    color_rule(ws, rng, 'N4="부적합 가능"', RED)
    color_rule(ws, rng, 'N4="정보 부족"', GREY)
    return ws


def sheet_feasibility(wb, cands):
    ws = wb.create_sheet("2_타당성평가")
    title(ws, "2. 타당성 평가", "각 항목 1–5점 (노란 셀 입력) · 판단 개입은 개입이 적을수록 고득점")
    cols = ["ID", "업무명"] + [lab for _, lab, _ in CRITERIA] + ["총점", "등급", "권장 자동화 범위",
                                                             "구현 난이도", "평가 근거", "비고"]
    header(ws, 3, cols, [8, 24, 10, 10, 10, 10, 10, 10, 8, 22, 10, 50, 24])
    for i, c in enumerate(cands):
        r = 4 + i
        ws.cell(row=r, column=1, value=f"='1_과제목록'!A{r}")
        ws.cell(row=r, column=2, value=f"='1_과제목록'!B{r}")
        sc = c.get("scores", {}) or {}
        reasons = []
        for j, (key, lab, _) in enumerate(CRITERIA):
            v = sc.get(key)
            s = v.get("score") if isinstance(v, dict) else v
            ws.cell(row=r, column=3 + j, value=s)
            if isinstance(v, dict) and v.get("reason"):
                reasons.append(f"[{lab}] {v['reason']}")
        ws.cell(row=r, column=8, value=f'=IF(COUNT(C{r}:G{r})<5,"평가 미완료",SUM(C{r}:G{r}))')
        ws.cell(row=r, column=9, value=f'=IF(ISNUMBER(H{r}),IF(H{r}>=20,"상",IF(H{r}>=12,"중","하")),"-")')
        scope = c.get("scope") or (f'=IF(I{r}="상","{SCOPES["상"]}",IF(I{r}="중","{SCOPES["중"]}",'
                                   f'IF(I{r}="하","{SCOPES["하"]}","-")))')
        ws.cell(row=r, column=10, value=scope)
        ws.cell(row=r, column=11, value=c.get("complexity"))
        ws.cell(row=r, column=12, value="\n".join(reasons) or c.get("score_reason"))
        ws.cell(row=r, column=13, value=c.get("notes"))
    last = 3 + max(len(cands), 1)
    body(ws, 4, last, len(cols), wrap_cols={2, 10, 12, 13}, center_cols={1, 3, 4, 5, 6, 7, 8, 9, 11},
         input_cols={3, 4, 5, 6, 7, 11})
    dv = DataValidation(type="whole", operator="between", formula1="1", formula2="5", allow_blank=True,
                        error="1~5 사이 정수를 입력하세요.", showErrorMessage=True)
    ws.add_data_validation(dv)
    dv.add(f"C4:G{last + 50}")
    add_list_validation(ws, f"K4:K{last + 50}", ["하", "중", "상"])
    rng = f"I4:I{last}"
    color_rule(ws, rng, 'I4="상"', GREEN)
    color_rule(ws, rng, 'I4="중"', AMBER)
    color_rule(ws, rng, 'I4="하"', RED)


def sheet_roi(wb, cands, meta):
    ws = wb.create_sheet("3_ROI")
    title(ws, "3. ROI 산정")
    params = [("월 근무시간 (h / 1 FTE)", meta.get("work_hours_per_month", 160)),
              ("시간당 인건비 (원)", meta.get("hourly_cost")),
              ("효과 기준 (월 절감시간 h)", meta.get("effect_threshold_hours", 20))]
    for i, (k, v) in enumerate(params, start=2):
        ws.cell(row=i, column=1, value=k).font = f(10, True, GREY)
        c = ws.cell(row=i, column=2, value=v)
        c.fill = INPUT_FILL
        c.border = BORDER
        c.number_format = "#,##0"
    cols = ["ID", "업무명", "월 처리건수", "건당 처리시간(분)", "월 소요시간(h)", "예상 자동화율",
            "월 절감시간(h)", "연간 절감시간(h)", "절감 FTE", "연간 절감액(원)"]
    hr = 6
    header(ws, hr, cols, [26, 24, 11, 12, 12, 11, 12, 13, 10, 16])
    for i, c in enumerate(cands):
        r, src = hr + 1 + i, 4 + i
        def ref(col):  # empty source cell -> "" (a plain reference would show 0)
            return f"=IF('1_과제목록'!{col}{src}=\"\",\"\",'1_과제목록'!{col}{src})"

        vals = [f"='1_과제목록'!A{src}", f"='1_과제목록'!B{src}", ref("G"), ref("H"), ref("I"),
                c.get("expected_automation_rate"),
                f'=IF(AND(ISNUMBER(E{r}),ISNUMBER(F{r})),E{r}*F{r},"")',
                f'=IF(ISNUMBER(G{r}),G{r}*12,"")',
                f'=IF(ISNUMBER(G{r}),G{r}/$B$2,"")',
                f'=IF(AND(ISNUMBER(H{r}),ISNUMBER($B$3)),H{r}*$B$3,"")']
        for j, v in enumerate(vals, start=1):
            ws.cell(row=r, column=j, value=v)
    last = hr + max(len(cands), 1)
    body(ws, hr + 1, last, len(cols), wrap_cols={2}, center_cols={1}, input_cols={6},
         num_fmt={3: "#,##0", 4: "#,##0.0", 5: "#,##0.0", 6: "0%", 7: "#,##0.0", 8: "#,##0", 9: "0.00",
                  10: "#,##0"})
    tr = last + 1
    ws.cell(row=tr, column=2, value="합계").font = f(10, True)
    for col in "EGHIJ":
        c = ws[f"{col}{tr}"]
        c.value = f"=SUM({col}{hr + 1}:{col}{last})"
        c.font = f(10, True)
        c.border = BORDER
        c.number_format = "0.00" if col == "I" else "#,##0"
    ws.freeze_panes = ws.cell(row=hr + 1, column=3)


def sheet_priority(wb, cands, meta):
    ws = wb.create_sheet("4_우선순위")
    title(ws, "4. 우선순위", "종합점수(100) = 적합도 60% + 효과 40%  ·  순위는 종합점수 기준")
    cols = ["순위", "ID", "업무명", "총점(25)", "등급", "월 절감시간(h)", "종합점수(100)", "우선순위 구분", "다음 단계"]
    header(ws, 3, cols, [7, 8, 26, 10, 8, 13, 13, 26, 22])
    thr = float(meta.get("effect_threshold_hours", 20))

    def composite(c):
        t = total_score(c.get("scores", {}))
        try:
            saved = float(c["monthly_volume"]) * float(c["minutes_per_case"]) / 60 * float(
                c["expected_automation_rate"])
        except (KeyError, TypeError, ValueError):
            saved = 0
        return -1 if t is None else t / 25 * 60 + min(saved / (thr * 2), 1) * 40

    order = sorted(range(len(cands)), key=lambda i: composite(cands[i]), reverse=True)
    n = len(cands)
    last = 3 + max(n, 1)
    for k, i in enumerate(order):
        r, fr, rr = 4 + k, 4 + i, 7 + i  # row here / feasibility row / ROI row
        vals = [f'=IFERROR(RANK(G{r},$G$4:$G${last}),"-")',
                f"='2_타당성평가'!A{fr}", f"='2_타당성평가'!B{fr}", f"='2_타당성평가'!H{fr}",
                f"='2_타당성평가'!I{fr}", f"='3_ROI'!G{rr}",
                f"=IF(ISNUMBER(D{r}),ROUND(D{r}/25*60+MIN(N(F{r})/('3_ROI'!$B$4*2),1)*40,1),\"\")",
                f'=IF(NOT(ISNUMBER(D{r})),"평가 미완료",IF(D{r}<12,"보류 (프로세스 개선 선행)",'
                f'IF(AND(D{r}>=20,N(F{r})>=\'3_ROI\'!$B$4),"1순위 (Quick Win)",'
                f'IF(OR(D{r}>=20,N(F{r})>=\'3_ROI\'!$B$4),"2순위","3순위"))))',
                f'=IF(LEFT(H{r},1)="1","PDD 작성 착수",IF(LEFT(H{r},1)="2","현업 인터뷰 후 PDD",'
                f'IF(LEFT(H{r},1)="3","후순위 검토",IF(LEFT(H{r},2)="보류","프로세스 표준화 선행","정보 보완"))))']
        for j, v in enumerate(vals, start=1):
            ws.cell(row=r, column=j, value=v)
    body(ws, 4, last, len(cols), wrap_cols={3}, center_cols={1, 2, 4, 5, 7},
         num_fmt={6: "#,##0.0", 7: "0.0"})
    rng = f"H4:H{last}"
    color_rule(ws, rng, 'LEFT(H4,1)="1"', GREEN)
    color_rule(ws, rng, 'LEFT(H4,1)="2"', AMBER)
    color_rule(ws, rng, 'LEFT(H4,2)="보류"', RED)


def sheet_questions(wb, cands):
    ws = wb.create_sheet("5_확인필요")
    title(ws, "5. 확인 필요 사항", "현업 인터뷰로 보완할 항목 — 상태를 '확인완료'로 바꾸며 관리")
    cols = ["ID", "업무명", "구분", "확인 사항", "상태", "답변 / 비고"]
    header(ws, 3, cols, [8, 24, 12, 60, 11, 40])
    rows = []
    for c in cands:
        for q in c.get("open_questions", []) or []:
            area, text = (q.get("area", ""), q.get("question", "")) if isinstance(q, dict) else ("", q)
            rows.append([c.get("id"), c.get("task_name"), area, text])
        for key, lab in [("monthly_volume", "월 처리건수"), ("minutes_per_case", "건당 처리시간")]:
            if c.get(key) in (None, ""):
                rows.append([c.get("id"), c.get("task_name"), "업무개요", f"{lab} 수치 확인 (ROI 산정 필수)"])
        if total_score(c.get("scores", {})) is None:
            rows.append([c.get("id"), c.get("task_name"), "평가", "타당성 평가 점수 미완료 항목 확인"])
        if c.get("expected_automation_rate") in (None, ""):
            rows.append([c.get("id"), c.get("task_name"), "ROI", "예상 자동화율 확인"])
    for i, row in enumerate(rows):
        for j, v in enumerate(row + ["미확인", ""], start=1):
            ws.cell(row=4 + i, column=j, value=v)
    last = 3 + max(len(rows), 1)
    body(ws, 4, last, len(cols), wrap_cols={2, 4, 6}, center_cols={1, 3, 5}, input_cols={5, 6})
    add_list_validation(ws, f"E4:E{last + 50}", ["미확인", "확인중", "확인완료"])
    color_rule(ws, f"E4:E{last}", 'E4="미확인"', RED)
    color_rule(ws, f"E4:E{last}", 'E4="확인완료"', GREEN)


INTAKE_COLS = [
    ("업무명", "한 줄로 (예: 월말 거래처 정산 ERP 등록)", 26),
    ("부서", "", 12),
    ("담당자", "", 10),
    ("업무 설명", "무엇을, 왜, 결과물은?", 40),
    ("수행 빈도", "일/주/월/분기/반기/연/수시", 10),
    ("월 처리건수", "숫자 (월 평균)", 11),
    ("건당 처리시간(분)", "숫자 (1건 평균)", 12),
    ("사용 시스템", "ERP, 그룹웨어, Excel 등 모두", 24),
    ("입력 데이터 형태", "정형/반정형/비정형", 22),
    ("판단 개입", "사람 판단이 필요한 정도", 16),
    ("예외 상황", "자주 생기는 오류·예외와 빈도", 30),
    ("비고", "시스템 변경 예정, 보안 이슈 등", 26),
]


def sheet_intake(wb, standalone=False):
    ws = wb.active if standalone else wb.create_sheet("6_현업작성양식")
    if standalone:
        ws.title = "6_현업작성양식"
    title(ws, "RPA 자동화 후보 업무 조사서", "반복적으로 하는 업무를 한 행에 하나씩 적어 주세요. 노란 셀만 입력합니다.")
    cols = [c for c, _, _ in INTAKE_COLS]
    header(ws, 3, cols, [w for _, _, w in INTAKE_COLS])
    for j, (_, desc, _) in enumerate(INTAKE_COLS, start=1):
        c = ws.cell(row=4, column=j, value=desc)
        c.font = f(9, color=GREY, italic=True)
        c.alignment = Alignment(wrap_text=True, vertical="center")
        c.border = BORDER
    example = ["(예시) 월말 거래처 정산 ERP 등록", "재무팀", "홍길동", "거래처별 정산 엑셀을 ERP에 전표로 입력",
               "월", 300, 4, "SAP, Excel, Outlook", INPUT_TYPES[0], JUDGMENT[1], "거래처 미등록 월 5건 내외",
               "내년 차세대 ERP 예정 없음"]
    for j, v in enumerate(example, start=1):
        c = ws.cell(row=5, column=j, value=v)
        c.font = f(9, color=GREY, italic=True)
        c.border = BORDER
        c.alignment = Alignment(wrap_text=True, vertical="center")
    body(ws, 6, 40, len(cols), wrap_cols=set(range(1, len(cols) + 1)), input_cols=set(range(1, len(cols) + 1)))
    add_list_validation(ws, "E6:E200", FREQ)
    add_list_validation(ws, "I6:I200", INPUT_TYPES)
    add_list_validation(ws, "J6:J200", JUDGMENT)
    for col in "FG":
        dv = DataValidation(type="decimal", operator="greaterThanOrEqual", formula1="0", allow_blank=True,
                            error="숫자를 입력하세요.", showErrorMessage=True)
        ws.add_data_validation(dv)
        dv.add(f"{col}6:{col}200")
    ws.freeze_panes = "B6"


def build(data: dict, out: Path) -> Path:
    meta, cands = data.get("meta", {}), data.get("candidates", [])
    wb = Workbook()
    sheet_guide(wb, meta)
    sheet_candidates(wb, cands)
    sheet_feasibility(wb, cands)
    sheet_roi(wb, cands, meta)
    sheet_priority(wb, cands, meta)
    sheet_questions(wb, cands)
    sheet_intake(wb)
    out.parent.mkdir(parents=True, exist_ok=True)
    wb.save(out)
    return out


def main() -> None:
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("spec", type=Path, nargs="?")
    ap.add_argument("-o", "--out", type=Path)
    ap.add_argument("--intake-template", action="store_true", help="현업 배포용 빈 조사 양식만 생성")
    a = ap.parse_args()
    if a.intake_template:
        out = a.out or Path("RPA_현업작성양식.xlsx")
        wb = Workbook()
        sheet_intake(wb, standalone=True)
        wb.save(out)
        print(f"[OK] {out}")
        return
    if not a.spec:
        ap.error("spec JSON is required (or use --intake-template)")
    data = json.loads(a.spec.read_text(encoding="utf-8-sig"))
    out = a.out or a.spec.with_name("RPA_과제평가표.xlsx")
    print(f"[OK] {build(data, out)}  ({len(data.get('candidates', []))} candidates)")


if __name__ == "__main__":
    main()
