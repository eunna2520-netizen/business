"""토목 시설물 공기산정 시트 2개와 법정공휴일 시트를 워크북에 추가한다."""
import json

from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

INPUT = PatternFill("solid", fgColor="FFF2CC")
HEAD = PatternFill("solid", fgColor="D9E1F2")
RESULT = PatternFill("solid", fgColor="E2EFDA")
BOLD = Font(bold=True)
THIN = Side(style="thin", color="999999")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

# (시설물, 산정공식 문자열, 하한(백만원), 상한(백만원), 필요변수 L,W,BL,D,S,SL,RL)
FACILITIES = [
    ("도로포장", "Y = -637.009 + 173.198×ln(L) + 0.049×C", 0, 35000, (1, 0, 0, 0, 0, 0, 0)),
    ("도로(토공+교량)", "Y = -160.855 - 14.288×W + 164.473×ln(L) - 1.474×BL + 0.052×C", 0, 35000, (1, 1, 1, 0, 0, 0, 0)),
    ("농업용수", "Y = -2251.569 + 415.137×ln(C)", 1000, 20000, (0, 0, 0, 0, 0, 0, 0)),
    ("상수도", "Y = -1175.174 + 119.731×S - 0.273×D + 222.426×ln(C)", 0, 8000, (0, 0, 0, 1, 1, 0, 0)),
    ("하수도", "Y = -452.433 + 98.364×ln(SL) + 0.083×C", 0, 15000, (0, 0, 0, 0, 0, 1, 0)),
    ("철도(궤도)", "Y = -1723.316 - 74.260×ln(RL) + 372.266×ln(C)", 0, 120000, (0, 0, 0, 0, 0, 0, 1)),
]
VARS = [("L", "도로연장(m)"), ("W", "도로폭원(m)"), ("BL", "교량연장(m)"), ("D", "관경(mm)"),
        ("S", "양수장/배수장/가압장 개수"), ("SL", "하수도 연장(m)"), ("RL", "궤도연장(m)")]
# 가이드라인 부록 2·3의 조건 순서 (data/weather.json과 같다). 기본 적용: 철근콘크리트공사 예시
COND_DEFAULT_ON = {1, 5, 6, 9, 13}


def _hdr(ws, row, cols):
    for c in cols:
        cell = ws.cell(row, c)
        cell.fill, cell.font, cell.border = HEAD, BOLD, BOX
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)


def add_holiday_sheet(wb, old_sheet="별표1-공유일"):
    ws = wb.create_sheet("법정공휴일")
    ws["A1"] = "법정 공휴일수 (2019-2035년)  ※ 2019~2025: 기존 별표1, 2026~2035: 2026년 가이드라인 부록 1"
    ws["A1"].font = BOLD
    ws.append([])
    ws.append(["연도", *[f"{m}월" for m in range(1, 13)], "소계"])
    _hdr(ws, 3, range(1, 15))
    old = wb[old_sheet]
    rows = {}
    for r in range(4, 14):
        y = int(str(old.cell(r, 2).value)[:4])
        rows[y] = [old.cell(r, c).value for c in range(3, 15)]
    for y, v in json.load(open("data/holidays_2026_2035.json")).items():
        rows[int(y)] = v
    for y in sorted(rows):
        if y < 2019 or y > 2035:
            continue
        ws.append([y, *rows[y], f"=SUM(B{ws.max_row + 1}:M{ws.max_row + 1})"])
    return ws.max_row  # 마지막 행


def add_facility_sheet(wb):
    ws = wb.create_sheet("시설물별공기(토목)")
    ws["A1"] = "토목 시설물 공사기간 산정 (가이드라인 부록 5 산정공식)"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = "※ 노란 칸에 입력합니다. 산정값 Y는 비작업일수를 포함하므로, 준비기간과 정리기간만 더합니다. 적용범위를 넘으면 산정공식을 쓰지 말고 '공기산정(1일작업량)' 시트로 산정합니다."
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:F2")
    ws.row_dimensions[2].height = 45

    ws["A4"], ws["B4"] = "시설물", "도로포장"
    ws["A5"], ws["B5"] = "총공사비 C (백만원)", None
    ws["C5"] = "추정금액(추정가격+부가가치세+관급자재비)"
    dv = DataValidation(type="list", formula1='"' + ",".join(f[0] for f in FACILITIES) + '"', allow_blank=False)
    ws.add_data_validation(dv); dv.add("B4")

    ft = 22  # 시설물 표 시작 행 (헤더)
    for i, (code, label) in enumerate(VARS):
        r = 6 + i
        ws.cell(r, 1).value = f"{label} {code}"
        flag_col = 6 + i  # 시설물 표의 필요변수 열 (F~L)
        ws.cell(r, 3).value = (f'=IF(INDEX({ws.cell(ft + 1, flag_col).coordinate}:{ws.cell(ft + 6, flag_col).coordinate},'
                               f'MATCH($B$4,$A${ft + 1}:$A${ft + 6},0))=1,"← 필요","")')
    for r in list(range(4, 6)) + list(range(6, 13)):
        ws.cell(r, 1).font = BOLD
        ws.cell(r, 2).fill, ws.cell(r, 2).border = INPUT, BOX
    ws["A14"], ws["B14"] = "준비기간 (일)", 60
    ws["A15"], ws["B15"] = "정리기간 (일)", 30
    for r in (14, 15):
        ws.cell(r, 1).font = BOLD
        ws.cell(r, 2).fill, ws.cell(r, 2).border = INPUT, BOX

    look = lambda col: f"INDEX(${col}${ft + 1}:${col}${ft + 6},MATCH($B$4,$A${ft + 1}:$A${ft + 6},0))"
    ws["A17"] = "적용범위 판정"
    ws["B17"] = (f'=IF(NOT(ISNUMBER(B5)),"총공사비를 입력하세요",IF(AND(B5>={look("C")},B5<={look("D")}),'
                 f'"적용범위 이내 (산정공식 사용 가능)","적용범위 초과 → 산정공식 사용 불가. 1일 작업량 방식으로 산정하세요"))')
    ws["A18"] = "산정공식 값 Y (일)"
    y = ('IFERROR(ROUND(IF($B$4="도로포장",-637.009+173.198*LN(B6)+0.049*B5,'
         'IF($B$4="도로(토공+교량)",-160.855-14.288*B7+164.473*LN(B6)-1.474*B8+0.052*B5,'
         'IF($B$4="농업용수",-2251.569+415.137*LN(B5),'
         'IF($B$4="상수도",-1175.174+119.731*B10-0.273*B9+222.426*LN(B5),'
         'IF($B$4="하수도",-452.433+98.364*LN(B11)+0.083*B5,'
         '-1723.316-74.26*LN(B12)+372.266*LN(B5)))))),0),"입력 확인")')
    ws["B18"] = f'=IF(LEFT(B17,4)="적용범위",IF(ISNUMBER(SEARCH("이내",B17)),{y},"적용불가"),"")'
    ws["A19"] = "공사기간 (Y + 준비 + 정리, 일)"
    ws["B19"] = '=IF(ISNUMBER(B18),B18+B14+B15,"-")'
    ws["A20"] = "환산 (개월)"
    ws["B20"] = '=IF(ISNUMBER(B19),ROUND(B19/30,1),"-")'
    for r in (17, 18, 19, 20):
        ws.cell(r, 1).font = BOLD
        ws.cell(r, 2).fill, ws.cell(r, 2).border = RESULT, BOX

    ws.cell(ft - 1, 1).value = "▼ 시설물별 산정공식과 적용범위 (부록 5)  Y=공사기간(일), 금액 단위는 백만원(농업용수·상수도 등 범위는 억원을 백만원으로 환산)"
    heads = ["시설물", "산정공식", "총공사비 하한(백만원)", "총공사비 상한(백만원)", "",
             *[c for c, _ in VARS]]
    for j, h in enumerate(heads):
        ws.cell(ft, 1 + j).value = h
    _hdr(ws, ft, range(1, 13))
    for i, (name, formula, lo, hi, flags) in enumerate(FACILITIES):
        r = ft + 1 + i
        ws.cell(r, 1).value, ws.cell(r, 2).value = name, formula
        ws.cell(r, 3).value, ws.cell(r, 4).value = lo, hi
        for k, f in enumerate(flags):
            ws.cell(r, 6 + k).value = f
        for c in range(1, 13):
            ws.cell(r, c).border = BOX
    ws.cell(ft + 8, 1).value = "※ ln은 자연로그. 상수도는 다종의 관일 때 물량이 가장 많거나 가장 큰 관경을 적용. 도로포장은 토공+교량이 함께 있으면 '도로(토공+교량)' 공식을 적용. 철도(궤도)는 설비 부분 공기 별도."
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 62
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    return ws


def add_workday_sheet(wb, n_data, n_list, n_hol):
    ws = wb.create_sheet("공기산정(1일작업량)")
    ws["A1"] = "1일 작업량에 의한 공사기간 산정 (가이드라인 제2장)"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("공사기간 = 준비기간 + 주공정(CP) 공종별 공사기간의 합 + 정리기간,  공종별 공사기간 = 작업일수 + 비작업일수.  "
                "월별 비작업일수 = A(기후)+B(공휴일)-C(A×B÷달력일수), 월 8일(주40시간) 미만이면 8일 적용.  "
                "공종은 입력 순서대로 이어서 시공하는 것으로 계산합니다. 준비·정리기간에는 비작업일수를 넣지 않습니다.")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:L2")
    ws.row_dimensions[2].height = 48

    labels = [("A4", "기상 지점", "B4", "서울"), ("A5", "본공사 착수 연도", "B5", 2026),
              ("A6", "본공사 착수 월", "B6", 7), ("A7", "준비기간 (일)", "B7", "='시설물별공기(토목)'!B14"),
              ("A8", "정리기간 (일)", "B8", "='시설물별공기(토목)'!B15")]
    for a, lab, b, val in labels:
        ws[a], ws[b] = lab, val
        ws[a].font = BOLD
        ws[b].fill, ws[b].border = INPUT, BOX
    ws["C4"] = f'=IFERROR("지점코드 "&INDEX(지점목록!$C$2:$C${n_list},MATCH(B4,지점목록!$A$2:$A${n_list},0)),"지점명을 확인하세요")'
    ws["C5"] = "※ 본공사 착수 월의 1일부터 계산 (준비기간은 그 이전)"
    dv = DataValidation(type="list", formula1=f"=지점목록!$A$2:$A${n_list}", allow_blank=False)
    ws.add_data_validation(dv); dv.add("B4")

    # ── 기상조건 선택표 (행 12~24)
    c0 = 11
    ws.cell(c0 - 1, 1).value = "▼ 적용할 기상조건 선택 (O=적용, X=미적용) — 선택한 지점의 월별 비작업일수(2015~2024, 가이드라인 부록 3)"
    ws.cell(c0 - 1, 1).font = BOLD
    for j, h in enumerate(["번호", "기상조건", "적용", *[f"{m}월" for m in range(1, 13)], "소계"]):
        ws.cell(c0, 1 + j).value = h
    _hdr(ws, c0, range(1, 16))
    dv2 = DataValidation(type="list", formula1='"O,X"', allow_blank=False)
    ws.add_data_validation(dv2)
    names = json.load(open("data/weather.json", encoding="utf-8"))
    for i, c in enumerate(names, 1):
        r = c0 + i
        ws.cell(r, 1).value, ws.cell(r, 2).value = i, c["name"]
        ws.cell(r, 3).value = "O" if i in COND_DEFAULT_ON else "X"
        ws.cell(r, 3).fill = INPUT
        dv2.add(ws.cell(r, 3))
        for m in range(12):
            col = "FGHIJKLMNOPQ"[m]
            ws.cell(r, 4 + m).value = (f"=SUMIFS(기상자료!${col}$2:${col}${n_data},기상자료!$A$2:$A${n_data},"
                                       f"INDEX(지점목록!$C$2:$C${n_list},MATCH($B$4,지점목록!$A$2:$A${n_list},0)),"
                                       f"기상자료!$D$2:$D${n_data},$A{r})")
        ws.cell(r, 16).value = f"=SUM(D{r}:O{r})"
        for cc in range(1, 17):
            ws.cell(r, cc).border = BOX
    sr = c0 + len(names) + 1
    ws.cell(sr, 2).value = "적용 조건 합계 (A: 기후여건)"
    ws.cell(sr, 2).font = BOLD
    for m in range(12):
        col = ws.cell(sr, 4 + m).column_letter
        ws.cell(sr, 4 + m).value = f'=SUMPRODUCT(($C${c0 + 1}:$C${c0 + 13}="O")*{col}{c0 + 1}:{col}{c0 + 13})'
        ws.cell(sr, 4 + m).fill = RESULT
    ws.cell(sr, 16).value = f"=SUM(D{sr}:O{sr})"
    A_ROW = sr

    # ── 월별 표 (60개월)
    mt = sr + 12 + 20  # 월별 표 헤더 행 (공종표 뒤)
    first, last = mt + 1, mt + 60

    # ── 공종 입력표
    wt = sr + 3
    ws.cell(wt - 1, 1).value = "▼ 공종별 작업량 입력 (주공정=Y인 공종만 공사기간에 합산, 위에서 아래 순서로 이어서 시공)"
    ws.cell(wt - 1, 1).font = BOLD
    heads = ["공종", "작업수량", "단위", "1일 작업량", "주공정(Y/N)", "작업일수", "누적 작업일수", "경과일수(누적)", "공종별 공사기간", "비작업일수"]
    for j, h in enumerate(heads):
        ws.cell(wt, 1 + j).value = h
    _hdr(ws, wt, range(1, 11))
    dv3 = DataValidation(type="list", formula1='"Y,N"', allow_blank=False)
    ws.add_data_validation(dv3)
    K, L, D_, J_ = (f"$K${first}:$K${last}", f"$L${first}:$L${last}", f"$D${first}:$D${last}", f"$J${first}:$J${last}")
    n_rows = 15
    demo = [("철골세우기", 4000, "톤", 50), ("데크플레이트 설치", 33800, "㎡", 440), ("철근 가공/조립", 940, "톤", 12),
            ("형틀 조립", 24500, "㎡", 320), ("콘크리트 타설", 10300, "㎥", 930)]
    for i in range(n_rows):
        r = wt + 1 + i
        d = demo[i] if i < len(demo) else (None, None, None, None)
        for j, v in enumerate(d):
            ws.cell(r, 1 + j).value = v
        ws.cell(r, 5).value = "Y" if d[0] else None
        dv3.add(ws.cell(r, 5))
        for c in (1, 2, 3, 4, 5):
            ws.cell(r, c).fill = INPUT
        ws.cell(r, 6).value = f'=IF(AND(ISNUMBER(B{r}),ISNUMBER(D{r}),N(D{r})>0),ROUNDUP(B{r}/D{r},0),"")'
        ws.cell(r, 7).value = f'=IF(F{r}="","",N(G{r - 1})+IF(E{r}="Y",F{r},0))'
        ws.cell(r, 8).value = (
            f'=IF(G{r}="","",IF(G{r}=0,0,IF(G{r}>INDEX({K},60)+INDEX({J_},60),"기간초과",'
            f'INDEX({L},MATCH(G{r},{K},1))+(G{r}-INDEX({K},MATCH(G{r},{K},1)))'
            f'*INDEX({D_},MATCH(G{r},{K},1))/INDEX({J_},MATCH(G{r},{K},1)))))')
        ws.cell(r, 9).value = f'=IF(F{r}="","",IF(E{r}="Y",IF(ISNUMBER(H{r}),ROUND(H{r}-N(H{r - 1}),1),"기간초과"),"-"))'
        ws.cell(r, 10).value = f'=IF(ISNUMBER(I{r}),ROUND(I{r}-F{r},1),"")'
        for c in range(1, 11):
            ws.cell(r, c).border = BOX
    t1, t2 = wt + 1, wt + n_rows
    res = t2 + 2
    ws.cell(res, 1).value = "▼ 결과"
    ws.cell(res, 1).font = BOLD
    out = [
        ("준비기간 (일)", "=B7"),
        ("주공정 공종별 공사기간의 합 (일)", f'=IFERROR(ROUNDUP(MAX(H{t1}:H{t2}),0),"기간초과")'),
        ("정리기간 (일)", "=B8"),
    ]
    for i, (lab, f) in enumerate(out):
        ws.cell(res + 1 + i, 1).value, ws.cell(res + 1 + i, 2).value = lab, f
    ws.cell(res + 4, 1).value = "공사기간 A (일)"
    ws.cell(res + 4, 2).value = f'=IF(ISNUMBER(B{res + 2}),B{res + 1}+B{res + 2}+B{res + 3},"-")'
    ws.cell(res + 5, 1).value = "환산 (개월)"
    ws.cell(res + 5, 2).value = f'=IF(ISNUMBER(B{res + 4}),ROUND(B{res + 4}/30,1),"-")'
    ws.cell(res + 6, 1).value = "산정공식 공사기간 B (일) — 시설물별공기(토목)"
    ws.cell(res + 6, 2).value = "='시설물별공기(토목)'!B19"
    ws.cell(res + 7, 1).value = "적정성 검토 (A가 B의 ±20% 이내인지)"
    ws.cell(res + 7, 2).value = (f'=IF(AND(ISNUMBER(B{res + 4}),ISNUMBER(B{res + 6})),'
                                 f'IF(ABS(B{res + 4}/B{res + 6}-1)<=0.2,"±20% 이내 (적정)","±20% 초과 → 산정 과정 재검토 (편차 "&TEXT(B{res + 4}/B{res + 6}-1,"0.0%")&")"),'
                                 f'"산정공식 적용범위 밖이거나 미입력 → 비교 생략")')
    for r in range(res + 1, res + 8):
        ws.cell(r, 1).font = BOLD
        ws.cell(r, 2).fill, ws.cell(r, 2).border = RESULT, BOX
    assert res + 8 < mt, (res, mt)

    # ── 월별 비작업일수 표
    ws.cell(mt - 1, 1).value = "▼ 월별 비작업일수 계산 (착수 월부터 60개월)"
    ws.cell(mt - 1, 1).font = BOLD
    heads = ["순번", "연도", "월", "달력일수", "법정공휴일 B", "기후여건 A", "중복일수 C=A×B÷달력일수",
             "A+B-C", "적용 비작업일수(월8일 이상, 정수)", "작업가능일수", "누적 작업가능(전월까지)", "누적 달력일수(전월까지)"]
    for j, h in enumerate(heads):
        ws.cell(mt, 1 + j).value = h
    _hdr(ws, mt, range(1, 13))
    ws.row_dimensions[mt].height = 48
    for i in range(60):
        r = first + i
        ws.cell(r, 1).value = i + 1
        ws.cell(r, 2).value = f"=YEAR(EDATE(DATE($B$5,$B$6,1),A{r}-1))"
        ws.cell(r, 3).value = f"=MONTH(EDATE(DATE($B$5,$B$6,1),A{r}-1))"
        ws.cell(r, 4).value = f"=DAY(EOMONTH(DATE(B{r},C{r},1),0))"
        ws.cell(r, 5).value = f"=IFERROR(INDEX(법정공휴일!$B$4:$M${n_hol},MATCH(B{r},법정공휴일!$A$4:$A${n_hol},0),C{r}),0)"
        ws.cell(r, 6).value = f"=INDEX($D${A_ROW}:$O${A_ROW},C{r})"
        ws.cell(r, 7).value = f"=ROUND(F{r}*E{r}/D{r},1)"
        ws.cell(r, 8).value = f"=F{r}+E{r}-G{r}"
        ws.cell(r, 9).value = f"=ROUND(MAX(H{r},8),0)"
        ws.cell(r, 10).value = f"=D{r}-I{r}"
        ws.cell(r, 11).value = 0 if i == 0 else f"=K{r - 1}+J{r - 1}"
        ws.cell(r, 12).value = 0 if i == 0 else f"=L{r - 1}+D{r - 1}"
        for c in range(1, 13):
            ws.cell(r, c).border = BOX
    ws.cell(last + 2, 1).value = ("※ 마지막 달은 '비작업일수 = 총 비작업일수 × 잔여작업일수 ÷ 그 달 총 작업가능일수' 규칙(가이드라인 1150행)으로 안분합니다. "
                                  "가이드라인 예시(철골세우기)는 20+23+12+10+10을 74일로 적었으나 실제 합은 75일입니다.")
    ws.column_dimensions["A"].width = 36
    ws.column_dimensions["B"].width = 34
    for col in "CDEFGHIJKL":
        ws.column_dimensions[col].width = 14
    return ws
