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
def openpyxl_col(n):
    from openpyxl.utils import get_column_letter
    return get_column_letter(n)


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


def add_workrate_sheet(wb):
    import workrates_civil
    ws = wb.create_sheet("1일작업량(토목)")
    ws.append(["선택명 (분류>공종 (작업조건))", "수량 단위", "1일 작업량 (수량/일)", "원문 표기 (부록 4)"])
    _hdr(ws, 1, range(1, 5))
    rows = workrates_civil.build()
    for key, unit, rate, src in rows:
        ws.append([key, unit, rate, src])
    for r in range(2, len(rows) + 2):
        ws.cell(r, 3).number_format = "#,##0.0000"
    ws.cell(1, 5).value = "기본 기상세트"
    _hdr(ws, 1, [5])
    default_set = {"공통가설공사": "세트2", "토공사": "세트2", "배수공사": "세트1", "포장공사": "세트3",
                   "교량공사": "세트1", "터널공사": "세트4", "부대공사": "세트2"}
    for r, (key, *_r) in enumerate(rows, 2):
        ws.cell(r, 5).value = default_set[key.split(">")[0]]
    ws.freeze_panes = "A2"
    ws.column_dimensions["A"].width = 70
    ws.column_dimensions["C"].width = 20
    ws.column_dimensions["D"].width = 22
    cats, first = [], {}
    for i, (key, *_r) in enumerate(rows):
        c = key.split(">")[0]
        if c not in first:
            first[c] = [i, 0]
            cats.append(c)
        first[c][1] += 1
    for j, h in enumerate(["분류 필터", "시작 위치", "개수"]):
        ws.cell(1, 6 + j).value = h
    _hdr(ws, 1, range(6, 9))
    ws.cell(2, 6).value, ws.cell(2, 7).value, ws.cell(2, 8).value = "전체", 1, len(rows)
    for j, c in enumerate(cats):
        ws.cell(3 + j, 6).value, ws.cell(3 + j, 7).value, ws.cell(3 + j, 8).value = c, first[c][0] + 1, first[c][1]
        assert [k.split(">")[0] for k, *_r in rows[first[c][0]:first[c][0] + first[c][1]]] == [c] * first[c][1], c
    ws.column_dimensions["F"].width = 16
    ncat = len(cats) + 1
    ws.cell(1, 10).value = "필터된 공종 목록 (자동, 공종 칸 드롭다운의 원본)"
    _hdr(ws, 1, [10])
    ws.column_dimensions["J"].width = 70
    sel = "MATCH('공기산정(1일작업량)'!$B$32,$F$2:$F$%d,0)" % (ncat + 1)
    for r in range(2, len(rows) + 2):
        k = f"ROWS($J$2:J{r})"
        ws.cell(r, 10).value = (f'=IF({k}<=INDEX($H$2:$H${ncat + 1},{sel}),'
                                f'INDEX($A$2:$A${len(rows) + 1},INDEX($G$2:$G${ncat + 1},{sel})+{k}-1),"")')
    ws.cell(len(rows) + 3, 1).value = "※ 가이드라인 부록 4 '토목분야 > 도로시설물'을 옮긴 것입니다. '일/개소' 형태는 1일 작업량을 1÷일수로 환산했습니다. 표에 없는 공종은 공종표에 직접 입력하세요."
    return len(rows) + 1, len(cats) + 1


def add_preset_sheet(wb):
    import presets_civil
    ws = wb.create_sheet("공종프리셋")
    ws["A18"] = "※ 공사 유형별 대표 공종(작업 순서)입니다. '공기산정(1일작업량)' 시트의 공사 유형을 고르면 공종표가 이 목록으로 채워집니다. 여기서 항목을 고쳐도 됩니다."
    for j, (name, items) in enumerate(presets_civil.PRESETS.items(), 1):
        ws.cell(1, j).value = name
        ws.cell(1, j).fill, ws.cell(1, j).font = HEAD, BOLD
        for i, it in enumerate(items):
            ws.cell(2 + i, j).value = it
        ws.column_dimensions[ws.cell(1, j).column_letter].width = 46
    return len(presets_civil.PRESETS)


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

    ft = 24  # 시설물 표 시작 행 (헤더)
    for i, (code, label) in enumerate(VARS):
        r = 6 + i
        ws.cell(r, 1).value = f"{label} {code}"
        flag_col = 6 + i  # 시설물 표의 필요변수 열 (F~L)
        ws.cell(r, 3).value = (f'=IF(INDEX({ws.cell(ft + 1, flag_col).coordinate}:{ws.cell(ft + 6, flag_col).coordinate},'
                               f'MATCH($B$4,$A${ft + 1}:$A${ft + 6},0))=1,"← 필요","")')
    for r in list(range(4, 6)) + list(range(6, 13)):
        ws.cell(r, 1).font = BOLD
        ws.cell(r, 2).fill, ws.cell(r, 2).border = INPUT, BOX
    ws["A14"], ws["B14"] = "준비기간 (일, 직접 입력)", 0
    ws["A15"], ws["B15"] = "정리기간 (일, 직접 입력)", 0
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
    ws["H4"] = "▼ 공사 유형별 준비기간(예시) — 가이드라인"
    ws["H4"].font = BOLD
    for j, h in enumerate(["공종", "준비기간", "공종", "준비기간"]):
        ws.cell(5, 8 + j).value = h
    _hdr(ws, 5, range(8, 12))
    for i, row in enumerate(PREP_EXAMPLES):
        for j, v in enumerate(row):
            ws.cell(6 + i, 8 + j).value = v
            ws.cell(6 + i, 8 + j).border = BOX
    for col, w in zip("HIJK", (18, 12, 18, 12)):
        ws.column_dimensions[col].width = w
    ws["A21"] = "※ 산정공식은 과거 공사 실적의 회귀식(평균값)이라, 소규모 공사는 실제 공기보다 크게 나올 수 있습니다. 이 값은 적정성 검토용이며 최종 공기는 1일 작업량 방식으로 산정하세요."
    ws["A21"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A21:F21")
    ws.row_dimensions[21].height = 32
    ws.cell(ft + 8, 1).value = "※ ln은 자연로그. 상수도는 다종의 관일 때 물량이 가장 많거나 가장 큰 관경을 적용. 도로포장은 토공+교량이 함께 있으면 '도로(토공+교량)' 공식을 적용. 철도(궤도)는 설비 부분 공기 별도."
    ws.column_dimensions["A"].width = 30
    ws.column_dimensions["B"].width = 62
    ws.column_dimensions["C"].width = 22
    ws.column_dimensions["D"].width = 22
    return ws


PREP_EXAMPLES = [  # 가이드라인 제2장 '건설공사 유형별 준비기간(예시)' (663~671행)
    ("공동주택", "45일", "상수도공사", "60일"),
    ("고속도로공사", "180일", "하천공사", "40일"),
    ("철도공사", "90일", "항만공사", "40일"),
    ("포장공사(신설)", "50일", "강교가설공사", "90일"),
    ("포장공사(수선)", "60일", "PC교량 공사", "70일"),
    ("공동구공사", "80일", "교량보수공사", "60일"),
]

SET_DEFAULTS = [
    ("콘크리트·구조물", {1, 5, 6, 9, 13}),
    ("토공·가설(옥외)", {1, 5, 6, 10, 13}),
    ("포장(아스팔트)", {5, 6, 9}),
    ("옥내·터널", {1, 5}),
]


def add_workday_sheet(wb, n_data, n_list, n_hol, n_rate, n_types, n_cat):
    from openpyxl.utils import get_column_letter as L
    import json as _json
    ws = wb.create_sheet("공기산정(1일작업량)")
    ws["A1"] = "1일 작업량에 의한 공사기간 산정 (가이드라인 제2장)"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("공사기간 = 준비기간 + 주공정(CP) 공종별 공사기간의 합 + 정리기간,  공종별 공사기간 = 작업일수 + 비작업일수.  "
                "월별 비작업일수 = A(기후)+B(공휴일)-C(A×B÷달력일수), 월 8일(주40시간) 미만이면 8일 적용.  "
                "공종마다 기상조건 세트를 골라 그 세트의 조건으로 비작업일수를 계산합니다. 공종은 입력 순서대로 이어서 시공하는 것으로 가정합니다. "
                "준비·정리기간에는 비작업일수를 넣지 않습니다.")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:L2")
    ws.row_dimensions[2].height = 60

    for a, lab, b, val in [("A4", "기상 지점", "B4", "제주"), ("A5", "본공사 착수 연도", "B5", 2026),
                           ("A6", "본공사 착수 월", "B6", 7), ("A7", "준비기간 (일)", "B7", "='시설물별공기(토목)'!B14"),
                           ("A8", "정리기간 (일)", "B8", "='시설물별공기(토목)'!B15")]:
        ws[a], ws[b] = lab, val
        ws[a].font = BOLD
        ws[b].fill, ws[b].border = INPUT, BOX
    ws["A9"], ws["B9"] = "공사 유형 (공종 자동 구성)", "하수도 관로공사"
    ws["A9"].font = BOLD
    ws["B9"].fill, ws["B9"].border = INPUT, BOX
    last_col = openpyxl_col(n_types)
    dvt = DataValidation(type="list", formula1=f"=공종프리셋!$A$1:${last_col}$1", allow_blank=False)
    ws.add_data_validation(dvt); dvt.add("B9")
    ws["C9"] = "※ 유형을 바꾸면 아래 공종표가 그 유형의 대표 공종으로 바뀝니다. 수량은 다시 입력하세요."
    code = f"INDEX(지점목록!$C$2:$C${n_list},MATCH($B$4,지점목록!$A$2:$A${n_list},0))"
    ws["C4"] = f'=IFERROR("지점코드 "&{code},"지점명을 확인하세요")'
    ws["C5"] = "※ 본공사 착수 월의 1일부터 계산 (준비기간은 그 이전)"
    dv = DataValidation(type="list", formula1=f"=지점목록!$A$2:$A${n_list}", allow_blank=False)
    ws.add_data_validation(dv); dv.add("B4")

    # ── 기상조건 세트 (행 11~30)
    ws["A11"] = "▼ 기상조건 세트 — 공종마다 다른 조건을 쓰도록 세트 4개를 정의합니다 (O=적용, X=미적용). 값은 선택한 지점의 월별 비작업일수(2015~2024, 부록 3)"
    ws["A11"].font = BOLD
    ws["A12"], ws["B12"] = "세트 이름 (수정 가능)", "※ 세트별 이름을 적어 두세요"
    ws["A12"].font = BOLD
    for p, (nm, _c) in enumerate(SET_DEFAULTS):
        ws.cell(12, 3 + p).value = nm
        ws.cell(12, 3 + p).fill, ws.cell(12, 3 + p).border = INPUT, BOX
        ws.cell(12, 3 + p).alignment = Alignment(wrap_text=True, horizontal="center")
    ws.row_dimensions[12].height = 32
    c0 = 13
    heads = ["번호", "기상조건", "세트1", "세트2", "세트3", "세트4", *[f"{m}월" for m in range(1, 13)], "소계"]
    for j, h in enumerate(heads):
        ws.cell(c0, 1 + j).value = h
    _hdr(ws, c0, range(1, 20))
    dv2 = DataValidation(type="list", formula1='"O,X"', allow_blank=False)
    ws.add_data_validation(dv2)
    names = _json.load(open("data/weather.json", encoding="utf-8"))
    for i, c in enumerate(names, 1):
        r = c0 + i
        ws.cell(r, 1).value, ws.cell(r, 2).value = i, c["name"]
        for p, (_nm, on) in enumerate(SET_DEFAULTS):
            cell = ws.cell(r, 3 + p)
            cell.value = "O" if i in on else "X"
            cell.fill = INPUT
            cell.alignment = Alignment(horizontal="center")
            dv2.add(cell)
        for m in range(12):
            col = "FGHIJKLMNOPQ"[m]
            ws.cell(r, 7 + m).value = (f"=SUMIFS(기상자료!${col}$2:${col}${n_data},기상자료!$A$2:$A${n_data},{code},"
                                       f"기상자료!$D$2:$D${n_data},$A{r})")
        ws.cell(r, 19).value = f"=SUM(G{r}:R{r})"
        for cc in range(1, 20):
            ws.cell(r, cc).border = BOX
    f1, f2 = c0 + 1, c0 + 13
    sum_row = {}
    for p in range(4):
        r = f2 + 1 + p
        sum_row[p] = r
        ws.cell(r, 2).value = f"세트{p + 1} 기후여건 A 합계"
        ws.cell(r, 2).font = BOLD
        fl = "CDEF"[p]
        for m in range(12):
            cl = L(7 + m)
            ws.cell(r, 7 + m).value = f'=SUMPRODUCT(({fl}${f1}:{fl}${f2}="O")*{cl}${f1}:{cl}${f2})'
            ws.cell(r, 7 + m).fill = RESULT
        ws.cell(r, 19).value = f"=SUM(G{r}:R{r})"
    last_sum = f2 + 4

    fr = last_sum + 2  # 분류 필터 행
    ws.cell(fr, 1).value, ws.cell(fr, 2).value = "공종 목록 분류 (선택 범위 좁히기)", "전체"
    ws.cell(fr, 1).font = BOLD
    ws.cell(fr, 2).fill, ws.cell(fr, 2).border = INPUT, BOX
    dvc = DataValidation(type="list", formula1=f"='1일작업량(토목)'!$F$2:$F${n_cat + 1}", allow_blank=False)
    ws.add_data_validation(dvc); dvc.add(f"B{fr}")
    ws.cell(fr, 3).value = "※ 분류를 고르면 아래 공종 칸의 목록이 그 분류의 공종만 보입니다."
    # 필터된 목록의 기준 셀은 add_workrate_sheet에서 B{fr}을 참조하도록 고정: 여기서는 위치 검증만 한다
    assert fr == 32, fr

    wt = fr + 2
    ws.cell(wt - 1, 1).value = "▼ 공종별 작업량 입력 (주공정=Y인 공종만 공사기간에 합산, 위에서 아래 순서로 이어서 시공)"
    ws.cell(wt - 1, 1).font = BOLD
    ws.cell(wt - 1, 6).value = "※ 공종 삭제=칸 지우기, 추가=빈 줄에서 목록 선택. 1일 작업량·기상세트는 덮어써서 수정 가능."
    heads = ["공종 (목록 선택 또는 직접 입력)", "작업수량", "단위", "1일 작업량 (수정 가능)", "주공정(Y/N)", "기상조건 세트",
             "작업일수", "시작(경과일)", "종료(경과일)", "공종별 공사기간", "비작업일수", "부록4 원문 표기",
             "세트번호", "시작월 순번", "시작시 누적작업가능", "종료월 순번"]
    for j, h in enumerate(heads):
        ws.cell(wt, 1 + j).value = h
    _hdr(ws, wt, range(1, 17))
    for c in range(13, 17):
        ws.cell(wt, c).fill = PatternFill("solid", fgColor="EDEDED")
    dv3 = DataValidation(type="list", formula1='"Y,N"', allow_blank=False)
    ws.add_data_validation(dv3)
    dvs = DataValidation(type="list", formula1=f"=$C${c0}:$F${c0}", allow_blank=True)
    ws.add_data_validation(dvs)
    dv4 = DataValidation(type="list", formula1=f"='1일작업량(토목)'!$J$2:$J${n_rate}", allow_blank=True)
    dv4.showErrorMessage = False
    ws.add_data_validation(dv4)

    n_rows = 15
    t1, t2 = wt + 1, wt + n_rows
    res = t2 + 2
    mt = res + 11
    first, last = mt + 1, mt + 60
    # 월별 표 참조 범위
    rng = lambda col: f"${L(col)}${first}:${L(col)}${last}"
    FC = rng(6)          # 누적 달력일수(전월까지)
    CAL = rng(4)         # 달력일수
    Kc = [rng(7 + 5 * p + 4) for p in range(4)]   # 누적 작업가능(전월까지)
    Vc = [rng(7 + 5 * p + 3) for p in range(4)]   # 작업가능일수
    choose = lambda lst: "CHOOSE(M{r}," + ",".join(lst) + ")"
    wk = "'1일작업량(토목)'"
    for i in range(n_rows):
        r = t1 + i
        ch = lambda lst: "CHOOSE($M%d,%s)" % (r, ",".join(lst))
        ws.cell(r, 1).value = (f"=IFERROR(INDEX(공종프리셋!$A$2:${last_col}$16,{i + 1},"
                               f"MATCH($B$9,공종프리셋!$A$1:${last_col}$1,0))&\"\",\"\")")
        dv4.add(ws.cell(r, 1))
        m = f"MATCH($A{r},{wk}!$A$2:$A${n_rate},0)"
        ws.cell(r, 3).value = f"=IFERROR(INDEX({wk}!$B$2:$B${n_rate},{m}),\"\")"
        ws.cell(r, 4).value = f"=IFERROR(INDEX({wk}!$C$2:$C${n_rate},{m}),\"\")"
        ws.cell(r, 5).value = f'=IF(A{r}="","","Y")'
        dv3.add(ws.cell(r, 5))
        ws.cell(r, 6).value = (f'=IF(A{r}="","",IFERROR(INDEX({wk}!$E$2:$E${n_rate},{m}),'
                               f'IF(ISNUMBER(SEARCH("포장",A{r})),"세트3","세트1")))')
        dvs.add(ws.cell(r, 6))
        for c in (1, 2, 3, 4, 5, 6):
            ws.cell(r, c).fill = INPUT
        ws.cell(r, 7).value = f'=IF(AND(ISNUMBER(B{r}),ISNUMBER(D{r}),N(D{r})>0),ROUNDUP(ROUND(B{r}/D{r},6),0),"")'
        ws.cell(r, 13).value = f'=IF(G{r}="","",IFERROR(MATCH(F{r},$C${c0}:$F${c0},0),1))'
        ws.cell(r, 8).value = f'=IF(G{r}="","",0)' if i == 0 else f'=IF(G{r}="","",MAX(0,MAX(I${t1}:I{r - 1})))'
        ws.cell(r, 14).value = f'=IF(G{r}="","",MATCH(H{r},{FC},1))'
        ws.cell(r, 15).value = (f'=IF(G{r}="","",INDEX({ch(Kc)},N{r})+(H{r}-INDEX({FC},N{r}))'
                                f'*INDEX({ch(Vc)},N{r})/INDEX({CAL},N{r}))')
        ws.cell(r, 16).value = (f'=IF(G{r}="","",IF(O{r}+G{r}>INDEX({ch(Kc)},60)+INDEX({ch(Vc)},60),"기간초과",'
                                f'MATCH(O{r}+G{r},{ch(Kc)},1)))')
        ws.cell(r, 9).value = (f'=IF(G{r}="","",IF(E{r}<>"Y",H{r},IF(P{r}="기간초과","기간초과",'
                               f'INDEX({FC},P{r})+(O{r}+G{r}-INDEX({ch(Kc)},P{r}))*INDEX({CAL},P{r})/INDEX({ch(Vc)},P{r}))))')
        ws.cell(r, 10).value = f'=IF(G{r}="","",IF(E{r}="Y",IF(ISNUMBER(I{r}),ROUND(I{r}-H{r},1),"기간초과"),"-"))'
        ws.cell(r, 11).value = f'=IF(ISNUMBER(J{r}),ROUND(J{r}-G{r},1),"")'
        ws.cell(r, 12).value = f"=IFERROR(INDEX({wk}!$D$2:$D${n_rate},{m}),\"\")"
        for c in range(1, 17):
            ws.cell(r, c).border = BOX
        for c in range(13, 17):
            ws.cell(r, c).font = Font(color="808080")

    ws.cell(res, 1).value = "▼ 결과"
    ws.cell(res, 1).font = BOLD
    out = [
        ("준비기간 (일)", "=B7"),
        ("주공정 공종별 공사기간의 합 (일)",
         f'=IF(COUNTIF(I{t1}:I{t2},"기간초과")>0,"기간초과",ROUNDUP(MAX(0,MAX(I{t1}:I{t2})),0))'),
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
    assert res + 8 < mt - 1, (res, mt)

    # ── 월별 비작업일수 표 (공통 + 세트 4개)
    ws.cell(mt - 1, 1).value = "▼ 월별 비작업일수 계산 (착수 월부터 60개월) — 세트별로 기후여건 A가 다릅니다"
    ws.cell(mt - 1, 1).font = BOLD
    base_heads = ["순번", "연도", "월", "달력일수", "법정공휴일 B", "누적 달력일수(전월까지)"]
    blk = ["기후여건 A", "중복일수 C", "적용 비작업일수(월8일 이상, 정수)", "작업가능일수", "누적 작업가능(전월까지)"]
    for j, h in enumerate(base_heads):
        ws.cell(mt, 1 + j).value = h
    for p in range(4):
        for j, h in enumerate(blk):
            ws.cell(mt, 7 + 5 * p + j).value = f"세트{p + 1} {h}"
    _hdr(ws, mt, range(1, 27))
    ws.row_dimensions[mt].height = 62
    for i in range(60):
        r = first + i
        ws.cell(r, 1).value = i + 1
        ws.cell(r, 2).value = f"=YEAR(EDATE(DATE($B$5,$B$6,1),A{r}-1))"
        ws.cell(r, 3).value = f"=MONTH(EDATE(DATE($B$5,$B$6,1),A{r}-1))"
        ws.cell(r, 4).value = f"=DAY(EOMONTH(DATE(B{r},C{r},1),0))"
        ws.cell(r, 5).value = f"=IFERROR(INDEX(법정공휴일!$B$4:$M${n_hol},MATCH(B{r},법정공휴일!$A$4:$A${n_hol},0),C{r}),0)"
        ws.cell(r, 6).value = 0 if i == 0 else f"=F{r - 1}+D{r - 1}"
        for p in range(4):
            b = 7 + 5 * p
            a_, c_, adj, av, k = (L(b + j) for j in range(5))
            ws.cell(r, b).value = f"=INDEX($G${sum_row[p]}:$R${sum_row[p]},C{r})"
            ws.cell(r, b + 1).value = f"=ROUND({a_}{r}*E{r}/D{r},1)"
            ws.cell(r, b + 2).value = f"=ROUND(MAX({a_}{r}+E{r}-{c_}{r},8),0)"
            ws.cell(r, b + 3).value = f"=D{r}-{adj}{r}"
            ws.cell(r, b + 4).value = 0 if i == 0 else f"={k}{r - 1}+{av}{r - 1}"
        for c in range(1, 27):
            ws.cell(r, c).border = BOX
    ws.cell(last + 2, 1).value = ("※ 각 공종은 앞 공종이 끝난 시점에서 시작하며, 그 공종의 기상세트 표로 작업가능일수를 채워 종료 시점을 찾습니다. "
                                  "한 달 안에서는 작업가능일수가 고르게 분포한다고 보고 안분합니다(가이드라인 1150행 규칙과 같은 방식).")
    ws.column_dimensions["A"].width = 60
    ws.column_dimensions["B"].width = 22
    for col in range(3, 27):
        ws.column_dimensions[L(col)].width = 13
    ws.column_dimensions["L"].width = 18
    return ws, dict(sum_row=sum_row, t1=t1, t2=t2, c0=c0)


# (구분, 작업, 동절기·저온, 강우, 바람, 눈, 혹서기·고온, 출처, 부록3에서 고를 조건 번호)
LIMIT_ROWS = [
    ("가설·장비", "이동식 크레인", "", "", "평균풍속 10 m/s 초과", "", "", "KCS 21 20 10", "13 (순간풍속 15m/s로 근사)"),
    ("가설·장비", "타워크레인 설치·인상·해체·점검·수리", "", "1 mm/hr 이상", "순간풍속 10 m/s 이상", "10 mm/hr 이상", "", "KCS 21 20 10", "13 (10m/s는 부록3에 없음, 15m/s로 근사)"),
    ("가설·장비", "타워크레인 운전작업", "", "", "순간풍속 15 m/s 이상", "", "", "KCS 21 20 10", "13"),
    ("가설·장비", "건설용 리프트", "", "", "평균풍속 15 m/s 초과", "", "", "KCS 21 20 10", "13"),
    ("가설·장비", "싣기 및 내리기(환경관리)", "", "", "평균풍속 8 m/s 이상", "", "", "KCS 21 20 15", "해당 조건 없음 (직접 검토)"),
    ("토공·옹벽", "연약지반 고결공", "4℃ 이하", "", "15 km/h 이상", "", "", "KCS 11 30 30", "5 (근사)"),
    ("토공·옹벽", "콘크리트 뿜어붙이기", "양생 후 3일간 10℃ 미만", "일기가 좋지 못한 기상조건", "일기가 좋지 못한 기상조건", "", "", "KCS 11 73 10", "5, 9"),
    ("토공·옹벽", "비탈면 녹화", "10℃ 이하", "", "", "", "25℃ 이상", "KCS 11 73 15", "해당 조건 없음 (직접 검토)"),
    ("토공·옹벽", "혼합토 표층개량제 포설 및 다짐", "일 최저기온 0℃", "", "", "", "", "EXCS 11 73 20", "5 (근사)"),
    ("토공·옹벽", "화강풍암토 비탈면 녹화", "", "", "", "", "일평균 기온 25℃ 이상", "EXCS 11 73 25", "해당 조건 없음 (직접 검토)"),
    ("토공·옹벽", "보강토 옹벽 뒷채움 다짐 및 블록 속채움", "1.5℃ 미만", "강우 시", "", "강설 시", "", "KCS 11 80 10", "5, 9, 6"),
    ("토공·옹벽", "돌(블록)쌓기 옹벽", "착수 전·쌓기 중·완료 후 48시간 동안 5℃ 미만", "", "", "", "30℃ 이상", "KCS 11 80 25", "5, 3 (근사)"),
    ("토공·옹벽", "H-PIPE 옹벽 엄지말뚝 이음(용접)", "0℃ 이하", "강우 시", "", "강설 시", "", "LHCS 11 80 30", "5, 9, 6"),
    ("굴착·복구", "굴착 및 복구공사 다짐", "0℃ 이하", "강우 시", "", "", "", "SMCS 11 85 10", "5, 9"),
    ("콘크리트", "일반 콘크리트", "", "강우 시", "", "강설 시", "", "KCS 14 20 10", "9, 6"),
    ("콘크리트", "일반 콘크리트 모르타르·그라우트", "작업 시작 전·중 5℃, 완료 후 48시간 동안 10℃ 이상", "", "", "", "작업 시작 전·중·완료 후 48시간 동안 30℃ 초과", "LHCS 14 20 10 15·20", "5, 3 (근사)"),
    ("콘크리트", "한중 콘크리트", "일평균 기온 -3℃ 이하", "", "", "", "", "EXCS 14 20 40", "5 (근사)"),
    ("콘크리트", "숏 콘크리트", "", "", "", "", "32℃ 이상", "KCS 14 20 51", "3 (33℃로 근사)"),
    ("콘크리트", "프리캐스트 콘크리트", "-5℃ 이하", "", "풍속 10 m/s 이상", "", "", "SMCS 14 20 52", "5, 13 (근사)"),
    ("강구조·교량", "일렉트로가스 용접", "", "", "풍속 2.7 m/s 이상", "", "", "KCS 14 31 20", "13 (근사)"),
    ("강구조·교량", "강구조 조립 및 설치(현장 접합)", "대기온도 -20℃ 이하", "강우 시, 강우 가능성이 있는 경우, 강우 직후", "풍속 2.0 m/s 이상", "", "", "KCS 14 31 30", "9, 13 (근사)"),
    ("강구조·교량", "강구조 도장", "5℃ 미만", "강우 시", "강풍", "강설 시", "43℃ 이상", "KCS 14 31 40", "9, 6, 13"),
    ("강구조·교량", "강교 제작 및 가설(현장 용접)", "5℃ 이하", "강우 시, 강우 가능성이 있는 경우, 강우 직후", "피복 아크용접 5 m/s 이상, 플럭스코어드 아크용접 2 m/s 이상", "", "", "EXCS 24 30 00", "5, 9, 13 (근사)"),
    ("강구조·교량", "신축이음", "5℃ 미만", "", "", "", "", "KCS 24 40 10", "5 (근사)"),
    ("강구조·교량", "교면방수", "5℃ 이하", "강우 시", "강풍 시", "", "30℃ 초과", "KCS 24 40 20", "5, 9, 13, 3"),
    ("강구조·교량", "일체식(노출) 시멘트콘크리트 교면포장", "7℃ 이하", "강우 시", "", "", "30℃ 이상", "EXCS 24 70 10", "5, 9, 3"),
    ("도로포장", "동상방지층·보조기층·기층 - 아스팔트콘크리트 기층", "5℃ 이하", "강우 시", "", "", "", "KCS 44 50 05", "5, 9"),
    ("도로포장", "동상방지층·보조기층·기층 - 시멘트안정처리 기층", "4℃ 이하", "강우 시", "", "", "", "KCS 44 50 05", "5, 9"),
    ("도로포장", "프라임코트", "10℃ 이하", "강우 시", "", "", "", "KCS 44 50 10", "5, 9"),
    ("도로포장", "택코트", "5℃ 이하", "강우 시", "", "", "", "KCS 44 50 10", "5, 9"),
    ("도로포장", "실코트", "10℃ 이하", "강우 시", "", "", "", "KCS 44 50 10", "5, 9"),
    ("도로포장", "아스팔트콘크리트 중간층·표층", "5℃ 이하", "강우 시", "", "", "", "KCS 44 50 10", "5, 9"),
    ("도로포장", "배수성 아스팔트콘크리트 포장", "5℃ 이하", "강우 시", "", "", "", "LHCS 44 50 10 30", "5, 9"),
    ("도로포장", "저소음 비배수성 아스팔트 포장", "5℃ 이하", "강우 시", "", "", "", "LHCS 44 50 10 35", "5, 9"),
    ("도로포장", "중온 아스팔트콘크리트 포장", "2℃ 이하", "강우 시", "", "", "", "LHCS 44 50 10 40", "5, 9"),
    ("도로포장", "투수 아스팔트콘크리트 포장", "5℃ 이하, 기층면 동결 시", "", "", "", "", "SMCS 44 50 10 45", "5"),
    ("도로포장", "시멘트콘크리트 포장(빈배합)", "일평균기온 4℃ 이하", "강우 시", "", "", "", "KCS 44 50 15", "5, 9"),
    ("도로포장", "경하중 시멘트콘크리트 포장", "4℃ 이하", "강우 시", "", "", "30℃ 이상", "LHCS 44 50 15 15", "5, 9, 3"),
    ("도로포장", "투수 콘크리트 포장", "5℃ 이하", "", "", "", "30℃ 이상", "SMCS 44 50 20", "5, 3"),
    ("도로포장", "경하중 아스팔트콘크리트 포장", "5℃ 이하", "강우 시", "", "", "", "SMCS 44 50 25 05", "5, 9"),
    ("도로포장", "인터로킹 블록포장", "바닥면이 얼어있을 경우", "강우 시", "", "", "", "SMCS 44 50 30 10", "5, 9"),
    ("도로안전", "노면표시(차선도색) 시공", "5℃ 이하", "", "", "", "", "KCS 44 60 05", "5 (근사)"),
    ("도로안전", "콘크리트 중앙분리대·시선유도 도장", "5℃ 미만", "", "", "", "43℃ 이상", "EXCS 44 60 05", "5"),
    ("도로안전", "시선유도시설(표지병) 시공", "5℃ 이하", "강우 시", "", "강설 시", "", "LHCS 44 60 05 15", "5, 9, 6"),
    ("도로안전", "미끄럼방지 포장", "5℃ 이하", "강우 시", "강풍 시", "", "", "LHCS 44 60 05 50", "5, 9, 13"),
    ("도로안전", "경계블록 및 L형 측구", "4℃ 이하", "강우 시", "", "", "30℃ 이상", "LHCS 44 60 05 55", "5, 9, 3"),
    ("도로포장", "타일포장", "4℃ 이하", "", "", "", "30℃ 이상", "LHCS 44 70 06", "5, 3"),
    ("도로포장", "투수시멘트콘크리트 포장", "5℃ 이하", "강우 시", "", "", "30℃ 이상", "LHCS 44 70 09", "5, 9, 3"),
    ("도로시설", "블록방음담장", "4℃ 이하", "강우 시", "", "", "30℃ 이상", "LHCS 44 80 06", "5, 9, 3"),
    ("도로시설", "비산먼지 방지시설(싣기 및 내리기)", "", "", "평균풍속 8 m/s 이상", "", "", "KCS 44 80 15", "해당 조건 없음 (직접 검토)"),
    ("도로유지", "콘크리트 표면보호제 도포(폴리머 시멘트 모르타르)", "5℃ 이하", "", "", "", "30℃ 이상", "EXCS 44 99 45", "5, 3"),
    ("상수도", "도복장 강관 용접접합", "-15℃ 이하", "강우 시", "", "강설 시", "35℃ 이상", "KCS 57 30 20", "8 (최저 -12℃ 이하로 근사), 9, 6, 2 (혹서 35℃는 체감/최고 35℃)"),
    ("상수도", "상수도 강관 접합", "-15℃ 이하", "강우 시", "", "강설 시", "35℃ 이상", "LHCS 57 30 20 05", "8 (근사), 9, 6, 2"),
    ("상수도", "주철관 접합", "-1℃ 이하", "", "", "", "", "LHCS 57 30 20 15", "5 (근사)"),
    ("상수도", "접합부 액상에폭시·폴리우레아 수지 도료 내부도장", "5℃ 이하", "강우 시", "강풍 시", "", "", "KCS 57 30 25", "5, 9, 13"),
    ("상수도", "정수처리시설 방수공사", "5℃ 미만", "강우 시, 강우 가능성이 있는 경우, 강우 직후", "강풍 시", "강설이 예상될 경우", "32℃ 초과", "KCS 57 40 15", "5, 9, 13, 6, 3"),
    ("하수관로", "굴착 및 되메우기, 포장공", "5℃ 이하", "강우 시", "", "", "", "KCS 61 20 15", "5, 9"),
    ("하수관로", "하수도관 관의 절단", "5℃ 이하", "", "", "", "35℃ 이상", "KCS 61 20 30", "5, 2"),
    ("하수관로", "비산먼지 방지시설(싣기 및 내리기)", "", "", "평균풍속 8 m/s 이상", "", "", "KCS 61 40 05", "해당 조건 없음 (직접 검토)"),
    ("하수관로", "기존 암거 보수(단면복구, 방청 도포재)", "5℃ 이하", "강우 시, 강우 가능성이 있는 경우", "", "", "30℃ 이상", "KCS 61 80 20", "5, 9, 3"),
]


def add_limit_reference_sheet(wb):
    ws = wb.create_sheet("작업제한기상조건(참고)")
    ws["A1"] = "작업제한 기상조건 참고표 (가이드라인 부록 2 - 토목 중심 발췌)"
    ws["A1"].font = Font(bold=True, size=13)
    ws["A2"] = ("부록 2의 기준값은 작업마다 다르고(4℃, 5℃, 10 m/s 등) 부록 3에서 미리 계산된 13개 조건과 딱 맞지 않는 경우가 많습니다. "
                "맨 오른쪽 '고를 조건 번호'는 가장 가까운 조건을 제가 제안한 것이므로, 현장 기준으로 판단해 '공기산정(1일작업량)' 시트의 기상조건 세트에 반영하고 "
                "근사한 경우에는 산정 근거에 이유를 적으세요. 전체 목록은 부록 2(원문 1355~1937행)를 확인하세요.")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:I2")
    ws.row_dimensions[2].height = 62
    heads = ["구분", "작업", "동절기·저온", "강우", "바람", "눈", "혹서기·고온", "출처", "고를 조건 번호(제안)"]
    for j, h in enumerate(heads):
        ws.cell(4, 1 + j).value = h
    _hdr(ws, 4, range(1, 10))
    for i, row in enumerate(LIMIT_ROWS):
        for j, v in enumerate(row):
            c = ws.cell(5 + i, 1 + j)
            c.value = v
            c.border = BOX
            c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "C5"
    ws.auto_filter.ref = f"A4:I{4 + len(LIMIT_ROWS)}"
    r0 = 6 + len(LIMIT_ROWS)
    ws.cell(r0, 1).value = "▼ 조건 번호 (공기산정 시트의 기상조건 표와 같은 번호)"
    ws.cell(r0, 1).font = BOLD
    names = json.load(open("data/weather.json", encoding="utf-8"))
    for i, c in enumerate(names, 1):
        ws.cell(r0 + i, 1).value, ws.cell(r0 + i, 2).value = i, c["name"]
    for col, w in zip("ABCDEFGHI", (12, 36, 24, 22, 24, 14, 20, 18, 30)):
        ws.column_dimensions[col].width = w
    return ws


def add_schedule_sheet(wb, info, n_hol):
    from openpyxl.utils import get_column_letter as L
    WD = "'공기산정(1일작업량)'!"
    ws = wb.create_sheet("월별산정표(자동)")
    ws["A1"] = "월별 산정표 — 착공 월과 작업일수를 넣으면 착공부터 종료까지 월별 표가 자동으로 만들어집니다"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = ("가이드라인 '공종별 공사기간 산정 예시(철골세우기)'와 같은 방식입니다. 월별 비작업일수 = A+B-C(월 8일 이상), 작업가능일수 = 달력일수-비작업일수, "
                "마지막 달은 '비작업일수 × 잔여작업일수 ÷ 그 달 작업가능일수'로 안분합니다. 지점은 '공기산정(1일작업량)' 시트의 기상 지점(B4)을 따릅니다.")
    ws["A2"].alignment = Alignment(wrap_text=True, vertical="top")
    ws.merge_cells("A2:N2")
    ws.row_dimensions[2].height = 48
    rows = [("착공 연도", f"={WD}B5"), ("착공 월", f"={WD}B6"), ("기상 지점", f"={WD}B4"),
            ("기상조건 세트", "세트1"), ("작업일수 (직접 입력)", None)]
    for i, (lab, val) in enumerate(rows):
        r = 4 + i
        ws.cell(r, 1).value, ws.cell(r, 2).value = lab, val
        ws.cell(r, 1).font = BOLD
        ws.cell(r, 2).border = BOX
        ws.cell(r, 2).fill = INPUT if lab != "기상 지점" else PatternFill("solid", fgColor="EDEDED")
    c0 = info["c0"]
    dv = DataValidation(type="list", formula1=f"={WD}$C${c0}:$F${c0}", allow_blank=False)
    ws.add_data_validation(dv); dv.add("B7")
    ws["C4"] = "※ 기본값은 공사기간 산정 시트의 착공 연·월입니다. 여기서 직접 바꿔 써도 됩니다."
    ws["C7"] = f'=IFERROR(INDEX({WD}$C$12:$F$12,MATCH(B7,{WD}$C${c0}:$F${c0},0)),"")'
    ws["C8"] = f'="비워 두면 공종표의 주공정 작업일수 합계("&SUMIF({WD}E{info["t1"]}:E{info["t2"]},"Y",{WD}G{info["t1"]}:G{info["t2"]})&"일)를 사용합니다"'
    ws["A9"], ws["B9"] = "적용 작업일수", (f'=IF(ISNUMBER(B8),B8,SUMIF({WD}E{info["t1"]}:E{info["t2"]},"Y",{WD}G{info["t1"]}:G{info["t2"]}))')
    ws["A9"].font = BOLD
    ws["B9"].fill, ws["B9"].border = RESULT, BOX
    ws["D7"] = f'=IFERROR(MATCH(B7,{WD}$C${c0}:$F${c0},0),1)'
    ws["D7"].font = Font(color="808080")

    first, last = 32, 91
    T = lambda col: f"${col}${first}:${col}${last}"
    ws["A11"] = "▼ 요약"
    ws["A11"].font = BOLD
    summary = [
        ("착공", f'=B4&"년 "&B5&"월"'),
        ("종료 예정", f'=IF(COUNT(A{first}:A{last})=0,"-",INDEX(B{first}:B{last},COUNT(A{first}:A{last})))'),
        ("소요 개월 수 (달력 기준)", f"=COUNT(A{first}:A{last})"),
        ("작업일수 합계", f"=SUM(K{first}:K{last})"),
        ("비작업일수 합계", f"=ROUND(SUM(L{first}:L{last}),1)"),
        ("산정 기간 (작업+비작업, 일)", f"=ROUNDUP(SUM(M{first}:M{last}),0)"),
        ("준비기간 (일)", f"={WD}B7"),
        ("정리기간 (일)", f"={WD}B8"),
        ("총 공사기간 (일)", "=B17+B18+B19"),
        ("환산 (개월)", "=ROUND(B20/30,1)"),
    ]
    for i, (lab, f) in enumerate(summary):
        r = 12 + i
        ws.cell(r, 1).value, ws.cell(r, 2).value = lab, f
        ws.cell(r, 1).font = BOLD
        ws.cell(r, 2).fill, ws.cell(r, 2).border = RESULT, BOX
    assert 12 + len(summary) < 27

    hr = first - 1
    ws.cell(hr - 1, 1).value = "▼ 월별 산정표"
    ws.cell(hr - 1, 1).font = BOLD
    heads = ["순번", "연월", "달력일수", "법정공휴일 B", "기후여건 A", "중복일수 C", "A+B-C", "적용 비작업일수(월8일↑)", "작업가능일수",
             "전월말 잔여 작업일수", "이 달 작업일수", "이 달 비작업일수", "이 달 소요일수", "월말 잔여 작업일수"]
    for j, h in enumerate(heads):
        ws.cell(hr, 1 + j).value = h
    _hdr(ws, hr, range(1, 15))
    ws.row_dimensions[hr].height = 48
    helper = ["k", "연", "월", "달력", "공휴", "기후A", "중복C", "적용", "가능", "전월누적", "표시"]
    for j, h in enumerate(helper):
        c = ws.cell(hr, 16 + j)
        c.value = h
        c.font = Font(color="808080", bold=True)
    sets = ",".join(f"{WD}$G${info['sum_row'][p]}:$R${info['sum_row'][p]}" for p in range(4))
    for i in range(60):
        r = first + i
        P, Q, R, S, Tt, U, V, W, X, Y, Z = (f"{L(16 + j)}{r}" for j in range(11))
        ws[P] = i + 1
        ws[Q] = f"=YEAR(EDATE(DATE($B$4,$B$5,1),{P}-1))"
        ws[R] = f"=MONTH(EDATE(DATE($B$4,$B$5,1),{P}-1))"
        ws[S] = f"=DAY(EOMONTH(DATE({Q},{R},1),0))"
        ws[Tt] = f"=IFERROR(INDEX(법정공휴일!$B$4:$M${n_hol},MATCH({Q},법정공휴일!$A$4:$A${n_hol},0),{R}),0)"
        ws[U] = f"=INDEX(CHOOSE($D$7,{sets}),1,{R})"
        ws[V] = f"=ROUND({U}*{Tt}/{S},1)"
        ws[W] = f"=ROUND(MAX({U}+{Tt}-{V},8),0)"
        ws[X] = f"={S}-{W}"
        ws[Y] = 0 if i == 0 else f"={L(25)}{r - 1}+{L(24)}{r - 1}"
        ws[Z] = f"=IF(AND(ISNUMBER($B$9),$B$9>{Y}),1,0)"
        show = lambda expr: f'=IF({Z}=1,{expr},"")'
        ws.cell(r, 1).value = show(P)
        ws.cell(r, 2).value = show(f'{Q}&"년 "&{R}&"월"')
        ws.cell(r, 3).value = show(S)
        ws.cell(r, 4).value = show(Tt)
        ws.cell(r, 5).value = show(U)
        ws.cell(r, 6).value = show(V)
        ws.cell(r, 7).value = show(f"ROUND({U}+{Tt}-{V},1)")
        ws.cell(r, 8).value = show(W)
        ws.cell(r, 9).value = show(X)
        ws.cell(r, 10).value = show(f"$B$9-{Y}")
        ws.cell(r, 11).value = show(f"MIN({X},$B$9-{Y})")
        ws.cell(r, 12).value = show(f"ROUND({W}*MIN({X},$B$9-{Y})/{X},1)")
        ws.cell(r, 13).value = show(f"K{r}+L{r}")
        ws.cell(r, 14).value = show(f"J{r}-K{r}")
        for c in range(1, 15):
            ws.cell(r, c).border = BOX
        for c in range(16, 27):
            ws.cell(r, c).font = Font(color="A0A0A0")
    ws.column_dimensions["A"].width = 34
    ws.column_dimensions["B"].width = 20
    for col in range(3, 15):
        ws.column_dimensions[L(col)].width = 13
    return ws
