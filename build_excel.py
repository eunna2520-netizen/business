"""data/weather.json + data/template.xlsx -> 공기산정_2026기준.xlsx"""
import json

import openpyxl
from openpyxl.styles import Alignment, Font
from openpyxl.worksheet.datavalidation import DataValidation

import sys
conds = json.load(open(sys.argv[1] if len(sys.argv) > 1 else "data/weather.json", encoding="utf-8"))
import sys
wb = openpyxl.load_workbook("data/template.xlsx")

# 1) 전국 지점 데이터 시트
ds = wb.create_sheet("기상자료")
ds.append(["지점코드", "행정구역", "지점명", "조건번호", "조건", *[f"{m}월" for m in range(1, 13)], "소계"])
codes = sorted(conds[0]["stations"], key=lambda k: (conds[0]["stations"][k]["region"], conds[0]["stations"][k]["name"]))
for code in codes:
    for i, c in enumerate(conds, 1):
        s = c["stations"][code]
        ds.append([int(code), s["region"], s["name"], i, c["name"], *s["months"], s["total"]])
for cell in ds[1]:
    cell.font = Font(bold=True)
ds.freeze_panes = "A2"
ds.column_dimensions["E"].width = 34
n = ds.max_row

ls = wb.create_sheet("지점목록")
ls.append(["지점명", "행정구역", "지점코드"])
for code in codes:
    s = conds[0]["stations"][code]
    ls.append([s["name"], s["region"], int(code)])
for cell in ls[1]:
    cell.font = Font(bold=True)
np_ = ls.max_row

# 2) 별표2 시트를 수식으로 전환
wx = wb["별표2-비작업일수"]
wx["A1"] = "별표 2. 기상상태에 따른 지역별 비작업일수(2015년-2024년)  ※ 출처: 2026년 적정 공사기간 확보를 위한 가이드라인 부록 3"
wx["A2"], wx["D2"] = "혹서기 기준", "체감온도"
wx["A2"].font = wx["D2"].font = Font(bold=True)
dv = DataValidation(type="list", formula1='"체감온도,일최고기온"', allow_blank=False)
wx.add_data_validation(dv); dv.add("D2")
wx["E2"] = "← 혹서기를 체감온도/일최고기온 중 무엇으로 볼지 선택 (가이드라인 부록3 (1)(2) 또는 (3)(4))"

hot = lambda a, b: f'IF($D$2="체감온도",{a},{b})'
slot_cond = [hot(1, 3), hot(2, 4), 5, 6, 7, 8, 10, 11, 12, 13]
slot_label = [
    f'=IF($D$2="체감온도","혹서기\n체감온도 33°C 이상","혹서기\n일최고기온 33°C 이상")',
    f'=IF($D$2="체감온도","혹서기\n체감온도 35°C 이상","혹서기\n일최고기온 35°C 이상")',
    "동절기\n일최고기온 0°C 이하", "동절기\n신적설 5㎝ 이상",
    "동절기 : 일최고기온 0°C 이하 및 신적설 5㎝ 이상", "동절기 : 일최저기온 -12°C 이하",
    "일강수량 : 5㎜ 이상", "일강수량 : 10㎜ 이상", "일강수량 : 20㎜ 이상",
    "일최대순간풍속 : 15m/s 이상",
]
wx["R3"] = "조건번호(자동)"
for first, default in ((5, "서귀포"), (19, "제주")):
    last = first + 9
    code_cell = f"$A${first}"
    wx.cell(first, 2).value = default
    wx.cell(first, 1).value = f"=INDEX(지점목록!$C$2:$C${np_},MATCH($B${first},지점목록!$A$2:$A${np_},0))"
    d = DataValidation(type="list", formula1=f"=지점목록!$A$2:$A${np_}", allow_blank=False)
    wx.add_data_validation(d); d.add(f"B{first}")
    for k in range(10):
        r = first + k
        wx.cell(r, 3).value = "2015~2024"
        wx.cell(r, 18).value = f"={slot_cond[k]}" if isinstance(slot_cond[k], str) else slot_cond[k]
        wx.cell(r, 17).value = slot_label[k]
        wx.cell(r, 17).alignment = Alignment(wrap_text=True, vertical="center")
        for m in range(12):
            col = openpyxl.utils.get_column_letter(6 + m)  # 기상자료 시트의 1~12월 열(F~Q)
            wx.cell(r, 4 + m).value = (
                f"=SUMIFS(기상자료!${col}$2:${col}${n},기상자료!$A$2:$A${n},{code_cell},기상자료!$D$2:$D${n},$R{r})")
        wx.cell(r, 16).value = f"=SUM(D{r}:O{r})"
wx["A3"] = wx["A17"] = "지점코드"
wx["B3"].value = wx["B17"].value = "지역 (선택)"
wx.column_dimensions["R"].width = 14

# 3) 공기산정 시트 연결
main = wb["공기산정"]
main["Y19"], main["Z19"] = "='별표2-비작업일수'!A5", "='별표2-비작업일수'!B5"
main["Y27"], main["Z27"] = "='별표2-비작업일수'!A19", "='별표2-비작업일수'!B19"
main["AN21"] = "=SUM(AB21:AM21)"
main["AA20"].value  # 표시용 셀 확인만

import sheets_civil
n_hol = sheets_civil.add_holiday_sheet(wb)
sheets_civil.add_facility_sheet(wb)
sheets_civil.add_workday_sheet(wb, n, np_, n_hol)
wb.calculation.fullCalcOnLoad = True
wb.save("공기산정_2026기준.xlsx")
print("saved", n - 1, "data rows")
