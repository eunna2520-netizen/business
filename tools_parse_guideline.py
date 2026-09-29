"""가이드라인 마크다운의 부록3 표를 data/weather.json으로 추출한다."""
import json
import re
import sys

src, out = sys.argv[1], sys.argv[2]
lines = open(src, encoding="utf-8").read().splitlines()
start = next(i for i, l in enumerate(lines) if l.startswith("(1) 혹서기") and "체감온도 33" in l and i > 2000)
end = next(i for i, l in enumerate(lines) if i > start and l.startswith("1. 토목분야"))
head = re.compile(r"^\((\d+)\)\s+(.+)$")
row = re.compile(r"^\s*(\d+)\s+(\S+)\s+(\d{2,3})\s+(\S+)\s+((?:\d+\.\d+\s+){12}\d+\.\d+)\s*$")

conds, cur = [], None
for l in lines[start:end]:
    m = head.match(l)
    if m and int(m.group(1)) == len(conds) + 1:
        cur = {"name": m.group(2).strip(), "stations": {}}
        conds.append(cur)
        continue
    m = row.match(l)
    if m and cur is not None:
        vals = [float(x) for x in m.group(5).split()]
        cur["stations"][m.group(3)] = {"region": m.group(2), "name": m.group(4), "months": vals[:12], "total": vals[12]}

for c in conds:
    print(len(c["stations"]), c["name"])
json.dump(conds, open(out, "w", encoding="utf-8"), ensure_ascii=False)
