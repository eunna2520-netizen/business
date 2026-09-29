"""기상청 ASOS 일자료 API로 비작업일수 표를 다시 계산해 data/weather_custom.json을 만든다.

사용: WEATHER_API_KEY=<Decoding 키> python recalc.py --stations 184 189 --start 2016 --end 2025
체감온도 기준(조건 1, 2)은 API에 없으므로 가이드라인 값을 그대로 둔다.
"""
import argparse
import json
import os
import sys
import time

import requests

URL = "http://apis.data.go.kr/1360000/AsosDalyInfoService/getWthrDataList"

# data/weather.json의 조건 순서와 같다 (None = API로 계산 불가, 가이드라인 값 유지)
CONDITIONS = [
    None, None,
    lambda d: d["maxTa"] is not None and d["maxTa"] >= 33,
    lambda d: d["maxTa"] is not None and d["maxTa"] >= 35,
    lambda d: d["maxTa"] is not None and d["maxTa"] <= 0,
    lambda d: d["snow"] >= 5,
    lambda d: d["maxTa"] is not None and d["maxTa"] <= 0 or d["snow"] >= 5,
    lambda d: d["minTa"] is not None and d["minTa"] <= -12,
    lambda d: d["rain"] >= 3,
    lambda d: d["rain"] >= 5,
    lambda d: d["rain"] >= 10,
    lambda d: d["rain"] >= 20,
    lambda d: d["wind"] is not None and d["wind"] >= 15,
]


def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def fetch_year(key, stn, year):
    params = {"serviceKey": key, "pageNo": 1, "numOfRows": 999, "dataType": "JSON",
              "dataCd": "ASOS", "dateCd": "DAY", "stnIds": stn,
              "startDt": f"{year}0101", "endDt": f"{year}1231"}
    items = []
    while True:
        r = requests.get(URL, params=params, timeout=60)
        r.raise_for_status()
        try:
            resp = r.json()["response"]
        except ValueError:
            raise SystemExit(f"JSON이 아닌 응답: {r.text[:200]}")
        if str(resp["header"]["resultCode"]).lstrip("0") not in ("",):
            raise SystemExit(f"API 오류: {resp['header']['resultMsg']}")
        body = resp["body"]
        got = body["items"]["item"] if body.get("items") else []
        items += got if isinstance(got, list) else [got]
        if params["pageNo"] * params["numOfRows"] >= int(body["totalCount"]):
            return items
        params["pageNo"] += 1
        time.sleep(0.1)


def to_day(item):
    return {"month": int(item["tm"][5:7]), "year": int(item["tm"][:4]),
            "maxTa": num(item.get("maxTa")), "minTa": num(item.get("minTa")),
            "rain": num(item.get("sumRn")) or 0.0, "snow": num(item.get("ddMefs")) or 0.0,
            "wind": num(item.get("maxInsWs"))}


def monthly(days):
    years = len({d["year"] for d in days}) or 1
    out = []
    for cond in CONDITIONS:
        if cond is None:
            out.append(None)
            continue
        row = [0] * 12
        for d in days:
            if cond(d):
                row[d["month"] - 1] += 1
        out.append([round(c / years, 1) for c in row])
    return out


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--stations", nargs="+", required=True, help="지점번호 (예: 184 189)")
    ap.add_argument("--start", type=int, required=True)
    ap.add_argument("--end", type=int, required=True)
    ap.add_argument("--out", default="data/weather_custom.json")
    a = ap.parse_args()
    key = os.environ.get("WEATHER_API_KEY", "").strip()
    if not key:
        sys.exit("환경변수 WEATHER_API_KEY에 Decoding 인증키를 넣으세요.")

    base = json.load(open("data/weather.json", encoding="utf-8"))
    for stn in a.stations:
        if stn not in base[0]["stations"]:
            sys.exit(f"지점 {stn}이(가) data/weather.json에 없습니다.")
        days = []
        for y in range(a.start, a.end + 1):
            days += [to_day(i) for i in fetch_year(key, stn, y)]
            print(stn, y, "누적", len(days), "일", flush=True)
        for c, row in zip(base, monthly(days)):
            if row is not None:
                s = c["stations"][stn]
                s["months"], s["total"] = row, round(sum(row), 1)
    json.dump(base, open(a.out, "w", encoding="utf-8"), ensure_ascii=False)
    print("저장:", a.out, f"(기간 {a.start}~{a.end}, 체감온도 조건은 가이드라인 값 유지)")


if __name__ == "__main__":
    main()
