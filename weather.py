import calendar
from collections import defaultdict

import requests

ASOS_URL = "http://apis.data.go.kr/1360000/AsosDalyInfoService/getWthrDataList"

# 조건 순서는 별표2 시트의 10개 행 순서와 같다
CONDITIONS = [
    lambda d: d["maxTa"] >= 33,
    lambda d: d["maxTa"] >= 35,
    lambda d: d["maxTa"] <= 0,
    lambda d: d["snow"] >= 5,
    lambda d: d["maxTa"] <= 0 and d["snow"] >= 5,
    lambda d: d["minTa"] <= -12,
    lambda d: d["rain"] >= 5,
    lambda d: d["rain"] >= 10,
    lambda d: d["rain"] >= 20,
    lambda d: d["wind"] >= 15,
]


def _num(v, default=None):
    try:
        return float(v)
    except (TypeError, ValueError):
        return default


def fetch_daily(api_key, station, year):
    params = {
        "serviceKey": api_key, "pageNo": 1, "numOfRows": 999, "dataType": "JSON",
        "dataCd": "ASOS", "dateCd": "DAY", "stnIds": station,
        "startDt": f"{year}0101", "endDt": f"{year}1231",
    }
    days = []
    while True:
        r = requests.get(ASOS_URL, params=params, timeout=30)
        r.raise_for_status()
        body = r.json()["response"]
        if body["header"]["resultCode"] != "00":
            raise RuntimeError(body["header"]["resultMsg"])
        items = body["body"]["items"]["item"]
        days += items
        if params["pageNo"] * params["numOfRows"] >= body["body"]["totalCount"]:
            return days
        params["pageNo"] += 1


def normalize(item):
    maxta, minta = _num(item.get("maxTa")), _num(item.get("minTa"))
    if maxta is None or minta is None:
        return None
    return {
        "month": int(item["tm"][5:7]),
        "maxTa": maxta,
        "minTa": minta,
        "rain": _num(item.get("sumRn"), 0.0),
        "snow": _num(item.get("ddMes"), 0.0),
        "wind": _num(item.get("maxInsWs"), 0.0),
    }


def monthly_averages(daily_items, years):
    """10개 조건 x 12개월의 '연평균 해당 일수' 표를 반환한다."""
    counts = [[0] * 12 for _ in CONDITIONS]
    for item in daily_items:
        d = normalize(item)
        if d is None:
            continue
        for i, cond in enumerate(CONDITIONS):
            if cond(d):
                counts[i][d["month"] - 1] += 1
    return [[round(c / years, 1) for c in row] for row in counts]
