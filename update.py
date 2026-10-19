"""
최근 2개월 Strong Buy 상위 100개 회사 — 웹페이지 데이터 생성 (GitHub Actions에서 실행)
- 데이터: Yahoo Finance (yfinance) 애널리스트 추천 집계 (이번 달 + 지난 달 합산)
- 대상: S&P 500 + Nasdaq-100 구성종목
- 결과: strong_buy_top100_YYYYMMDD.xlsx (첨부 이미지와 같은 서식)

설치:  pip install yfinance pandas openpyxl lxml requests
실행:  python strong_buy_top100.py
       python strong_buy_top100.py --test  (10종목만 빠르게 테스트)
"""
import sys, time, datetime as dt
from io import StringIO

import pandas as pd
import requests
import yfinance as yf
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

TOP_N = 100
UA = {"User-Agent": "Mozilla/5.0"}


# ---------- 1. 대상 종목 ----------
def get_universe():
    tickers = set()
    srcs = [
        ("https://en.wikipedia.org/wiki/List_of_S%26P_500_companies", "Symbol"),
        ("https://en.wikipedia.org/wiki/Nasdaq-100", "Ticker"),
    ]
    for url, col in srcs:
        try:
            html = requests.get(url, headers=UA, timeout=20).text
            for t in pd.read_html(StringIO(html)):
                c = next((x for x in (col, "Ticker", "Symbol") if x in t.columns), None)
                if c and len(t) >= 90:
                    tickers |= set(t[c].astype(str).str.strip())
                    break
        except Exception as e:
            print(f"[경고] 종목 목록 실패: {url} ({e})")
    # Yahoo 표기: BRK.B -> BRK-B
    return sorted(t.replace(".", "-") for t in tickers if t and t != "nan")


# ---------- 2. 종목별 데이터 ----------
def fetch(ticker):
    tk = yf.Ticker(ticker)
    rec = tk.recommendations  # period: 0m(이번 달), -1m, -2m, -3m
    if rec is None or rec.empty:
        return None
    rec = rec[rec["period"].isin(["0m", "-1m"])]
    s = rec[["strongBuy", "buy", "hold", "sell", "strongSell"]].sum()
    try:
        mcap = tk.fast_info["market_cap"]
    except Exception:
        mcap = (tk.info or {}).get("marketCap")
    return {
        "티커": ticker,
        "시총(억USD)": (mcap or 0) / 1e8,
        "S_buy": int(s["strongBuy"]), "buy": int(s["buy"]), "hold": int(s["hold"]),
        "sell": int(s["sell"]), "S_sell": int(s["strongSell"]),
    }


def collect(tickers):
    rows = []
    for i, t in enumerate(tickers, 1):
        for attempt in range(3):
            try:
                r = fetch(t)
                if r:
                    rows.append(r)
                break
            except Exception as e:
                if attempt == 2:
                    print(f"  {t} 실패: {e}")
                time.sleep(2)
        if i % 25 == 0:
            print(f"  {i}/{len(tickers)} 완료")
        time.sleep(0.3)  # 과도한 요청 방지
    return pd.DataFrame(rows)


# ---------- 3. 엑셀 서식 ----------
def month_label(today):
    prev = (today.replace(day=1) - dt.timedelta(days=1))
    return f"{prev.strftime('%b')}~{today.strftime('%b')}"


def write_excel(df, today, path):
    wb = Workbook(); ws = wb.active; ws.title = "Strong Buy Top100"
    thin = Side(style="thin", color="999999")
    box = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    black = PatternFill("solid", fgColor="000000")
    white_b = Font(name="맑은 고딕", bold=True, color="FFFFFF", size=11)
    base = Font(name="맑은 고딕", size=10)

    for col, w in zip("ABCDEFGHI", [12, 12, 13, 8, 8, 8, 8, 8, 2]):
        ws.column_dimensions[col].width = w

    # 상단
    ws["C1"] = "■ 필터"; ws["C1"].font = Font(name="맑은 고딕", size=11)
    ws["A2"] = "갱신 날짜"; ws["A2"].font = base
    ws["C2"] = f"- 최근 2개월 strong_buy 상위 {TOP_N}개 회사"; ws["C2"].font = Font(name="맑은 고딕", size=11)
    ws["A3"] = today.strftime("%Y. %m. %d"); ws["A3"].font = Font(name="맑은 고딕", bold=True, size=10)

    # 헤더 (4~5행)
    for rng, txt in [("A4:A5", "No."), ("B4:B5", "티커"), ("C4:C5", "시총\n(억USD)"),
                     ("D4:H4", f"2달간 애널리스트 의견({month_label(today)})")]:
        ws.merge_cells(rng)
        c = ws[rng.split(":")[0]]; c.value = txt; c.font = white_b; c.alignment = center
    for row in ws["A4:H5"]:
        for c in row:
            c.fill = black; c.border = box; c.alignment = center
    colors = {"S_buy": "00B050", "buy": "92D050", "hold": "FFFF00", "sell": "FFC000", "S_sell": "FF0000"}
    for col, (k, color) in zip("DEFGH", colors.items()):
        c = ws[f"{col}5"]; c.value = k; c.fill = PatternFill("solid", fgColor=color)
        c.font = Font(name="맑은 고딕", bold=True, size=10,
                      color="FFFFFF" if k in ("S_buy", "S_sell") else "000000")
    ws.row_dimensions[4].height = 20; ws.row_dimensions[5].height = 18

    # 데이터
    for i, rec in enumerate(df.itertuples(index=False), 1):
        r = 5 + i
        vals = [i] + list(rec)
        for col, v in zip("ABCDEFGH", vals):
            c = ws[f"{col}{r}"]; c.value = v; c.font = base; c.border = box
            c.alignment = Alignment(horizontal="right" if col == "C" else "center", vertical="center")
        ws[f"C{r}"].number_format = "#,##0.00"
    ws.freeze_panes = "A6"
    wb.save(path)


def main():
    import json, os
    today = (dt.datetime.utcnow() + dt.timedelta(hours=9)).date()  # 한국 날짜
    tickers = get_universe()
    if "--test" in sys.argv:
        tickers = tickers[:10]
    if not tickers:
        sys.exit("종목 목록을 가져오지 못했습니다.")
    print(f"대상 {len(tickers)}개 종목 수집 시작…")
    df = collect(tickers)
    if df.empty:
        sys.exit("데이터를 하나도 받지 못했습니다. 기존 페이지를 유지합니다.")
    df = (df.sort_values(["S_buy", "buy", "시총(억USD)"], ascending=False)
            .head(TOP_N).reset_index(drop=True))
    cols = ["티커", "시총(억USD)", "S_buy", "buy", "hold", "sell", "S_sell"]

    os.makedirs("docs", exist_ok=True)
    json_path = "docs/data.json"

    # 직전 갱신 순위 (순위 변동 표시용)
    prev_rank, prev_date = {}, None
    if os.path.exists(json_path):
        with open(json_path, encoding="utf-8") as f:
            old = json.load(f)
        if old.get("updated") != today.isoformat():
            prev_rank = {r["ticker"]: r["rank"] for r in old["rows"]}
            prev_date = old.get("updated")
        else:  # 같은 날 재실행이면 그 전 비교 기준 유지
            prev_rank = {r["ticker"]: r["rank"] - (r["change"] or 0) for r in old["rows"]
                         if r.get("change") is not None}
            prev_date = old.get("prev_updated")

    rows = []
    for i, r in enumerate(df[cols].itertuples(index=False), 1):
        t = r[0]
        rows.append({
            "rank": i, "ticker": t, "mcap": round(float(r[1]), 2),
            "sbuy": int(r[2]), "buy": int(r[3]), "hold": int(r[4]),
            "sell": int(r[5]), "ssell": int(r[6]),
            "change": (prev_rank[t] - i) if t in prev_rank else None,
            "new": bool(prev_rank) and t not in prev_rank,
        })
    data = {
        "updated": today.isoformat(),
        "prev_updated": prev_date,
        "period": month_label(today),
        "universe": "S&P 500 + Nasdaq-100",
        "count": len(tickers),
        "rows": rows,
    }
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)
    write_excel(df[cols], today, "docs/strong_buy_top100.xlsx")
    print(f"저장 완료: {json_path}, docs/strong_buy_top100.xlsx")


if __name__ == "__main__":
    main()
