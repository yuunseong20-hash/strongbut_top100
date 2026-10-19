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
UA = {"User-Agent": "StrongBuyTop100/1.0 (https://github.com/yuunseong20-hash/strongbut_top100) python-requests"}

# Wikipedia 접속 실패 대비 Nasdaq-100 기본 목록 (2026년 기준)
NDX_FALLBACK = """ADBE AMD ABNB ALNY GOOGL GOOG AMZN AEP AMGN ADI AAPL AMAT APP ARM ASML ALAB ADSK ADP AXON BKR BKNG AVGO CDNS CTAS CSCO CCEP CMCSA CEG CPRT CRWV COST CRWD CSX DDOG DXCM FANG DASH EA EXC FAST FER FTNT GEHC GILD HON IDXX INTC INTU ISRG KDP KLAC KHC LRCX LIN LITE MAR MRVL MELI META MCHP MU MSFT MSTR MDLZ MPWR MNST NBIS NFLX NVDA NXPI ORLY ODFL PCAR PLTR PANW PAYX PYPL PDD PEP QCOM REGN RKLB ROP ROST SNDK STX SHOP SBUX SNPS TMUS TTWO TER TSLA TXN TRI VRTX WMT WBD WDC WDAY XEL""".split()


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
            found = False
            for t in pd.read_html(StringIO(html)):
                c = next((x for x in (col, "Ticker", "Symbol") if x in t.columns), None)
                if c and len(t) >= 90:
                    tickers |= set(t[c].astype(str).str.strip())
                    print(f"  {url.rsplit('/',1)[-1]}: {len(t)}개")
                    found = True
                    break
            if not found:
                print(f"[경고] 표를 찾지 못함: {url}")
        except Exception as e:
            print(f"[경고] 종목 목록 실패: {url} ({e})")
    tickers |= set(NDX_FALLBACK)
    # Yahoo 표기: BRK.B -> BRK-B
    return sorted(t.replace(".", "-") for t in tickers if t and t != "nan")


# ---------- 2. 종목별 데이터 ----------
EVENT_DAYS = 90          # 목표가 변경 내역 보관 기간(일)
EVENTS = []              # 증권사별 의견·목표가 변경 기록
TARGETS = {}             # 종목별 컨센서스 목표가


def _num(v):
    try:
        v = float(v)
        return round(v, 2) if v == v and v > 0 else None
    except (TypeError, ValueError):
        return None


def fetch_targets(tk, ticker):
    """목표가 컨센서스와 최근 증권사별 변경 내역 (실패해도 본 표에는 영향 없음)"""
    price = None
    try:
        price = _num(tk.fast_info["last_price"])
    except Exception:
        pass
    try:
        pt = tk.analyst_price_targets or {}
        TARGETS[ticker] = {
            "price": price or _num(pt.get("current")),
            "mean": _num(pt.get("mean")), "median": _num(pt.get("median")),
            "high": _num(pt.get("high")), "low": _num(pt.get("low")),
        }
    except Exception as e:
        print(f"  {ticker} 목표가 실패: {e}")
    try:
        ud = tk.upgrades_downgrades
        if ud is None or ud.empty:
            return
        ud = ud.reset_index()
        datecol = "GradeDate" if "GradeDate" in ud.columns else ud.columns[0]
        ud[datecol] = pd.to_datetime(ud[datecol], utc=True, errors="coerce")
        cutoff = pd.Timestamp.utcnow() - pd.Timedelta(days=EVENT_DAYS)
        for r in ud[ud[datecol] >= cutoff].to_dict("records"):
            EVENTS.append({
                "date": r[datecol].strftime("%Y-%m-%d"),
                "ticker": ticker,
                "firm": str(r.get("Firm") or ""),
                "from_grade": str(r.get("FromGrade") or ""),
                "to_grade": str(r.get("ToGrade") or ""),
                "action": str(r.get("Action") or ""),          # up/down/main/init/reit
                "pt_action": str(r.get("priceTargetAction") or ""),  # Raises/Lowers/Maintains/Announces
                "pt_from": _num(r.get("priorPriceTarget")),
                "pt_to": _num(r.get("currentPriceTarget")),
                "price": price,
            })
    except Exception as e:
        print(f"  {ticker} 변경내역 실패: {e}")


def fetch(ticker):
    tk = yf.Ticker(ticker)
    fetch_targets(tk, ticker)
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
            .reset_index(drop=True))   # 전체 순위 (웹은 상위 100 기본 표시, 검색 시 전체)
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
            "new": bool(prev_rank) and t not in prev_rank and i <= TOP_N,
        })
    data = {
        "updated": today.isoformat(),
        "prev_updated": prev_date,
        "period": month_label(today),
        "universe": "S&P 500 + Nasdaq-100",
        "count": len(tickers),
        "rows": rows,
    }
    for row in rows:
        t = TARGETS.get(row["ticker"], {})
        row["price"], row["tp"] = t.get("price"), t.get("mean")
        row["upside"] = round((t["mean"] / t["price"] - 1) * 100, 1) if t.get("mean") and t.get("price") else None
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=1)

    # 목표가 변경 내역
    ev = sorted(EVENTS, key=lambda e: (e["date"], e["ticker"]), reverse=True)
    top_set = {r["ticker"] for r in rows if r["rank"] <= TOP_N}
    for e in ev:
        e["top100"] = e["ticker"] in top_set
        if e["pt_from"] and e["pt_to"]:
            e["pt_chg"] = round((e["pt_to"] / e["pt_from"] - 1) * 100, 1)
        else:
            e["pt_chg"] = None
        e["upside"] = round((e["pt_to"] / e["price"] - 1) * 100, 1) if e["pt_to"] and e["price"] else None
    with open("docs/targets.json", "w", encoding="utf-8") as f:
        json.dump({"updated": today.isoformat(), "days": EVENT_DAYS,
                   "events": ev, "targets": TARGETS}, f, ensure_ascii=False, separators=(",", ":"))
    print(f"목표가 변경 {len(ev)}건 저장")
    write_excel(df[cols].head(TOP_N), today, "docs/strong_buy_top100.xlsx")
    print(f"저장 완료: {json_path}, docs/strong_buy_top100.xlsx")


if __name__ == "__main__":
    main()
