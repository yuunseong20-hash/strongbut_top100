# Strong Buy Top 100 (자동 갱신 웹페이지)

최근 2개월 애널리스트 Strong Buy 의견이 많은 미국 대형주(S&P 500 + Nasdaq-100) 상위 100개를 보여주는 웹페이지입니다.
GitHub Actions가 한국시간 월~금 오전 7시에 Yahoo Finance 데이터를 수집하고, GitHub Pages가 페이지를 게시합니다.

## 구성
- `update.py` — 데이터 수집 → `docs/data.json`, `docs/strong_buy_top100.xlsx` 생성
- `docs/index.html` — 웹페이지 (검색, 열 정렬, 순위 변동, 의견 분포 막대, 엑셀 다운로드)
- `.github/workflows/update.yml` — 자동 실행 일정

## 설정 (처음 한 번)
1. GitHub에서 새 저장소를 만들고 이 폴더 내용을 그대로 올립니다.
2. 저장소 **Settings → Pages** → Source: `Deploy from a branch`, Branch: `main`, 폴더: `/docs` → Save
3. **Actions** 탭 → `Strong Buy Top100 자동 갱신` → **Run workflow** 로 첫 데이터를 생성합니다 (약 10분).
4. 몇 분 뒤 `https://<아이디>.github.io/<저장소이름>/` 에서 확인합니다.

## 바꾸고 싶을 때
- 갱신 시각: `update.yml`의 `cron` (UTC 기준, 한국시간 −9시간)
- 순위 개수: `update.py`의 `TOP_N`
- 대상 종목: `update.py`의 `get_universe()`
