"""資源エネルギー庁「石油製品価格調査」の週次ファイルから、都道府県別の灯油・ガソリン価格サイトを作る。

使い方:
  python build.py            # 最新のExcelを取って site/ に書き出す
  python build.py --no-fetch # data/s5.xlsx をそのまま使う
"""
import datetime as dt
import html
import json
import re
import sys
import urllib.request
from pathlib import Path

import openpyxl

ROOT = Path(__file__).resolve().parent
DATA = ROOT / "data"
SITE = ROOT / "site"
SITE_NAME = "灯油とガソリンの値段（都道府県別）"
SITE_URL = "https://toyu-kakaku.com"
BASE_URL = "https://www.enecho.meti.go.jp/statistics/petroleum_and_lpgas/pl007/"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/141.0 Safari/537.36"

SHEETS = {  # シート名 -> (キー, 表示名, 単位)
    "灯油店頭": ("kerosene", "灯油（店頭）", "円/18L"),
    "灯油配達": ("kerosene_dlv", "灯油（配達）", "円/18L"),
    "レギュラー": ("regular", "レギュラー", "円/L"),
    "ハイオク": ("premium", "ハイオク", "円/L"),
    "軽油": ("diesel", "軽油", "円/L"),
}
# 列見出し -> 都道府県名（北海道・沖縄は「局」の列がその県の値）
RENAME = {"北海道局": "北海道", "沖縄局": "沖縄"}
PREFS = ["北海道", "青森", "岩手", "宮城", "秋田", "山形", "福島", "茨城", "栃木", "群馬", "埼玉", "千葉", "東京",
         "神奈川", "新潟", "富山", "石川", "福井", "山梨", "長野", "岐阜", "静岡", "愛知", "三重", "滋賀", "京都",
         "大阪", "兵庫", "奈良", "和歌山", "鳥取", "島根", "岡山", "広島", "山口", "徳島", "香川", "愛媛", "高知",
         "福岡", "佐賀", "長崎", "熊本", "大分", "宮崎", "鹿児島", "沖縄"]
SLUG = dict(zip(PREFS, ["hokkaido", "aomori", "iwate", "miyagi", "akita", "yamagata", "fukushima", "ibaraki",
                        "tochigi", "gunma", "saitama", "chiba", "tokyo", "kanagawa", "niigata", "toyama",
                        "ishikawa", "fukui", "yamanashi", "nagano", "gifu", "shizuoka", "aichi", "mie", "shiga",
                        "kyoto", "osaka", "hyogo", "nara", "wakayama", "tottori", "shimane", "okayama",
                        "hiroshima", "yamaguchi", "tokushima", "kagawa", "ehime", "kochi", "fukuoka", "saga",
                        "nagasaki", "kumamoto", "oita", "miyazaki", "kagoshima", "okinawa"]))
FULL = {p: p + ("都" if p == "東京" else "道" if p == "北海道" else "府" if p in ("京都", "大阪") else "県")
        for p in PREFS}
START = dt.date(2004, 6, 14)  # これ以前は税抜きで、県別の値もない


def fetch():
    req = urllib.request.Request(BASE_URL + "results.html", headers={"User-Agent": UA})
    page = urllib.request.urlopen(req, timeout=60).read().decode("utf-8", "replace")
    m = re.search(r'href="[^"]*?(xlsx/\d+s5\.xlsx)"', page)
    if not m:
        sys.exit("週次ファイルのリンクが見つかりません")
    req = urllib.request.Request(BASE_URL + m.group(1),
                                 headers={"User-Agent": UA, "Referer": BASE_URL + "results.html"})
    body = urllib.request.urlopen(req, timeout=120).read()
    if not body.startswith(b"PK"):
        sys.exit("Excelではないものが返りました")
    DATA.mkdir(exist_ok=True)
    (DATA / "s5.xlsx").write_bytes(body)


def load():
    wb = openpyxl.load_workbook(DATA / "s5.xlsx", read_only=True)
    out = {}
    for ws in wb.worksheets:
        if ws.title not in SHEETS:
            continue
        key = SHEETS[ws.title][0]
        rows = list(ws.iter_rows(values_only=True))
        head = [RENAME.get(re.sub(r"\s", "", str(h)), re.sub(r"\s", "", str(h))) for h in rows[0]]
        cols = {p: head.index(p) for p in ["全国"] + PREFS}
        series = {"dates": [], **{p: [] for p in cols}}
        for r in rows[1:]:
            if not isinstance(r[1], dt.datetime) or r[1].date() < START:
                continue
            series["dates"].append(r[1].date().isoformat())
            for p, i in cols.items():
                v = r[i]
                series[p].append(int(v) if isinstance(v, (int, float)) else None)
        out[key] = series
    return out


def last(vals, back=0):
    xs = [v for v in vals if v is not None]
    return xs[-1 - back] if len(xs) > back else None


def year_ago(s, p):
    d = dt.date.fromisoformat(s["dates"][-1]) - dt.timedelta(days=364)
    best = None
    for ds, v in zip(s["dates"], s[p]):
        if dt.date.fromisoformat(ds) <= d and v is not None:
            best = v
    return best


def sign(n):
    return "±0" if n == 0 else f"{n:+,}"


def ranking(s):
    vals = sorted(((s[p][-1], p) for p in PREFS if s[p][-1] is not None), reverse=True)
    return {p: i + 1 for i, (_, p) in enumerate(vals)}, vals


CSS = """
:root{--bg:#fbfaf7;--fg:#222;--muted:#666;--line:#e3e0d8;--accent:#c4501b;--up:#c4301b;--down:#1b6fc4;--card:#fff}
@media (prefers-color-scheme:dark){:root{--bg:#18181a;--fg:#eee;--muted:#aaa;--line:#333;--accent:#f08a50;--up:#f07060;--down:#60a8f0;--card:#222}}
*{box-sizing:border-box}body{margin:0;background:var(--bg);color:var(--fg);font:16px/1.7 system-ui,-apple-system,"Hiragino Sans","Noto Sans JP",sans-serif}
main,header,footer{max-width:860px;margin:0 auto;padding:0 16px}header{padding-top:16px}header a{color:var(--fg);text-decoration:none;font-weight:700}
h1{font-size:1.5rem;line-height:1.4;margin:.6em 0 .3em}h2{font-size:1.15rem;margin:1.8em 0 .5em;border-left:4px solid var(--accent);padding-left:8px}
.big{font-size:2.2rem;font-weight:700}.muted{color:var(--muted);font-size:.9rem}.up{color:var(--up)}.down{color:var(--down)}
.cards{display:grid;grid-template-columns:repeat(auto-fit,minmax(150px,1fr));gap:10px}.card{background:var(--card);border:1px solid var(--line);border-radius:8px;padding:10px 12px}
.card b{display:block;font-size:1.3rem}table{width:100%;border-collapse:collapse;font-size:.95rem}th,td{padding:6px 8px;border-bottom:1px solid var(--line);text-align:right}
th:first-child,td:first-child,th:nth-child(2),td:nth-child(2){text-align:left}a{color:var(--accent)}.chart{position:relative;height:300px}
.prefs{display:flex;flex-wrap:wrap;gap:6px 12px}footer{margin-top:40px;padding-bottom:40px;font-size:.85rem;color:var(--muted);border-top:1px solid var(--line)}
"""


def page(title, desc, body, path, extra_head=""):
    return f"""<!doctype html><html lang="ja"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width,initial-scale=1"><title>{html.escape(title)}</title>
<meta name="description" content="{html.escape(desc)}"><link rel="canonical" href="{SITE_URL}{path}">
<style>{CSS}</style>{extra_head}</head><body>
<header><a href="/">{SITE_NAME}</a></header><main>{body}</main>
<footer><p>出典：資源エネルギー庁「石油製品価格調査」（<a href="{BASE_URL}results.html">週次ファイル</a>）を加工して作成。
数値は各週月曜時点の消費税込みの小売価格。灯油は18リットルあたり、ガソリン・軽油は1リットルあたり。</p>
<p><a href="/about.html">このサイトについて</a></p></footer></body></html>"""


CHART_JS = '<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.4/dist/chart.umd.min.js"></script>'


def chart_script(cid, labels, datasets):
    return f"""<script>new Chart(document.getElementById("{cid}"),{{type:"line",
data:{{labels:{json.dumps(labels)},datasets:{json.dumps(datasets, ensure_ascii=False)}}},
options:{{maintainAspectRatio:false,spanGaps:true,pointRadius:0,interaction:{{mode:"index",intersect:false}},
elements:{{point:{{radius:0}},line:{{borderWidth:2}}}},scales:{{x:{{ticks:{{maxTicksLimit:8}}}}}}}}}});</script>"""


def winter_table(s, p):
    """毎年の冬（11〜3月）の平均と最高。"""
    rows = []
    years = sorted({int(d[:4]) for d in s["dates"]})
    for y in years:
        vs = [v for d, v in zip(s["dates"], s[p]) if v is not None and
              (d[:7] >= f"{y}-11" and d[:7] <= f"{y + 1}-03")]
        if len(vs) >= 15:
            rows.append((f"{y}年冬", round(sum(vs) / len(vs)), max(vs)))
    return rows[-6:]


def build(data):
    SITE.mkdir(exist_ok=True)
    k = data["kerosene"]
    latest = k["dates"][-1]
    rank, vals = ranking(k)
    nat, nat_prev, nat_ya = last(k["全国"]), last(k["全国"], 1), year_ago(k, "全国")
    upd = dt.date.fromisoformat(latest)

    # トップ
    rows = []
    for v, p in vals:
        d = v - (k[p][-2] or v)
        ya = year_ago(k, p)
        rows.append(f'<tr><td>{rank[p]}</td><td><a href="/{SLUG[p]}.html">{FULL[p]}</a></td><td>{v:,}円</td>'
                    f'<td class="{"up" if d > 0 else "down" if d < 0 else ""}">{sign(d)}</td>'
                    f'<td>{sign(v - ya) if ya else "-"}</td></tr>')
    tail = -260  # 約5年
    ds = [{"label": "全国 灯油（店頭）", "data": k["全国"][tail:], "borderColor": "#c4501b"}]
    body = f"""<h1>灯油18Lの値段、今週の都道府県ランキング</h1>
<p class="muted">{upd.month}月{upd.day}日（月）時点。毎週水曜ごろ、国の調査が出たら更新します。</p>
<div class="cards"><div class="card">全国平均<b>{nat:,}円</b><span class="muted">18Lあたり</span></div>
<div class="card">先週から<b class="{"up" if nat > nat_prev else "down"}">{sign(nat - nat_prev)}円</b></div>
<div class="card">1年前から<b>{sign(nat - nat_ya)}円</b></div>
<div class="card">一番高い<b>{FULL[vals[0][1]]}</b><span class="muted">{vals[0][0]:,}円</span></div>
<div class="card">一番安い<b>{FULL[vals[-1][1]]}</b><span class="muted">{vals[-1][0]:,}円</span></div></div>
<h2>全国平均の推移（5年）</h2><div class="chart"><canvas id="c"></canvas></div>
<h2>都道府県ランキング（高い順）</h2>
<table><tr><th>順位</th><th>都道府県</th><th>18L</th><th>先週比</th><th>1年前比</th></tr>{"".join(rows)}</table>
{chart_script("c", k["dates"][tail:], ds)}"""
    (SITE / "index.html").write_text(page(f"灯油の値段 都道府県ランキング（{upd.month}/{upd.day}時点）| {SITE_NAME}",
        f"灯油18Lの店頭価格を都道府県別に毎週更新。全国平均{nat:,}円、最高は{FULL[vals[0][1]]}。先週比・1年前比・5年の推移グラフ。",
        body, "/", CHART_JS), encoding="utf-8")

    # 都道府県
    for p in PREFS:
        v, prev = last(k[p]), last(k[p], 1)
        ya = year_ago(k, p)
        r = data["regular"]
        rv, rprev = last(r[p]), last(r[p], 1)
        rr, _ = ranking(r)
        diff_nat = v - nat
        hist = [x for x in k[p] if x is not None]
        hi = max(hist[-260:])
        hi_d = k["dates"][len(k[p]) - 260 + k[p][-260:].index(hi)]
        ds1 = [{"label": f"{FULL[p]} 店頭", "data": k[p][tail:], "borderColor": "#c4501b"},
               {"label": f"{FULL[p]} 配達", "data": data["kerosene_dlv"][p][tail:], "borderColor": "#d9a03a"},
               {"label": "全国 店頭", "data": k["全国"][tail:], "borderColor": "#999", "borderDash": [4, 4]}]
        ds2 = [{"label": f"{FULL[p]} レギュラー", "data": r[p][tail:], "borderColor": "#1b6fc4"},
               {"label": "全国 レギュラー", "data": r["全国"][tail:], "borderColor": "#999", "borderDash": [4, 4]}]
        wt = "".join(f"<tr><td>{a}</td><td></td><td>{b:,}円</td><td>{c:,}円</td></tr>" for a, b, c in winter_table(k, p))
        near = sorted(PREFS, key=lambda q: abs(PREFS.index(q) - PREFS.index(p)))[1:7]
        body = f"""<h1>{FULL[p]}の灯油の値段（18L）とガソリン価格の推移</h1>
<p class="muted">{upd.month}月{upd.day}日（月）時点。</p>
<div class="cards"><div class="card">灯油 店頭<b>{v:,}円</b><span class="muted">18Lあたり</span></div>
<div class="card">先週から<b class="{"up" if v > prev else "down" if v < prev else ""}">{sign(v - prev)}円</b></div>
<div class="card">1年前から<b>{sign(v - ya) if ya else "-"}円</b></div>
<div class="card">全国で<b>{rank[p]}番目に高い</b><span class="muted">全国平均と{sign(diff_nat)}円</span></div>
<div class="card">レギュラー<b>{rv:,}円/L</b><span class="muted">先週{sign(rv - rprev)}円・高い順{rr[p]}位</span></div></div>
<p>{FULL[p]}の灯油は、18Lで{v:,}円です。全国平均（{nat:,}円）より{abs(diff_nat):,}円{"高く" if diff_nat > 0 else "安く" if diff_nat < 0 else "同じで"}、
47都道府県で{rank[p]}番目に高い値段です。この5年で一番高かったのは{hi_d[:4]}年{int(hi_d[5:7])}月の{hi:,}円でした。</p>
<h2>灯油の推移（5年・店頭と配達）</h2><div class="chart"><canvas id="c1"></canvas></div>
<h2>冬ごとの灯油の値段（11〜3月）</h2>
<table><tr><th>冬</th><th></th><th>平均</th><th>最高</th></tr>{wt}</table>
<h2>レギュラーガソリンの推移（5年）</h2><div class="chart"><canvas id="c2"></canvas></div>
<h2>近くの県</h2><div class="prefs">{"".join(f'<a href="/{SLUG[q]}.html">{FULL[q]}</a>' for q in near)}</div>
<h2>ほかの都道府県</h2><div class="prefs">{"".join(f'<a href="/{SLUG[q]}.html">{q}</a>' for q in PREFS if q != p)}</div>
{chart_script("c1", k["dates"][tail:], ds1)}{chart_script("c2", r["dates"][tail:], ds2)}"""
        (SITE / f"{SLUG[p]}.html").write_text(page(
            f"{FULL[p]}の灯油価格 18L {v:,}円（{upd.month}/{upd.day}）推移と全国順位 | {SITE_NAME}",
            f"{FULL[p]}の灯油18Lは{v:,}円（{upd.month}月{upd.day}日）。先週比{sign(v - prev)}円、全国{rank[p]}位。店頭・配達・ガソリンの5年推移。",
            body, f"/{SLUG[p]}.html", CHART_JS), encoding="utf-8")

    about = """<h1>このサイトについて</h1>
<p>資源エネルギー庁が毎週公表している「石油製品価格調査」の数字を、都道府県ごとに見やすくまとめたサイトです。
国の調査が出たら（通常は水曜）、自動で更新します。</p>
<p>数値は国の調査をもとにこちらで集計・グラフ化したもので、個々のお店の値段とは違います。
調査の方法は<a href="https://www.enecho.meti.go.jp/statistics/petroleum_and_lpgas/pl007/">資源エネルギー庁のページ</a>をご覧ください。</p>"""
    (SITE / "about.html").write_text(page(f"このサイトについて | {SITE_NAME}", "データの出典と更新について", about, "/about.html"),
                                     encoding="utf-8")
    urls = ["/", "/about.html"] + [f"/{SLUG[p]}.html" for p in PREFS]
    (SITE / "sitemap.xml").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n<urlset xmlns="http://www.sitemaps.org/schemas/sitemap/0.9">\n'
        + "".join(f"<url><loc>{SITE_URL}{u}</loc><lastmod>{latest}</lastmod></url>\n" for u in urls)
        + "</urlset>\n", encoding="utf-8")
    (SITE / "robots.txt").write_text(f"User-agent: *\nAllow: /\nSitemap: {SITE_URL}/sitemap.xml\n", encoding="utf-8")
    print(f"{latest} 時点で {len(PREFS) + 2} ページ書き出し → {SITE}")


if __name__ == "__main__":
    if "--no-fetch" not in sys.argv:
        fetch()
    build(load())
