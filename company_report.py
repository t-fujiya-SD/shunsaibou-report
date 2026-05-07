#!/usr/bin/env python3
"""
旬彩坊 企業別HTMLレポート生成スクリプト
使い方: python3 company_report.py
出力: output/report_{企業名}.html (8ファイル)
"""

import pandas as pd
from pathlib import Path
import json
from datetime import datetime

BASE_DIR     = Path(__file__).parent
OUTPUT_DIR   = BASE_DIR / "output"
SUMMARY_PATH = OUTPUT_DIR / "weekly_summary.xlsx"
STORE_LIST   = BASE_DIR / "master" / "主要店舗リスト.csv"

# 企業名 → ファイル名の対応
COMPANY_FILES = {
    "株式会社キープウィルダイニング": "キープウィルダイニング",
    "株式会社ゴリラカンパニー":       "ゴリラカンパニー",
    "アイティープラス":               "アイティープラス",
    "いせ久":                         "いせ久",
    "Ｖａｍｏ株式会社":               "Vamo",
    "望滇山":                         "望滇山",
    "株式会社ジャパンダイニング":     "ジャパンダイニング",
    "いそいそグループ":               "いそいそグループ",
}

SOURCE_COLORS = {
    "インフォマート": "#4e9af1",
    "タノム":         "#6abf69",
    "アスピット":     "#ffca28",
    "販売大臣":       "#f06292",
}

def load_summary():
    df = pd.read_excel(SUMMARY_PATH, sheet_name="店舗別週次サマリ")
    df = df[df["集計週"].notna() & (df["ソース"] != "【合計】")].copy()
    df["合計金額（円）"] = pd.to_numeric(df["合計金額（円）"], errors="coerce").fillna(0).astype(int)
    return df

def load_detail():
    df = pd.read_excel(SUMMARY_PATH, sheet_name="全明細")
    df["金額"] = pd.to_numeric(df["金額"], errors="coerce").fillna(0).astype(int)
    df["単価"] = pd.to_numeric(df["単価"], errors="coerce").fillna(0)
    df["数量"] = pd.to_numeric(df["数量"], errors="coerce").fillna(0)
    df["取引日"] = pd.to_datetime(df["取引日"], errors="coerce")
    df = df.dropna(subset=["取引日"])
    return df

def load_store_map():
    if not STORE_LIST.exists():
        return {}
    df = pd.read_csv(STORE_LIST, dtype=str)
    result = {}
    for _, row in df.iterrows():
        key = str(row["店舗名"]).strip().replace("\u3000", " ")
        val = str(row["企業名"]).strip()
        result[key] = val
    return result

def week_to_month(week_str):
    year, w = week_str.split("-W")
    monday = datetime.fromisocalendar(int(year), int(w), 1)
    return f"{monday.year}年{monday.month}月"

def week_to_label(week_str):
    """2026-W18 → '4月27日〜'"""
    year, w = week_str.split("-W")
    monday = datetime.fromisocalendar(int(year), int(w), 1)
    return f"{monday.month}月{monday.day}日〜"

def fmt_yen(val):
    return f"¥{val:,}"

def build_company_data(df_summary, df_detail, store_map, company):
    """企業1社分のデータを構築"""
    # 企業に属する店舗を特定
    stores = [s for s, c in store_map.items() if c == company]

    # 週次サマリを企業でフィルタ
    df_s = df_summary.copy()
    df_s["企業名"] = df_s["店舗名"].apply(
        lambda x: store_map.get(str(x).strip().replace("\u3000", " "),
                                str(x).strip().replace("\u3000", " "))
    )
    df_s = df_s[df_s["企業名"] == company]

    # 全明細を企業でフィルタ
    df_d = df_detail.copy()
    df_d["企業名"] = df_d["店舗名"].apply(
        lambda x: store_map.get(str(x).strip().replace("\u3000", " "),
                                str(x).strip().replace("\u3000", " "))
    )
    df_d = df_d[df_d["企業名"] == company]

    if df_s.empty:
        return None

    weeks = sorted(df_s["集計週"].unique())

    # 週次データ
    weekly = {}
    for w in weeks:
        wdf = df_s[df_s["集計週"] == w]
        by_source = {}
        for src in ["インフォマート", "タノム", "アスピット", "販売大臣"]:
            by_source[src] = int(wdf[wdf["ソース"] == src]["合計金額（円）"].sum())
        by_store = {}
        for store in stores:
            by_store[store] = int(wdf[wdf["店舗名"].str.strip().str.replace("\u3000", " ") == store]["合計金額（円）"].sum())
        weekly[w] = {
            "total": int(wdf["合計金額（円）"].sum()),
            "by_source": by_source,
            "by_store": by_store,
            "month": week_to_month(w),
        }

    # 月別データ
    month_to_weeks = {}
    for w in weeks:
        m = weekly[w]["month"]
        if m not in month_to_weeks:
            month_to_weeks[m] = []
        month_to_weeks[m].append(w)

    monthly = {}
    if not df_d.empty:
        df_d["月"] = df_d["取引日"].dt.strftime("%Y年%-m月")
        df_d["日付"] = df_d["取引日"].dt.strftime("%Y-%m-%d")
        for month, mdf in df_d.groupby("月"):
            dates = sorted(mdf["日付"].unique())
            daily_totals = [int(mdf[mdf["日付"] == d]["金額"].sum()) for d in dates]
            monthly[month] = {
                "dates": dates,
                "daily_totals": daily_totals,
                "month_total": int(mdf["金額"].sum()),
            }

    # 検索データ
    records = []
    for _, r in df_d.iterrows():
        records.append({
            "date": r["取引日"].strftime("%Y-%m-%d"),
            "store": str(r["店舗名"]),
            "product": str(r["商品名"]),
            "price": float(r["単価"]) if r["単価"] else 0,
            "qty": float(r["数量"]) if r["数量"] else 0,
            "amount": int(r["金額"]),
        })

    min_date = df_d["取引日"].min().strftime("%Y-%m-%d") if not df_d.empty else ""
    max_date = df_d["取引日"].max().strftime("%Y-%m-%d") if not df_d.empty else ""

    return {
        "company": company,
        "weeks": weeks,
        "week_labels": {w: week_to_label(w) for w in weeks},
        "weekly": weekly,
        "month_to_weeks": month_to_weeks,
        "monthly": monthly,
        "stores": sorted(stores),
        "source_colors": SOURCE_COLORS,
        "search": {
            "records": records,
            "stores": sorted(set(r["store"] for r in records)),
            "min_date": min_date,
            "max_date": max_date,
        }
    }

def build_html(data, generated_at):
    company = data["company"]
    weeks = data["weeks"]
    latest_total = data["weekly"][weeks[-1]]["total"] if weeks else 0
    cumulative = sum(v["total"] for v in data["weekly"].values())
    stores = data["stores"]
    # 前週比
    if len(weeks) >= 2:
        prev_total = data["weekly"][weeks[-2]]["total"]
        wow_pct = (latest_total - prev_total) / prev_total * 100 if prev_total else 0
        wow_str = f"{wow_pct:+.1f}%"
        wow_color = "#27AE60" if wow_pct >= 0 else "#E74C3C"
    else:
        wow_str = "－"
        wow_color = "#888"
    store_colors = ["#4e9af1", "#f06292", "#ffca28", "#6abf69", "#ab47bc", "#26c6da", "#ff7043", "#ff8f00"]

    data_json   = json.dumps(data, ensure_ascii=False)
    search_json = json.dumps(data["search"], ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>{company} 仕入れレポート</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Hiragino Sans', 'Meiryo', sans-serif; background: #f5f6fa; color: #333; }}
  header {{ background: #1a5276; color: white; padding: 24px 32px; }}
  header h1 {{ font-size: 20px; font-weight: bold; }}
  header p {{ font-size: 12px; opacity: 0.7; margin-top: 4px; }}
  .container {{ max-width: 1100px; margin: 32px auto; padding: 0 24px; }}
  .card {{ background: white; border-radius: 12px; padding: 28px; margin-bottom: 28px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); }}
  .card h2 {{ font-size: 16px; font-weight: bold; color: #1a5276; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 2px solid #ecf0f1; }}
  .chart-wrap {{ position: relative; height: 320px; }}
  .chart-wrap-tall {{ position: relative; height: 360px; }}
  .chart-wrap-xl {{ position: relative; height: 960px; }}
  .summary-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 28px; }}
  .summary-card {{ background: white; border-radius: 12px; padding: 20px 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); }}
  .summary-card .label {{ font-size: 12px; color: #888; margin-bottom: 6px; }}
  .summary-card .value {{ font-size: 24px; font-weight: bold; color: #1a5276; }}
  .nav-bar {{ background: white; border-radius: 12px; padding: 16px 24px; margin-bottom: 28px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }}
  .nav-bar span {{ font-size: 13px; color: #666; font-weight: bold; }}
  .nav-btn {{ padding: 8px 20px; border-radius: 20px; border: 2px solid #1a5276; background: white; color: #1a5276; font-size: 13px; cursor: pointer; transition: all 0.2s; font-family: inherit; }}
  .nav-btn:hover {{ background: #1a5276; color: white; }}
  .nav-btn.active {{ background: #1a5276; color: white; }}
  .month-select {{ padding: 8px 16px; border-radius: 20px; border: 2px solid #E67E22; color: #E67E22; background: white; font-size: 13px; cursor: pointer; font-family: inherit; }}
  .search-btn-nav {{ border-color: #8E44AD; color: #8E44AD; }}
  .search-btn-nav:hover, .search-btn-nav.active {{ background: #8E44AD; color: white; border-color: #8E44AD; }}
  .divider {{ width: 1px; background: #ddd; align-self: stretch; margin: 0 4px; }}
  .search-box {{ background: white; border-radius: 12px; padding: 24px 28px; margin-bottom: 28px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); }}
  .search-row {{ display: flex; align-items: center; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; }}
  .search-row label {{ font-size: 13px; color: #666; font-weight: bold; min-width: 60px; }}
  .search-row input, .search-row select {{ padding: 8px 12px; border: 1px solid #ddd; border-radius: 8px; font-size: 13px; font-family: inherit; }}
  .search-input {{ width: 280px; }}
  .search-input-wrap {{ position: relative; }}
  .store-suggestions {{ position: absolute; background: white; border: 1px solid #ddd; border-radius: 8px; max-height: 200px; overflow-y: auto; z-index: 100; width: 280px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }}
  .store-suggestions div {{ padding: 8px 12px; cursor: pointer; font-size: 13px; }}
  .store-suggestions div:hover {{ background: #f0f4ff; }}
  .search-btn-run {{ padding: 9px 24px; background: #8E44AD; color: white; border: none; border-radius: 20px; font-size: 13px; cursor: pointer; font-family: inherit; }}
  .search-btn-run:hover {{ background: #7D3C98; }}
  .csv-btn {{ padding: 9px 24px; background: #27AE60; color: white; border: none; border-radius: 20px; font-size: 13px; cursor: pointer; font-family: inherit; display: none; }}
  .result-table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  .result-table th {{ background: #1a5276; color: white; padding: 10px 12px; text-align: left; }}
  .result-table th.num {{ text-align: right; }}
  .result-table td {{ padding: 9px 12px; border-bottom: 1px solid #ecf0f1; }}
  .result-table td.num {{ text-align: right; }}
  .result-table tr:hover {{ background: #f8f9fa; }}
  .result-table tr:nth-child(even) {{ background: #fafbfc; }}
  .result-summary {{ font-size: 13px; color: #666; margin-bottom: 12px; }}
  .store-filter {{ margin-bottom: 16px; }}
  .store-filter-bar {{ display: flex; align-items: center; gap: 10px; margin-bottom: 10px; flex-wrap: wrap; }}
  .store-filter-bar input[type=text] {{ padding: 6px 12px; border: 1px solid #ddd; border-radius: 8px; font-size: 13px; font-family: inherit; width: 240px; }}
  .filter-btn {{ padding: 5px 14px; border-radius: 16px; border: 1px solid #1a5276; background: white; color: #1a5276; font-size: 12px; cursor: pointer; font-family: inherit; }}
  .filter-btn:hover {{ background: #1a5276; color: white; }}
  .store-checkbox-list {{ display: flex; flex-wrap: wrap; gap: 6px; max-height: 120px; overflow-y: auto; padding: 8px; background: #f9f9f9; border-radius: 8px; border: 1px solid #eee; }}
  .store-checkbox-list label {{ display: flex; align-items: center; gap: 5px; font-size: 12px; cursor: pointer; white-space: nowrap; padding: 3px 8px; background: white; border-radius: 12px; border: 1px solid #ddd; }}
  .store-checkbox-list label:hover {{ border-color: #1a5276; }}
  .store-checkbox-list input[type=checkbox] {{ cursor: pointer; }}
</style>
</head>
<body>
<header>
  <h1>{company}　仕入れレポート</h1>
  <p>生成日時: {generated_at}　　旬彩坊 提供</p>
</header>
<div class="container">

  <div class="summary-grid">
    <div class="summary-card">
      <div class="label">最新週</div>
      <div class="value">{weeks[-1] if weeks else '-'}</div>
    </div>
    <div class="summary-card">
      <div class="label">最新週 仕入れ合計</div>
      <div class="value">{fmt_yen(latest_total)}</div>
    </div>
    <div class="summary-card">
      <div class="label">前週比</div>
      <div class="value" style="color:{wow_color}">{wow_str}</div>
    </div>
    <div class="summary-card">
      <div class="label">累計仕入れ</div>
      <div class="value">{fmt_yen(cumulative)}</div>
    </div>
  </div>

  <div class="nav-bar">
    <button class="nav-btn active" id="weekly-btn" onclick="showWeeklyView()">週次取引</button>
    <div class="divider"></div>
    <span>月別：</span>
    <select class="month-select" id="month-select" onchange="onMonthSelect(this)"></select>
    <div class="divider"></div>
    <button class="nav-btn search-btn-nav" id="search-toggle" onclick="showSearchView()">🔍 明細検索</button>
  </div>

  <!-- 週次ビュー -->
  <div id="weekly-view">
    <div class="card">
      <h2>① 週次仕入れ合計推移</h2>
      <div class="chart-wrap"><canvas id="chart1"></canvas></div>
    </div>
    <div class="card">
      <h2>② 店舗別の週次推移</h2>
      <div class="store-filter">
        <div class="store-filter-bar">
          <input type="text" id="store-filter-input" placeholder="店舗名で絞り込み..." oninput="filterStoreCheckboxes()">
          <button class="filter-btn" onclick="selectAllStores(true)">全選択</button>
          <button class="filter-btn" onclick="selectAllStores(false)">全解除</button>
        </div>
        <div class="store-checkbox-list" id="store-checkbox-list"></div>
      </div>
      <div class="chart-wrap-xl"><canvas id="chart3"></canvas></div>
    </div>
  </div>

  <!-- 月次ビュー -->
  <div id="monthly-view" style="display:none;">
    <div class="card">
      <h2 id="monthly-title1">① 日次仕入れ</h2>
      <div class="chart-wrap-tall"><canvas id="chart_m1"></canvas></div>
    </div>
  </div>

  <!-- 検索ビュー -->
  <div id="search-view" style="display:none;">
    <div class="search-box">
      <div class="search-row">
        <label>期間</label>
        <input type="date" id="date-from" />
        <span style="color:#999;">〜</span>
        <input type="date" id="date-to" />
      </div>
      <div class="search-row">
        <label>店舗</label>
        <div class="search-input-wrap">
          <input type="text" id="store-input" class="search-input" placeholder="店舗名を入力（空欄で全店舗）" autocomplete="off" oninput="filterSuggestions()" onblur="hideSuggestions()" />
          <div class="store-suggestions" id="suggestions" style="display:none;"></div>
        </div>
        <select id="store-select" style="padding:8px 12px;border:1px solid #ddd;border-radius:8px;font-size:13px;font-family:inherit;min-width:200px;">
          <option value="">（プルダウンで選択）</option>
        </select>
      </div>
      <div class="search-row">
        <button class="search-btn-run" onclick="runSearch()">検索</button>
        <button class="csv-btn" id="csv-btn" onclick="downloadCSV()">CSVダウンロード</button>
      </div>
    </div>
    <div class="card">
      <h2 id="search-result-title">商品別明細</h2>
      <div class="result-summary" id="result-summary"></div>
      <div style="overflow-x:auto;">
        <table class="result-table">
          <thead>
            <tr>
              <th>商品名</th>
              <th class="num">単価</th>
              <th class="num">数量</th>
              <th class="num">仕入れ金額</th>
            </tr>
          </thead>
          <tbody id="result-body"></tbody>
        </table>
      </div>
    </div>
  </div>

</div>
<script>
Chart.register(ChartDataLabels);
const ALL_DATA = {data_json};
const SEARCH_DATA = {search_json};
const weeks = ALL_DATA.weeks;
const weekly = ALL_DATA.weekly;
const monthly = ALL_DATA.monthly;
const stores = ALL_DATA.stores;
const storeColors = {json.dumps(store_colors[:len(stores)], ensure_ascii=False)};

// 月別プルダウン初期化
const monthSelect = document.getElementById('month-select');
const months = Object.keys(ALL_DATA.month_to_weeks);
months.forEach(m => {{
  const opt = document.createElement('option');
  opt.value = m; opt.textContent = m;
  monthSelect.appendChild(opt);
}});
if (months.length > 0) monthSelect.value = months[months.length - 1];

// 店舗プルダウン初期化
const storeSelect = document.getElementById('store-select');
SEARCH_DATA.stores.forEach(s => {{
  const opt = document.createElement('option');
  opt.value = s; opt.textContent = s;
  storeSelect.appendChild(opt);
}});
storeSelect.addEventListener('change', () => {{
  document.getElementById('store-input').value = storeSelect.value;
}});
document.getElementById('date-from').value = SEARCH_DATA.min_date;
document.getElementById('date-to').value = SEARCH_DATA.max_date;

// ナビ切替
function setNav(id) {{
  document.querySelectorAll('.nav-btn').forEach(b => b.classList.remove('active'));
  if (id) document.getElementById(id).classList.add('active');
}}
function showWeeklyView() {{
  document.getElementById('weekly-view').style.display = 'block';
  document.getElementById('monthly-view').style.display = 'none';
  document.getElementById('search-view').style.display = 'none';
  setNav('weekly-btn');
  updateWeeklyCharts(weeks);
}}
function onMonthSelect(sel) {{
  const month = sel.value;
  if (!month) return;
  document.getElementById('weekly-view').style.display = 'none';
  document.getElementById('monthly-view').style.display = 'block';
  document.getElementById('search-view').style.display = 'none';
  setNav(null);
  updateMonthlyCharts(month);
}}
function showSearchView() {{
  document.getElementById('weekly-view').style.display = 'none';
  document.getElementById('monthly-view').style.display = 'none';
  document.getElementById('search-view').style.display = 'block';
  setNav('search-toggle');
  runSearch();
}}

// チャート設定
const datalabels = {{
  display: true, align: 'top', anchor: 'end',
  font: {{ size: 10 }},
  formatter: v => v > 0 ? '¥' + Math.round(v/10000) + '万' : '',
  color: '#555',
}};
const mkOpts = (legend) => ({{
  responsive: true, maintainAspectRatio: false,
  plugins: {{
    legend: legend ? {{ position: 'bottom', labels: {{ boxWidth: 12, font: {{ size: 11 }} }} }} : {{ display: false }},
    datalabels,
  }},
  scales: {{ y: {{ ticks: {{ callback: v => '¥' + v.toLocaleString() }} }} }}
}});
const barOpts = (legend) => ({{
  responsive: true, maintainAspectRatio: false,
  plugins: {{
    legend: legend ? {{ position: 'bottom', labels: {{ boxWidth: 12, font: {{ size: 11 }} }} }} : {{ display: false }},
    datalabels: {{ display: true, anchor: 'end', align: 'top', font: {{ size: 10 }},
      formatter: v => v > 0 ? '¥' + Math.round(v/10000) + '万' : '', color: '#555' }},
  }},
  scales: {{ y: {{ ticks: {{ callback: v => '¥' + v.toLocaleString() }} }} }}
}});

// 週次チャート
const chart1 = new Chart(document.getElementById('chart1'), {{
  type: 'line',
  data: {{ labels: [], datasets: [{{ label: '仕入れ合計', borderColor: '#1a5276', backgroundColor: '#1a527633', tension: 0.3, fill: true, pointRadius: 6, data: [] }}] }},
  options: mkOpts(false),
}});
const chart3 = new Chart(document.getElementById('chart3'), {{
  type: 'line',
  data: {{ labels: [], datasets: stores.map((s, i) => ({{
    label: s, borderColor: storeColors[i % storeColors.length],
    backgroundColor: storeColors[i % storeColors.length] + '33',
    tension: 0.3, fill: false, pointRadius: 6, data: []
  }})) }},
  options: mkOpts(true),
}});

// 月次チャート
const chart_m1 = new Chart(document.getElementById('chart_m1'), {{
  type: 'bar',
  data: {{ labels: [], datasets: [{{ label: '日次仕入れ', backgroundColor: '#1a527688', data: [] }}] }},
  options: barOpts(false),
}});

function updateWeeklyCharts(ws) {{
  const labels = ws.map(w => ALL_DATA.week_labels[w]);
  chart1.data.labels = labels;
  chart1.data.datasets[0].data = ws.map(w => weekly[w].total);
  chart1.update();
  chart3.data.labels = labels;
  stores.forEach((s, i) => {{ chart3.data.datasets[i].data = ws.map(w => weekly[w].by_store[s] || 0); }});
  chart3.update();
}}

// 店舗フィルター
function buildStoreCheckboxes() {{
  const list = document.getElementById('store-checkbox-list');
  list.innerHTML = '';
  stores.forEach((s, i) => {{
    const label = document.createElement('label');
    const cb = document.createElement('input');
    cb.type = 'checkbox';
    cb.checked = true;
    cb.dataset.idx = i;
    cb.onchange = () => applyStoreFilter();
    const dot = document.createElement('span');
    dot.style.cssText = `display:inline-block;width:8px;height:8px;border-radius:50%;background:${{storeColors[i % storeColors.length]}};`;
    label.appendChild(cb);
    label.appendChild(dot);
    label.appendChild(document.createTextNode(' ' + s));
    list.appendChild(label);
  }});
}}

function filterStoreCheckboxes() {{
  const val = document.getElementById('store-filter-input').value.toLowerCase();
  document.querySelectorAll('#store-checkbox-list label').forEach(label => {{
    label.style.display = label.textContent.toLowerCase().includes(val) ? '' : 'none';
  }});
}}

function selectAllStores(checked) {{
  document.querySelectorAll('#store-checkbox-list input[type=checkbox]').forEach(cb => {{
    if (cb.closest('label').style.display !== 'none') cb.checked = checked;
  }});
  applyStoreFilter();
}}

function applyStoreFilter() {{
  document.querySelectorAll('#store-checkbox-list input[type=checkbox]').forEach(cb => {{
    chart3.data.datasets[cb.dataset.idx].hidden = !cb.checked;
  }});
  chart3.update();
}}

function updateMonthlyCharts(month) {{
  const md = monthly[month];
  if (!md) return;
  document.getElementById('monthly-title1').textContent = '① ' + month + ' 日次仕入れ';
  chart_m1.data.labels = md.dates;
  chart_m1.data.datasets[0].data = md.daily_totals;
  chart_m1.update();
}}

// 検索UI
function filterSuggestions() {{
  const val = document.getElementById('store-input').value.toLowerCase();
  const box = document.getElementById('suggestions');
  if (!val) {{ box.style.display = 'none'; return; }}
  const matches = SEARCH_DATA.stores.filter(s => s.toLowerCase().includes(val)).slice(0, 15);
  if (!matches.length) {{ box.style.display = 'none'; return; }}
  box.innerHTML = '';
  matches.forEach(s => {{
    const d = document.createElement('div');
    d.textContent = s;
    d.onmousedown = () => {{ document.getElementById('store-input').value = s; box.style.display = 'none'; }};
    box.appendChild(d);
  }});
  box.style.display = 'block';
}}
function hideSuggestions() {{
  setTimeout(() => {{ document.getElementById('suggestions').style.display = 'none'; }}, 150);
}}

function runSearch() {{
  const from  = document.getElementById('date-from').value;
  const to    = document.getElementById('date-to').value;
  const store = document.getElementById('store-input').value.trim();
  const filtered = SEARCH_DATA.records.filter(r => {{
    if (from && r.date < from) return false;
    if (to   && r.date > to)   return false;
    if (store && !r.store.includes(store)) return false;
    return true;
  }});
  const label = store || '全店舗';
  document.getElementById('search-result-title').textContent =
    label + '　' + from + ' 〜 ' + to + '　商品別明細';
  const agg = {{}};
  filtered.forEach(r => {{
    if (!agg[r.product]) agg[r.product] = {{price: r.price, qty: 0, amount: 0}};
    agg[r.product].qty    += r.qty;
    agg[r.product].amount += r.amount;
  }});
  const rows = Object.entries(agg)
    .map(([name, v]) => ({{name, price: v.price, qty: v.qty, amount: v.amount}}))
    .sort((a, b) => b.amount - a.amount);
  const total = rows.reduce((s, r) => s + r.amount, 0);
  document.getElementById('result-summary').textContent =
    '商品数: ' + rows.length + '種　合計: ¥' + total.toLocaleString();
  const tbody = document.getElementById('result-body');
  tbody.innerHTML = '';
  rows.forEach(r => {{
    const tr = document.createElement('tr');
    const qty_str = r.qty % 1 === 0 ? r.qty.toLocaleString() : r.qty.toFixed(2);
    const price_str = r.price > 0 ? '¥' + r.price.toLocaleString() : '-';
    tr.innerHTML = '<td>' + r.name + '</td>' +
      '<td class="num">' + price_str + '</td>' +
      '<td class="num">' + qty_str + '</td>' +
      '<td class="num">¥' + r.amount.toLocaleString() + '</td>';
    tbody.appendChild(tr);
  }});
  document.getElementById('csv-btn').style.display = rows.length > 0 ? 'inline-block' : 'none';
  window._csvRows = rows;
}}

function downloadCSV() {{
  const rows = window._csvRows || [];
  const from = document.getElementById('date-from').value;
  const to   = document.getElementById('date-to').value;
  const store = document.getElementById('store-input').value.trim() || '全店舗';
  const sep = String.fromCharCode(13, 10);
  const lines = ['\\uFEFF商品名,単価,数量,仕入れ金額'];
  rows.forEach(r => {{
    const qty = r.qty % 1 === 0 ? r.qty : r.qty.toFixed(2);
    const name = '"' + r.name.replace(/"/g, '""') + '"';
    lines.push(name + ',' + r.price + ',' + qty + ',' + r.amount);
  }});
  const csv = lines.join(sep);
  const blob = new Blob([csv], {{type: 'text/csv;charset=utf-8;'}});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a');
  a.href = url;
  a.download = store + '_' + from + '_' + to + '.csv';
  a.click();
  URL.revokeObjectURL(url);
}}

// 初期表示
buildStoreCheckboxes();
showWeeklyView();
</script>
</body>
</html>"""

def main():
    if not SUMMARY_PATH.exists():
        print("[ERROR] weekly_summary.xlsx が見つかりません。先に aggregate.py を実行してください。")
        return

    print("■ データ読み込み中...")
    df_summary = load_summary()
    df_detail  = load_detail()
    store_map  = load_store_map()
    print(f"  集計週: {sorted(df_summary['集計週'].unique())}")
    print(f"  明細行数: {len(df_detail):,}行")

    OUTPUT_DIR.mkdir(exist_ok=True)
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")

    print("■ 企業別HTML生成中...")
    for company, filename in COMPANY_FILES.items():
        data = build_company_data(df_summary, df_detail, store_map, company)
        if data is None:
            print(f"  ⚠️  {company}: データなし（スキップ）")
            continue
        html = build_html(data, generated_at)
        out_path = OUTPUT_DIR / f"report_{filename}.html"
        out_path.write_text(html, encoding="utf-8")
        total = sum(v["total"] for v in data["weekly"].values())
        print(f"  ✅ {company}: ¥{total:,}円 → {out_path.name}")

    print("\n✅ 全企業のHTMLを output/ に出力しました。")

if __name__ == "__main__":
    main()
