#!/usr/bin/env python3
"""
旬彩坊 週次集計 HTMLレポート生成スクリプト v2
使い方: python3 report.py
出力: output/report.html
"""

import pandas as pd
from pathlib import Path
import json
from datetime import datetime

BASE_DIR     = Path(__file__).parent
OUTPUT_DIR   = BASE_DIR / "output"
OUTPUT_PATH  = OUTPUT_DIR / "report.html"
SUMMARY_PATH = OUTPUT_DIR / "weekly_summary.xlsx"
STORE_LIST   = BASE_DIR / "master" / "主要店舗リスト.csv"

LARGE_COMPANIES = [
    "株式会社キープウィルダイニング",
    "株式会社ゴリラカンパニー",
    "アイティープラス",
    "いせ久",
]
SMALL_COMPANIES = [
    "Ｖａｍｏ株式会社",
    "望滇山",
    "株式会社ジャパンダイニング",
    "いそいそグループ",
]

PALETTE_LARGE = ["#4e9af1", "#f06292", "#ffca28", "#6abf69"]
PALETTE_SMALL = ["#6abf69", "#ab47bc", "#26c6da", "#ff7043"]
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
        print(f"[WARN] 主要店舗リストが見つかりません: {STORE_LIST}")
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

def build_weekly_data(df, store_map):
    weeks = sorted(df["集計週"].unique())
    df = df.copy()
    df["企業名"] = df["店舗名"].apply(
        lambda x: store_map.get(str(x).strip().replace("\u3000", " "),
                                str(x).strip().replace("\u3000", " "))
    )
    weekly = {}
    for w in weeks:
        wdf = df[df["集計週"] == w]
        by_source = {}
        for src in ["インフォマート", "タノム", "アスピット", "販売大臣"]:
            by_source[src] = int(wdf[wdf["ソース"] == src]["合計金額（円）"].sum())
        by_company = {}
        for company in LARGE_COMPANIES + SMALL_COMPANIES:
            by_company[company] = int(wdf[wdf["企業名"] == company]["合計金額（円）"].sum())
        weekly[w] = {
            "total": int(wdf["合計金額（円）"].sum()),
            "by_source": by_source,
            "by_company": by_company,
            "month": week_to_month(w),
        }
    month_to_weeks = {}
    for w in weeks:
        m = weekly[w]["month"]
        if m not in month_to_weeks:
            month_to_weeks[m] = []
        month_to_weeks[m].append(w)
    return weeks, weekly, month_to_weeks

def build_monthly_data(detail_df, store_map):
    """月別の日次データを構築"""
    detail_df = detail_df.copy()
    detail_df["企業名"] = detail_df["店舗名"].apply(
        lambda x: store_map.get(str(x).strip().replace("\u3000", " "),
                                str(x).strip().replace("\u3000", " "))
    )
    detail_df["月"] = detail_df["取引日"].dt.strftime("%Y年%-m月")
    detail_df["日付"] = detail_df["取引日"].dt.strftime("%Y-%m-%d")

    monthly = {}
    for month, mdf in detail_df.groupby("月"):
        dates = sorted(mdf["日付"].unique())
        daily_totals = []
        for d in dates:
            daily_totals.append(int(mdf[mdf["日付"] == d]["金額"].sum()))

        large_totals = {}
        for company in LARGE_COMPANIES:
            large_totals[company] = int(mdf[mdf["企業名"] == company]["金額"].sum())

        small_totals = {}
        for company in SMALL_COMPANIES:
            small_totals[company] = int(mdf[mdf["企業名"] == company]["金額"].sum())

        monthly[month] = {
            "dates": dates,
            "daily_totals": daily_totals,
            "large_totals": large_totals,
            "small_totals": small_totals,
            "month_total": int(mdf["金額"].sum()),
        }
    return monthly

def build_search_data(detail_df, store_map):
    """検索UI用データ: 全明細を日付文字列に変換してリスト化"""
    records = []
    for _, r in detail_df.iterrows():
        store = str(r["店舗名"])
        company = store_map.get(store.replace("　", " "), store)
        records.append({
            "date": r["取引日"].strftime("%Y-%m-%d"),
            "store": store,
            "company": company,
            "product": str(r["商品名"]),
            "price": float(r["単価"]) if r["単価"] else 0,
            "qty": float(r["数量"]) if r["数量"] else 0,
            "amount": int(r["金額"]),
        })
    stores = sorted(detail_df["店舗名"].dropna().unique().tolist())
    min_date = detail_df["取引日"].min().strftime("%Y-%m-%d")
    max_date = detail_df["取引日"].max().strftime("%Y-%m-%d")
    # 企業→店舗リストのマッピング（既知企業のみ）
    known_companies = LARGE_COMPANIES + SMALL_COMPANIES
    company_stores = {}
    for rec in records:
        c = rec["company"]
        if c in known_companies:
            if c not in company_stores:
                company_stores[c] = []
            if rec["store"] not in company_stores[c]:
                company_stores[c].append(rec["store"])
    for c in company_stores:
        company_stores[c] = sorted(company_stores[c])
    return {"records": records, "stores": stores, "min_date": min_date, "max_date": max_date,
            "company_stores": company_stores, "known_companies": known_companies}

def fmt_yen(val):
    return f"¥{val:,}"

def build_html(all_data, search_data, generated_at):
    weeks = all_data["weeks"]
    latest_total = all_data["weekly"][weeks[-1]]["total"] if weeks else 0
    cumulative = sum(v["total"] for v in all_data["weekly"].values())
    total_weeks = len(weeks)
    data_json = json.dumps(all_data, ensure_ascii=False)
    search_json = json.dumps(search_data, ensure_ascii=False)

    return f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>旬彩坊 週次売上レポート</title>
<script src="https://cdn.jsdelivr.net/npm/chart.js@4.4.0/dist/chart.umd.min.js"></script>
<script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2.2.0/dist/chartjs-plugin-datalabels.min.js"></script>
<style>
  * {{ box-sizing: border-box; margin: 0; padding: 0; }}
  body {{ font-family: 'Hiragino Sans', 'Meiryo', sans-serif; background: #f5f6fa; color: #333; }}
  header {{ background: #2C3E50; color: white; padding: 24px 32px; }}
  header h1 {{ font-size: 22px; font-weight: bold; }}
  header p {{ font-size: 12px; opacity: 0.7; margin-top: 4px; }}
  .container {{ max-width: 1100px; margin: 32px auto; padding: 0 24px; }}
  .card {{ background: white; border-radius: 12px; padding: 28px; margin-bottom: 28px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); }}
  .card h2 {{ font-size: 16px; font-weight: bold; color: #2C3E50; margin-bottom: 20px; padding-bottom: 10px; border-bottom: 2px solid #ecf0f1; }}
  .chart-wrap {{ position: relative; height: 320px; }}
  .chart-wrap-tall {{ position: relative; height: 360px; }}
  .summary-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 16px; margin-bottom: 28px; }}
  .summary-card {{ background: white; border-radius: 12px; padding: 20px 24px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); }}
  .summary-card .label {{ font-size: 12px; color: #888; margin-bottom: 6px; }}
  .summary-card .value {{ font-size: 24px; font-weight: bold; color: #2C3E50; }}
  .period-selector {{ background: white; border-radius: 12px; padding: 20px 28px; margin-bottom: 28px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }}
  .period-selector span {{ font-size: 13px; color: #666; font-weight: bold; }}
  .divider {{ width: 1px; background: #ddd; align-self: stretch; margin: 0 4px; }}
  .period-btn {{ padding: 8px 20px; border-radius: 20px; border: 2px solid #2C3E50; background: white; color: #2C3E50; font-size: 13px; cursor: pointer; transition: all 0.2s; font-family: inherit; }}
  .period-btn:hover {{ background: #2C3E50; color: white; }}
  .period-btn.active {{ background: #2C3E50; color: white; }}
  .period-btn.month-btn {{ border-color: #E67E22; color: #E67E22; }}
  .period-btn.month-btn:hover, .period-btn.month-btn.active {{ background: #E67E22; color: white; }}
  .search-btn {{ border-color: #8E44AD; color: #8E44AD; }}
  .search-btn:hover, .search-btn.active {{ background: #8E44AD; color: white; }}
  .search-box {{ background: white; border-radius: 12px; padding: 24px 28px; margin-bottom: 28px; box-shadow: 0 2px 8px rgba(0,0,0,0.07); }}
  .search-row {{ display: flex; align-items: center; gap: 12px; flex-wrap: wrap; margin-bottom: 16px; }}
  .search-row label {{ font-size: 13px; color: #666; font-weight: bold; min-width: 60px; }}
  .search-row input, .search-row select {{ padding: 8px 12px; border: 1px solid #ddd; border-radius: 8px; font-size: 13px; font-family: inherit; }}
  .search-row select {{ min-width: 200px; }}
  .search-btn-run {{ padding: 9px 24px; background: #8E44AD; color: white; border: none; border-radius: 20px; font-size: 13px; cursor: pointer; font-family: inherit; }}
  .search-btn-run:hover {{ background: #7D3C98; }}
  .result-table {{ width: 100%; border-collapse: collapse; font-size: 13px; }}
  .result-table th {{ background: #2C3E50; color: white; padding: 10px 12px; text-align: left; }}
  .result-table th.num {{ text-align: right; }}
  .result-table td {{ padding: 9px 12px; border-bottom: 1px solid #ecf0f1; }}
  .result-table td.num {{ text-align: right; }}
  .result-table tr:hover {{ background: #f8f9fa; }}
  .result-table tr:nth-child(even) {{ background: #fafbfc; }}
  .result-table tr:nth-child(even):hover {{ background: #f0f2f5; }}
  .result-summary {{ font-size: 13px; color: #666; margin-bottom: 12px; }}
  .search-input {{ padding: 8px 12px; border: 1px solid #ddd; border-radius: 8px; font-size: 13px; font-family: inherit; width: 280px; }}
  .store-suggestions {{ position: absolute; background: white; border: 1px solid #ddd; border-radius: 8px; max-height: 200px; overflow-y: auto; z-index: 100; width: 280px; box-shadow: 0 4px 12px rgba(0,0,0,0.1); }}
  .store-suggestions div {{ padding: 8px 12px; cursor: pointer; font-size: 13px; }}
  .store-suggestions div:hover {{ background: #f0f4ff; }}
  .search-input-wrap {{ position: relative; }}
  .mode-tabs {{ display: flex; gap: 8px; margin-bottom: 20px; }}
  .mode-tab {{ padding: 7px 20px; border-radius: 20px; border: 2px solid #8E44AD; background: white; color: #8E44AD; font-size: 13px; cursor: pointer; font-family: inherit; }}
  .mode-tab.active {{ background: #8E44AD; color: white; }}
  .store-checkboxes {{ display: flex; flex-wrap: wrap; gap: 8px; margin-top: 8px; max-height: 160px; overflow-y: auto; padding: 8px; border: 1px solid #eee; border-radius: 8px; background: #fafafa; }}
  .store-checkbox-item {{ display: flex; align-items: center; gap: 6px; font-size: 13px; cursor: pointer; padding: 4px 8px; border-radius: 6px; }}
  .store-checkbox-item:hover {{ background: #f0f4ff; }}
  .check-all-btn {{ padding: 5px 14px; border: 1px solid #8E44AD; border-radius: 16px; background: white; color: #8E44AD; font-size: 12px; cursor: pointer; font-family: inherit; }}
  .check-all-btn:hover {{ background: #8E44AD; color: white; }}
  .store-checkbox-item {{ display: flex; align-items: center; gap: 8px; padding: 4px 0; font-size: 13px; cursor: pointer; }}
  .store-checkbox-item input {{ cursor: pointer; }}
  .company-result-title {{ font-size: 13px; color: #888; margin-bottom:4px; }}
  .view-tabs {{ display: flex; gap: 8px; margin-bottom: 16px; }}
  .view-tab {{ padding: 7px 18px; border-radius: 20px; border: 2px solid #2C3E50; background: white; color: #2C3E50; font-size: 13px; cursor: pointer; font-family: inherit; }}
  .view-tab.active {{ background: #2C3E50; color: white; }}
  .period-btn.month-btn {{ border-color: #E67E22; color: #E67E22; }}
  .period-btn.month-btn:hover, .period-btn.month-btn.active {{ background: #E67E22; color: white; }}
</style>
</head>
<body>
<header>
  <h1>旬彩坊 週次売上レポート</h1>
  <p>生成日時: {generated_at}</p>
</header>
<div class="container">

  <div class="summary-grid">
    <div class="summary-card">
      <div class="label">最新週</div>
      <div class="value">{weeks[-1] if weeks else '-'}</div>
    </div>
    <div class="summary-card">
      <div class="label">最新週 売上合計</div>
      <div class="value">{fmt_yen(latest_total)}</div>
    </div>
    <div class="summary-card">
      <div class="label">集計週数</div>
      <div class="value">{total_weeks}週</div>
    </div>
    <div class="summary-card">
      <div class="label">累計売上</div>
      <div class="value">{fmt_yen(cumulative)}</div>
    </div>
  </div>

  <div class="period-selector">
    <button class="period-btn active" id="weekly-btn" onclick="showWeeklyViewAll()">週次取引</button>
    <div class="divider"></div>
    <span>月別：</span>
    <select id="month-select" class="period-btn" style="padding:8px 16px;border-radius:20px;border:2px solid #E67E22;color:#E67E22;background:white;font-size:13px;cursor:pointer;font-family:inherit;" onchange="onMonthSelect(this)">
    </select>
    <div class="divider"></div>
    <button class="period-btn search-btn" id="search-toggle" onclick="showSearchView()">🔍 明細検索</button>
  </div>

  <!-- 週次ビュー -->
  <div id="weekly-view">
    <div class="card">
      <h2>① 全体の週次売上推移</h2>
      <div class="chart-wrap"><canvas id="chart1"></canvas></div>
    </div>
    <div class="card">
      <h2>② ソース別の週次売上推移</h2>
      <div class="chart-wrap"><canvas id="chart2"></canvas></div>
    </div>
    <div class="card">
      <h2>③ 中規模以上企業の週次売上推移</h2>
      <div class="chart-wrap"><canvas id="chart3"></canvas></div>
    </div>
    <div class="card">
      <h2>④ 個人店舗の週次売上推移</h2>
      <div class="chart-wrap"><canvas id="chart4"></canvas></div>
    </div>
  </div>

  <!-- 月次ビュー -->
  <div id="monthly-view" style="display:none;">
    <div class="card">
      <h2 id="monthly-title1">① 日次売上</h2>
      <div class="chart-wrap-tall"><canvas id="chart_m1"></canvas></div>
    </div>
    <div class="card">
      <h2 id="monthly-title2">② 中規模以上企業 累計売上</h2>
      <div class="chart-wrap"><canvas id="chart_m2"></canvas></div>
    </div>
    <div class="card">
      <h2 id="monthly-title3">③ 個人店舗 累計売上</h2>
      <div class="chart-wrap"><canvas id="chart_m3"></canvas></div>
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
      <div style="display:flex;gap:32px;flex-wrap:wrap;">
        <div style="flex:1;min-width:300px;">
          <div style="font-size:12px;color:#888;font-weight:bold;margin-bottom:8px;">● 店舗で検索</div>
          <div class="search-row" style="margin-bottom:8px;">
            <div class="search-input-wrap">
              <input type="text" id="store-input" class="search-input" placeholder="店舗名を入力（空欄で全店舗）" autocomplete="off" oninput="filterSuggestions()" onblur="hideSuggestions()" style="width:100%;" />
              <div class="store-suggestions" id="suggestions" style="display:none;"></div>
            </div>
          </div>
          <div class="search-row">
            <select id="store-select" style="padding:8px 12px;border:1px solid #ddd;border-radius:8px;font-size:13px;font-family:inherit;min-width:200px;">
              <option value="">（プルダウンで選択）</option>
            </select>
          </div>
          <div class="search-row" style="margin-top:12px;">
            <button class="search-btn-run" onclick="runStoreSearch()">店舗で検索</button>
            <button class="search-btn-run" id="csv-btn" onclick="downloadCSV()" style="background:#27AE60;display:none;">CSVダウンロード</button>
          </div>
        </div>
        <div style="flex:1;min-width:300px;">
          <div style="font-size:12px;color:#888;font-weight:bold;margin-bottom:8px;">● 企業で検索</div>
          <div class="search-row" style="margin-bottom:8px;">
            <select id="company-select" style="padding:8px 12px;border:1px solid #ddd;border-radius:8px;font-size:13px;font-family:inherit;min-width:200px;" onchange="updateStoreCheckboxes()">
              <option value="">企業を選択</option>
            </select>
          </div>
          <div id="store-checkboxes" style="max-height:160px;overflow-y:auto;border:1px solid #ddd;border-radius:8px;padding:8px 12px;display:none;"></div>
          <div class="search-row" style="margin-top:12px;">
            <button class="search-btn-run" onclick="runCompanySearch()">企業で検索</button>
            <button class="search-btn-run" id="csv-btn2" onclick="downloadCSV2()" style="background:#27AE60;display:none;">CSVダウンロード</button>
          </div>
        </div>
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
              <th class="num">購入金額</th>
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
const SEARCH_DATA = {search_json};


// 初期化
const storeSelect = document.getElementById('store-select');
SEARCH_DATA.stores.forEach(s => {{
  const opt = document.createElement('option');
  opt.value = s; opt.textContent = s;
  storeSelect.appendChild(opt);
}});
storeSelect.addEventListener('change', () => {{
  document.getElementById('store-input').value = storeSelect.value;
}});
// 企業プルダウン初期化
const companySelect = document.getElementById('company-select');
SEARCH_DATA.known_companies.forEach(c => {{
  if (SEARCH_DATA.company_stores[c]) {{
    const opt = document.createElement('option');
    opt.value = c; opt.textContent = c;
    companySelect.appendChild(opt);
  }}
}});

function updateStoreCheckboxes() {{
  const company = companySelect.value;
  const box = document.getElementById('store-checkboxes');
  box.innerHTML = '';
  if (!company) {{ box.style.display = 'none'; return; }}
  const stores = SEARCH_DATA.company_stores[company] || [];
  stores.forEach(s => {{
    const label = document.createElement('label');
    label.className = 'store-checkbox-item';
    const cb = document.createElement('input');
    cb.type = 'checkbox'; cb.value = s; cb.checked = true;
    label.appendChild(cb);
    label.appendChild(document.createTextNode(s));
    box.appendChild(label);
  }});
  box.style.display = 'block';
}}

let _companyRows = [];
function runCompanySearch() {{
  const company = companySelect.value;
  if (!company) {{ alert('企業を選択してください'); return; }}
  const checkboxes = document.querySelectorAll('#store-checkboxes input[type=checkbox]:checked');
  const selectedStores = Array.from(checkboxes).map(cb => cb.value);
  if (selectedStores.length === 0) {{ alert('店舗を1つ以上選択してください'); return; }}
  const from = document.getElementById('date-from').value;
  const to   = document.getElementById('date-to').value;
  const filtered = SEARCH_DATA.records.filter(r => {{
    if (from && r.date < from) return false;
    if (to   && r.date > to)   return false;
    return selectedStores.includes(r.store);
  }});
  const agg = {{}};
  filtered.forEach(r => {{
    if (!agg[r.product]) agg[r.product] = {{price: r.price, qty: 0, amount: 0}};
    agg[r.product].qty    += r.qty;
    agg[r.product].amount += r.amount;
  }});
  _companyRows = Object.entries(agg)
    .map(([name, v]) => ({{name, price: v.price, qty: v.qty, amount: v.amount}}))
    .sort((a, b) => b.amount - a.amount);
  const total = _companyRows.reduce((s, r) => s + r.amount, 0);
  document.getElementById('search-result-title').textContent =
    company + '　' + from + ' 〜 ' + to + '　商品別明細';
  document.getElementById('result-summary').textContent =
    '対象店舗: ' + selectedStores.length + '店　商品数: ' + _companyRows.length + '種　合計: ¥' + total.toLocaleString();
  const tbody = document.getElementById('result-body');
  tbody.innerHTML = '';
  _companyRows.forEach(r => {{
    const tr = document.createElement('tr');
    const qty_str = r.qty % 1 === 0 ? r.qty.toLocaleString() : r.qty.toFixed(2);
    const price_str = r.price > 0 ? '¥' + r.price.toLocaleString() : '-';
    tr.innerHTML = '<td>' + r.name + '</td>' +
      '<td class="num">' + price_str + '</td>' +
      '<td class="num">' + qty_str + '</td>' +
      '<td class="num">¥' + r.amount.toLocaleString() + '</td>';
    tbody.appendChild(tr);
  }});
  document.getElementById('csv-btn2').style.display = _companyRows.length > 0 ? 'inline-block' : 'none';
}}

function downloadCSV2() {{
  const rows = _companyRows || [];
  const company = companySelect.value;
  const from = document.getElementById('date-from').value;
  const to   = document.getElementById('date-to').value;
  const sep = String.fromCharCode(13, 10);
  const lines = ['\uFEFF商品名,単価,数量,購入金額'];
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
  a.download = company + '_' + from + '_' + to + '.csv';
  a.click();
  URL.revokeObjectURL(url);
}}

document.getElementById('date-from').value = SEARCH_DATA.min_date;
document.getElementById('date-to').value = SEARCH_DATA.max_date;

// 検索窓サジェスト
function filterSuggestions() {{
  const val = document.getElementById('store-input').value.toLowerCase();
  const box = document.getElementById('suggestions');
  if (!val) {{ box.style.display = 'none'; return; }}
  const matches = SEARCH_DATA.stores.filter(s => s.toLowerCase().includes(val)).slice(0, 15);
  if (matches.length === 0) {{ box.style.display = 'none'; return; }}
  box.innerHTML = '';
  matches.forEach(s => {{
    const d = document.createElement('div');
    d.textContent = s;
    d.onmousedown = () => {{
      document.getElementById('store-input').value = s;
      box.style.display = 'none';
    }};
    box.appendChild(d);
  }});
  box.style.display = 'block';
}}
function hideSuggestions() {{
  setTimeout(() => {{ document.getElementById('suggestions').style.display = 'none'; }}, 150);
}}

function showSearchView() {{
  document.getElementById('weekly-view').style.display = 'none';
  document.getElementById('monthly-view').style.display = 'none';
  document.getElementById('search-view').style.display = 'block';
  document.querySelectorAll('.period-btn').forEach(b => b.classList.remove('active'));
  document.getElementById('search-toggle').classList.add('active');
  runStoreSearch();
}}

function runStoreSearch() {{
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
  const lines = ['\uFEFF商品名,単価,数量,購入金額'];
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


const ALL_DATA = {data_json};
const weeks = ALL_DATA.weeks;
const weekly = ALL_DATA.weekly;
const monthly = ALL_DATA.monthly;

function setActiveBtn(activeBtn) {{
  document.querySelectorAll('.period-btn').forEach(b => b.classList.remove('active'));
  activeBtn.classList.add('active');
}}

function showWeeklyViewAll() {{
  setActiveBtn(document.getElementById('weekly-btn'));
  showWeeklyView(weeks);
}}

// 月別プルダウン
const monthSelect = document.getElementById('month-select');
const months = Object.keys(monthly);
months.forEach(month => {{
  const opt = document.createElement('option');
  opt.value = month;
  opt.textContent = month;
  monthSelect.appendChild(opt);
}});
if (months.length > 0) {{
  monthSelect.value = months[months.length - 1];
}}

function onMonthSelect(sel) {{
  const month = sel.value;
  if (!month) return;
  document.querySelectorAll('.period-btn').forEach(b => b.classList.remove('active'));
  showMonthlyView(month);
}}

// 週次チャート
const datalabels = {{
  display: true, align: 'top', anchor: 'end',
  font: {{ size: 10 }},
  formatter: v => v > 0 ? '¥' + Math.round(v/10000) + '万' : '',
  color: '#555',
}};
const lineOpts = (legend) => ({{
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
    datalabels: {{
      display: true, anchor: 'end', align: 'top',
      font: {{ size: 10 }},
      formatter: v => v > 0 ? '¥' + Math.round(v/10000) + '万' : '',
      color: '#555',
    }},
  }},
  scales: {{ y: {{ ticks: {{ callback: v => '¥' + v.toLocaleString() }} }} }}
}});

const chart1 = new Chart(document.getElementById('chart1'), {{
  type: 'line',
  data: {{ labels: [], datasets: [{{ label: '週次売上合計', borderColor: '#4e9af1', backgroundColor: '#4e9af133', tension: 0.3, fill: true, pointRadius: 6, data: [] }}] }},
  options: lineOpts(false),
}});
const srcKeys = Object.keys(ALL_DATA.source_colors);
const chart2 = new Chart(document.getElementById('chart2'), {{
  type: 'line',
  data: {{ labels: [], datasets: srcKeys.map(src => ({{
    label: src, borderColor: ALL_DATA.source_colors[src],
    backgroundColor: ALL_DATA.source_colors[src] + '33',
    tension: 0.3, fill: false, pointRadius: 6, data: []
  }})) }},
  options: lineOpts(true),
}});
const chart3 = new Chart(document.getElementById('chart3'), {{
  type: 'line',
  data: {{ labels: [], datasets: ALL_DATA.large_companies.map((c, i) => ({{
    label: c, borderColor: ALL_DATA.palette_large[i],
    backgroundColor: ALL_DATA.palette_large[i] + '33',
    tension: 0.3, fill: false, pointRadius: 6, data: []
  }})) }},
  options: lineOpts(true),
}});
const chart4 = new Chart(document.getElementById('chart4'), {{
  type: 'line',
  data: {{ labels: [], datasets: ALL_DATA.small_companies.map((c, i) => ({{
    label: c, borderColor: ALL_DATA.palette_small[i],
    backgroundColor: ALL_DATA.palette_small[i] + '33',
    tension: 0.3, fill: false, pointRadius: 6, data: []
  }})) }},
  options: lineOpts(true),
}});

// 月次チャート
const chart_m1 = new Chart(document.getElementById('chart_m1'), {{
  type: 'bar',
  data: {{ labels: [], datasets: [{{ label: '日次売上', backgroundColor: '#4e9af1aa', data: [] }}] }},
  options: barOpts(false),
}});
const chart_m2 = new Chart(document.getElementById('chart_m2'), {{
  type: 'bar',
  data: {{ labels: ALL_DATA.large_companies, datasets: [{{ label: '累計売上', backgroundColor: ALL_DATA.palette_large, data: [] }}] }},
  options: barOpts(false),
}});
const chart_m3 = new Chart(document.getElementById('chart_m3'), {{
  type: 'bar',
  data: {{ labels: ALL_DATA.small_companies, datasets: [{{ label: '累計売上', backgroundColor: ALL_DATA.palette_small, data: [] }}] }},
  options: barOpts(false),
}});

function showWeeklyView(filtered) {{
  document.getElementById('weekly-view').style.display = 'block';
  document.getElementById('monthly-view').style.display = 'none';
  chart1.data.labels = filtered;
  chart1.data.datasets[0].data = filtered.map(w => weekly[w].total);
  chart1.update();
  chart2.data.labels = filtered;
  srcKeys.forEach((src, i) => {{ chart2.data.datasets[i].data = filtered.map(w => weekly[w].by_source[src] || 0); }});
  chart2.update();
  chart3.data.labels = filtered;
  ALL_DATA.large_companies.forEach((c, i) => {{ chart3.data.datasets[i].data = filtered.map(w => weekly[w].by_company[c] || 0); }});
  chart3.update();
  chart4.data.labels = filtered;
  ALL_DATA.small_companies.forEach((c, i) => {{ chart4.data.datasets[i].data = filtered.map(w => weekly[w].by_company[c] || 0); }});
  chart4.update();
}}

function showMonthlyView(month) {{
  document.getElementById('weekly-view').style.display = 'none';
  document.getElementById('monthly-view').style.display = 'block';
  const md = monthly[month];
  document.getElementById('monthly-title1').textContent = '① ' + month + ' 日次売上';
  document.getElementById('monthly-title2').textContent = '② ' + month + ' 中規模以上企業 累計売上';
  document.getElementById('monthly-title3').textContent = '③ ' + month + ' 個人店舗 累計売上';
  chart_m1.data.labels = md.dates;
  chart_m1.data.datasets[0].data = md.daily_totals;
  chart_m1.update();
  chart_m2.data.datasets[0].data = ALL_DATA.large_companies.map(c => md.large_totals[c] || 0);
  chart_m2.update();
  chart_m3.data.datasets[0].data = ALL_DATA.small_companies.map(c => md.small_totals[c] || 0);
  chart_m3.update();
}}

showWeeklyView(weeks);
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
    print(f"  店舗マスタ: {len(store_map)}件")
    print("■ 集計中...")
    weeks, weekly, month_to_weeks = build_weekly_data(df_summary, store_map)
    monthly = build_monthly_data(df_detail, store_map)
    search_data = build_search_data(df_detail, store_map)
    all_data = {
        "weeks": weeks,
        "weekly": weekly,
        "month_to_weeks": month_to_weeks,
        "monthly": monthly,
        "large_companies": LARGE_COMPANIES,
        "small_companies": SMALL_COMPANIES,
        "source_colors": SOURCE_COLORS,
        "palette_large": PALETTE_LARGE,
        "palette_small": PALETTE_SMALL,
    }
    print("■ HTML生成中...")
    generated_at = datetime.now().strftime("%Y-%m-%d %H:%M")
    html = build_html(all_data, search_data, generated_at)
    OUTPUT_DIR.mkdir(exist_ok=True)
    OUTPUT_PATH.write_text(html, encoding="utf-8")
    print(f"✅ 出力完了: {OUTPUT_PATH}")
    print(f"   ブラウザで開いてください: open '{OUTPUT_PATH}'")

if __name__ == "__main__":
    main()
