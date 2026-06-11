#!/bin/bash
# 旬彩坊 週次集計 自動実行スクリプト
# 毎週水曜 10:00 / 23:00 に launchd から起動される

set -euo pipefail

PROJECT_DIR="$HOME/Documents/旬彩坊_週次集計"
LOG_FILE="$PROJECT_DIR/run_weekly.log"

log() {
  echo "[$(date '+%Y-%m-%d %H:%M:%S')] $*" | tee -a "$LOG_FILE"
}

# ── 対象週を計算（月曜始まりのISO週番号）──────────────────────────
WEEK=$(python3 -c "
import datetime
today = datetime.date.today()
iso = today.isocalendar()
print(f'{iso[0]}-W{iso[1]:02d}')
")

INPUT_DIR="$PROJECT_DIR/input/$WEEK"
DONE_FLAG="$PROJECT_DIR/input/$WEEK/.done"

log "========================================"
log "対象週: $WEEK"

# ── 既に集計済みなら終了 ─────────────────────────────────────────
if [ -f "$DONE_FLAG" ]; then
  log "集計済み（.done フラグあり）。スキップします。"
  exit 0
fi

# ── inputフォルダがなければ終了 ──────────────────────────────────
if [ ! -d "$INPUT_DIR" ] || [ -z "$(ls -A "$INPUT_DIR" 2>/dev/null)" ]; then
  log "input/$WEEK が存在しないか空のため、スキップします。"
  exit 0
fi

log "input/$WEEK を確認。集計を開始します。"

cd "$PROJECT_DIR"
source .venv/bin/activate

# ── XLS → XLSX 変換（xlrd 経由で pandas が読めるため不要だが念のため確認）──
log "販売大臣ファイルを確認中..."
XLS_FILE=$(find "$INPUT_DIR" -iname "*.xls" ! -iname "*.xlsx" | head -1)
if [ -n "$XLS_FILE" ]; then
  XLSX_FILE="${XLS_FILE%.*}.xlsx"
  if [ ! -f "$XLSX_FILE" ]; then
    log "  XLS検出: $XLS_FILE → xlrd で直接読み込みます（変換不要）"
  fi
fi

# ── 集計実行 ─────────────────────────────────────────────────────
log "aggregate.py 実行中..."
python3 aggregate.py "$WEEK" 2>&1 | tee -a "$LOG_FILE"

# ── レポート生成 ──────────────────────────────────────────────────
log "report.py 実行中..."
python3 report.py 2>&1 | tee -a "$LOG_FILE"

log "company_report.py 実行中..."
python3 company_report.py 2>&1 | tee -a "$LOG_FILE"

# ── GitHub Pages へ push ─────────────────────────────────────────
log "git push 中..."
git add output/
git commit -m "$WEEK 集計追加（自動）"
git push 2>&1 | tee -a "$LOG_FILE"

# ── 完了フラグを立てる ────────────────────────────────────────────
touch "$DONE_FLAG"
log "完了。レポート URL: https://t-fujiya-sd.github.io/shunsaibou-report/output/report.html"
log "========================================"
