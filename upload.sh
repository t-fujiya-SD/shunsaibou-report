#!/bin/bash
cd ~/Documents/旬彩坊_週次集計
git add output/report.html output/report_*.html
git commit -m "Weekly update $(date '+%Y-%m-%d')"
git push origin main
echo "✅ GitHubにアップロード完了"
