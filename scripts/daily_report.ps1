# daily_report.ps1 — タスクスケジューラから呼び出すラッパー
#
# タスクスケジューラ登録手順:
#   タスクスケジューラ → 基本タスクの作成
#   トリガー: 毎日 17:30
#   操作: プログラムの開始
#     プログラム: powershell.exe
#     引数:  -NonInteractive -File "C:\Users\hassy\project\outlook-cli\scripts\daily_report.ps1"
#     開始: C:\Users\hassy\project\outlook-cli

$projectDir = "C:\Users\hassy\project\outlook-cli"
$python = "$projectDir\.venv\Scripts\python.exe"

& $python -m daily_report
