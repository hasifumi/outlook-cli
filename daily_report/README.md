# daily_report

Outlook・ActivityWatch・Mattermost のデータをローカル LLM に渡し、
日報 Markdown を自動生成して Obsidian vault に保存するスタンドアロンツール。

## 動作の流れ

```
python -m daily_report
  │
  ├─ LLM API サーバの死活確認（未起動なら自動起動）
  ├─ Outlook カレンダー収集（今日 + 明日）
  ├─ 送信メール収集（当日分）
  ├─ ActivityWatch 作業ログ収集（REST API → SQLite フォールバック）
  ├─ LLM に 3 テキストブロックを渡して日報生成
  └─ YYYY-MM-DD.md として保存 → Windows 通知
```

## セットアップ

### 1. `.env` を設定する

プロジェクトルートの `.env` を編集する（`.gitignore` 済みなので直接書いてOK）：

```
DAILY_REPORT_API_ENDPOINT=http://192.168.1.76:8089/v1/chat/completions
DAILY_REPORT_API_HEALTH_URL=http://192.168.1.76:8089/v1/models
DAILY_REPORT_API_START_CMD=Start-Process "path\to\server.exe"
DAILY_REPORT_OUTPUT_DIR=V:\obsidian
ACTIVITYWATCH_DB=
```

| 変数 | 説明 | デフォルト |
|------|------|-----------|
| `DAILY_REPORT_API_ENDPOINT` | LLM API の `/v1/chat/completions` URL | `http://localhost:8089/v1/chat/completions` |
| `DAILY_REPORT_API_HEALTH_URL` | 死活確認 URL（`/v1/models` 等） | `http://localhost:8089/v1/models` |
| `DAILY_REPORT_API_START_CMD` | 未起動時に実行する PowerShell コマンド | （空の場合は自動起動しない） |
| `DAILY_REPORT_OUTPUT_DIR` | 出力先ディレクトリ | `.`（カレントディレクトリ） |
| `ACTIVITYWATCH_DB` | ActivityWatch SQLite DB パス | 空の場合は REST API 優先 |

### 2. 動作確認（モック環境）

```powershell
$env:OUTLOOK_MOCK = 1
.venv\Scripts\python.exe -m daily_report --output-dir C:\tmp --no-notify
# C:\tmp\2026-06-21.md が生成されることを確認
```

### 3. タスクスケジューラへの登録

`scripts\daily_report.ps1` をタスクスケジューラに登録する：

1. タスクスケジューラを開く（`taskschd.msc`）
2. 「基本タスクの作成」
3. トリガー: 毎日 17:30
4. 操作: プログラムの開始
   - プログラム: `powershell.exe`
   - 引数: `-NonInteractive -File "C:\Users\hassy\project\outlook-cli\scripts\daily_report.ps1"`
   - 開始: `C:\Users\hassy\project\outlook-cli`

## 使い方

```powershell
# 今日の日報を生成
python -m daily_report

# 出力先を上書き
python -m daily_report --output-dir C:\tmp

# 特定の日付で生成
python -m daily_report --date 2026-06-20

# Windows 通知なしで実行
python -m daily_report --no-notify
```

## 生成される日報のフォーマット

```markdown
## やったこと
- 〇〇 MTG に参加（10:00-11:00）
- 田中さんへプロジェクト進捗を送付

## 明日やること
- 週次レビュー（10:00-11:00）
- 資料作成の続き

## 所感
全体的にスムーズに進んだ一日でした。...
```

## アーキテクチャ

```
daily_report/
├── __main__.py   Click CLI エントリポイント・全工程のオーケストレーション
├── collect.py    データ収集（Outlook / ActivityWatch / Mattermost stub）
├── llm.py        LLM API 呼び出し・APIサーバ起動確認
└── report.py     Markdown 保存・win10toast 通知
```

`collect.py` は `outlook_cli.base.OutlookBase` の実装（`OutlookCOM` または `OutlookMock`）を
ライブラリとして使う。`outlook-cli` を事前起動する必要はない（`win32com` が COM サーバを自動起動する）。

## データソースの詳細

### カレンダー（`collect_calendar`）
`OutlookBase.get_calendar()` で今日・明日の予定を取得してテキスト化する。

### 送信メール + Mattermost（`collect_sent_mail` + `collect_mattermost`）
`OutlookBase.sent_today()` で当日の送信メールを取得する。
Mattermost は現在 stub（空文字を返す）。トークン取得後に実装予定。

### ActivityWatch（`collect_activitywatch`）
1. REST API（`http://localhost:5600/api/0/buckets`）を優先して試みる
2. 失敗した場合は SQLite DB を直接読む
3. アプリ別の使用時間を集計して上位10件を出力する
4. AW が完全に使えない場合は空文字を返し、日報生成を続行する

## テスト

```powershell
uv run pytest tests/test_daily_report.py -v
```

13 テストケースがある（collect・save・llm の純粋関数を unittest.mock でテスト）。

## 未実装・今後の予定

- Mattermost API 統合（`collect_mattermost()` を実装）
- `DAILY_REPORT_API_START_CMD` の具体的なコマンド設定
