# outlook-cli

Windows 11 + Outlook M365 をターミナルから操作する CLI / TUI ツール。
Outlook の GUI を開かずにメール確認・送受信・カレンダー参照・空き時間検索ができる。

## 動作環境

| 項目 | 内容 |
|------|------|
| OS | Windows 11 |
| Outlook | M365 Apps for Enterprise（クイック実行） |
| Python | 3.11 以上 |
| パッケージ管理 | [uv](https://docs.astral.sh/uv/) |

## セットアップ

```powershell
# リポジトリのクローン
git clone https://github.com/hasifumi/outlook-cli.git
cd outlook-cli

# 依存パッケージのインストール
uv sync

# win32com（会社PC用）
uv add pywin32
```

### 環境切り替え

```powershell
# 自宅（JSON モック）
$env:OUTLOOK_MOCK = 1

# 会社PC（Outlook COM）— 環境変数を設定しなければ自動的に COM を使用
Remove-Item Env:OUTLOOK_MOCK
```

## CLI コマンド

```powershell
# 共通プレフィックス
.venv\Scripts\python.exe -m outlook_cli.cli <コマンド> [オプション]
# または uv sync 後に
outlook <コマンド> [オプション]
```

### メール操作

| コマンド | 説明 | 主なオプション |
|---|---|---|
| `list` | メール一覧 | `--folder inbox/sent/drafts` `--limit 20` `--json-output` |
| `search <keyword>` | メール検索 | `--days 7` `--from sender@example.com` `--json-output` |
| `read <mail_id>` | 本文表示 | `--json-output` |
| `send` | メール送信 | `--to` `--subject` `--body` |
| `reply <mail_id>` | 返信 | `--body` |
| `unread-count` | フォルダ別未読件数 | `--folder` `--json-output` |
| `unread-summary` | 未読メール本文冒頭サマリー | `--folder` `--limit` `--json-output` |
| `sent-today` | 当日の送信メール一覧 | `--date YYYY-MM-DD` `--json-output` |
| `flagged` | フラグ付き／期限設定メール | `--folder` `--days` `--json-output` |

### カレンダー操作

| コマンド | 説明 | 主なオプション |
|---|---|---|
| `cal today` | 今日の予定一覧 | `--date YYYY-MM-DD` `--json-output` |
| `cal week` | 今週（月〜金）の予定一覧 | `--date YYYY-MM-DD` `--json-output` |
| `find-slot` | 複数人の空き時間候補（上位5件） | `--attendees "a@co.jp;b@co.jp"` `--duration 60` `--days 5` `--work-start 9` `--work-end 18` `--json-output` |

### 使用例

```powershell
# 未読件数を確認
outlook unread-count

# 未読メールの内容をざっと把握
outlook unread-summary --limit 5

# 今日の予定を確認
outlook cal today

# 田中さんと佐藤さんの 60 分空き時間を今後5日で探す
outlook find-slot --attendees "tanaka@company.com;sato@company.com" --duration 60

# find-slot の出力例
# 空き時間候補（60分）
# ──────────────────────────────────────
# 1. ○  2026-05-18 (月) 09:00-10:00
# 2. ○  2026-05-19 (火) 14:00-15:00
# 3. △  2026-05-20 (水) 09:00-10:00  ※仮予定あり: tanaka@company.com
```

## MCP サーバー

CLIコマンドと同じ12個の操作（list/search/read/send/reply/unread_count/unread_summary/
sent_today/flagged/cal_today/cal_week/find_slot）をMCP（Model Context Protocol）ツールとして、
Streamable HTTPで公開する。外部ツール（MCPクライアント）から直接Outlookを操作できる。

```powershell
# 自宅（モック）で起動
$env:OUTLOOK_MOCK = 1
outlook-mcp
# または
.venv\Scripts\python.exe -m outlook_cli.mcp_server
```

デフォルトでは `http://0.0.0.0:8764/mcp` で待ち受ける。認証は付けていないため、
社内LANなど信頼できるネットワーク内での利用を前提とする。

| 環境変数 | 説明 | デフォルト |
|---|---|---|
| `OUTLOOK_MCP_HOST` | 待ち受けホスト | `0.0.0.0` |
| `OUTLOOK_MCP_PORT` | 待ち受けポート | `8764` |

## TUI

```powershell
# TUI 起動
outlook-tui
# または
.venv\Scripts\python.exe -m outlook_cli.tui
```

キーバインド: `j/k` でリスト移動、`h/l` でペイン切り替え、`q` で終了。

## 日報自動生成（daily_report）

Outlook・ActivityWatch・Mattermost のデータをローカル LLM に渡し、
日報 Markdown ファイルを自動生成して Obsidian vault に保存する。

```powershell
# 初回: .env を設定
copy .env.example .env   # または直接編集
# DAILY_REPORT_API_ENDPOINT=http://192.168.1.76:8089/v1/chat/completions
# DAILY_REPORT_OUTPUT_DIR=V:\obsidian

# 手動実行
.venv\Scripts\python.exe -m daily_report

# 出力先を一時的に変更
.venv\Scripts\python.exe -m daily_report --output-dir C:\tmp

# 過去日付で生成
.venv\Scripts\python.exe -m daily_report --date 2026-06-20
```

タスクスケジューラへの登録は `scripts\daily_report.ps1` のコメントを参照。
詳細は [`daily_report/README.md`](daily_report/README.md) を参照。

## アーキテクチャ

```
CLI / TUI / MCPサーバー（mcp_server.py）
  └── OutlookBase（抽象クラス）
        ├── OutlookMock  — JSON ファイルで動作（自宅開発用）
        └── OutlookCOM   — win32com 経由で Outlook に直接アクセス（会社PC用）

daily_report（スタンドアロン）
  ├── collect.py  — OutlookBase + ActivityWatch からデータ収集
  ├── llm.py      — ローカル LLM API 呼び出し（OpenAI 互換）
  └── report.py   — Markdown 保存・Windows 通知
```

新しいメソッドを追加するときは `base.py` `mock.py` `com.py` の3箇所に実装する。

## 開発

```powershell
# モック環境でテスト実行
$env:OUTLOOK_MOCK = 1
uv run pytest

# CLI を直接実行（インストール不要）
$env:OUTLOOK_MOCK = 1
.venv\Scripts\python.exe -m outlook_cli.cli cal today

# daily_report をモック環境で実行
$env:OUTLOOK_MOCK = 1
.venv\Scripts\python.exe -m daily_report --output-dir C:\tmp --no-notify
```

## ライセンス

MIT
