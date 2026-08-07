# Outlook CLI — CLAUDE.md

## プロジェクト概要

会社PC（Panasonic Let's Note CF-SV1 / Windows 11）でOutlookのGUIを開かずに
メール操作・未読確認・日次振り返りをCLI/Skill経由で行うツール群。

---

## 環境

| 項目 | 内容 |
|------|------|
| 会社PC | Panasonic Let's Note CF-SV1 |
| OS | Windows 11 |
| Outlook | M365 Apps for Enterprise（クイック実行）|
| Python管理 | uv（winget経由） |
| ターミナル | PowerShell / Windows Terminal |
| エディタ | Neovim / Claude Code |

---

## アーキテクチャ

```
Neovim / Claude Code
  │
  ├── outlook-cli  （Click CLI）
  │     └── OutlookBase（抽象クラス）
  │           ├── OutlookMock  ← 自宅開発用（JSONファイル）
  │           └── OutlookCOM   ← 会社PC用（win32com）
  │
  ├── outlook-tui  （Textual TUIアプリ）
  │     └── 同じ OutlookBase 経由
  │
  ├── outlook-mcp  （MCPサーバー、Streamable HTTP）
  │     └── 同じ OutlookBase 経由。CLIコマンド相当の12ツールを公開
  │
  └── daily_report  （スタンドアロン日報生成）
        ├── OutlookBase をライブラリとしてインポート
        ├── ActivityWatch REST API / SQLite
        └── ローカル LLM（OpenAI 互換 API）→ Obsidian に Markdown 保存
```

### フォルダ構成

```
outlook-cli/
├── CLAUDE.md
├── pyproject.toml
├── .env                   ← LLM設定・出力先（.gitignore済み）
├── mock_data.json         ← モックデータ（自宅開発用）
├── outlook_cli/
│   ├── base.py            ← OutlookBase 抽象クラス
│   ├── mock.py            ← OutlookMock
│   ├── com.py             ← OutlookCOM（win32com）
│   ├── cli.py             ← Click エントリポイント
│   ├── tui.py             ← Textual TUI
│   └── mcp_server.py      ← MCPサーバー（Streamable HTTP, port 8764）
├── daily_report/
│   ├── __main__.py        ← python -m daily_report エントリポイント
│   ├── collect.py         ← カレンダー・送信メール・AW収集
│   ├── llm.py             ← LLM API呼び出し・サーバ起動確認
│   └── report.py          ← Markdown保存・win10toast通知
├── scripts/
│   └── daily_report.ps1   ← タスクスケジューラ用ラッパー
└── tests/
    └── test_daily_report.py
```

outlook_cli に新しいメソッドを追加するときは **base.py・mock.py・com.py の3箇所**に実装する。

### 環境切り替え

```powershell
# 自宅（モック）
$env:OUTLOOK_MOCK=1; .venv\Scripts\python.exe -m outlook_cli.tui

# 会社PC（COM）
.venv\Scripts\python.exe -m outlook_cli.tui
```

---

## 実装済み CLI コマンド

| コマンド | 説明 | 主なオプション |
|---|---|---|
| `list` | メール一覧 | `--folder` `--limit` `--json-output` |
| `search <keyword>` | メール検索 | `--days` `--from` `--json-output` |
| `read <mail_id>` | 本文表示 | `--json-output` |
| `send` | メール送信 | `--to` `--subject` `--body` |
| `reply <mail_id>` | 返信 | `--body` |
| `unread-count` | フォルダ別未読件数 | `--folder` `--json-output` |
| `unread-summary` | 未読メール本文冒頭サマリー | `--folder` `--limit` `--json-output` |
| `sent-today` | 当日の送信メール一覧 | `--date` `--json-output` |
| `cal today` | 今日の予定一覧 | `--date` `--json-output` |
| `cal week` | 今週（月〜金）の予定一覧 | `--date` `--json-output` |
| `find-slot` | 複数人の空き時間候補 | `--attendees` `--duration` `--days` `--work-start` `--work-end` `--json-output` |

上記コマンドはすべてMCPツールとしても公開されている（[MCPサーバー](#mcpサーバーoutlook-mcp)参照）。

---

## MCPサーバー（outlook-mcp）

```
Diffcoder（MCPクライアント機能を実装予定）
    │  Streamable HTTP (0.0.0.0:8764/mcp)
    ▼
outlook_cli/mcp_server.py
    │  OutlookBase
    ▼
OutlookMock / OutlookCOM
```

CLIコマンド相当の12ツール（list/search/read/send/reply/unread_count/unread_summary/
sent_today/flagged/cal_today/cal_week/find_slot）を公開。認証なし（社内LAN前提）。
起動: `outlook-mcp`（環境変数 `OUTLOOK_MCP_HOST` / `OUTLOOK_MCP_PORT` で上書き可）。

旧来検討していたLoChaBot向けRESTブリッジ（`outlook_bridge.py`, port 5050）構想は、
LoChaBotの廃止（Discord bot移行）に伴い不要と判断し、MCPサーバーに統合済み。

---

## 振り返りワークフロー

```
朝イチ  → outlook unread-count      # 今日の負荷を把握
随時    → outlook unread-summary    # Outlookを開かずに内容確認
夕方    → python -m daily_report    # 日報自動生成（タスクスケジューラで自動実行）
```

日報は `V:\obsidian\YYYY-MM-DD.md` に保存される。win10toast で完了通知あり。

---

## daily_report の設定（.env）

```
DAILY_REPORT_API_ENDPOINT=http://192.168.1.76:8089/v1/chat/completions
DAILY_REPORT_API_HEALTH_URL=http://192.168.1.76:8089/v1/models
DAILY_REPORT_API_START_CMD=   # APIサーバの起動コマンド（未設定時は自動起動しない）
DAILY_REPORT_OUTPUT_DIR=V:\obsidian
ACTIVITYWATCH_DB=             # 空の場合はREST API優先（localhost:5600）
```

---

## 次のアクション

1. 会社PCで `OutlookCOM` の動作確認（`unread-summary` / `sent-today` / `daily-report`）
2. `DAILY_REPORT_API_START_CMD` に LLM サーバの起動コマンドを設定
3. Mattermost トークン取得後に `daily_report/collect.py` の `collect_mattermost()` を実装
4. タスクスケジューラに `scripts\daily_report.ps1` を登録（毎日17:30）
5. 会社PCで `outlook-mcp`（`OutlookCOM`経由）の動作確認、Diffcoderからの接続確認

---

## Textual 実装の注意点

### `ListView.clear()` は必ず `await` する

```python
# NG: 古いアイテムがDOMに残り DuplicateIds エラー
mail_list.clear()

# OK
await mail_list.clear()
```

### ウィジェットIDの制約

英数字・アンダースコア・ハイフンのみ。日本語フォルダ名はインデックスで代替。
- フォルダ: `folder-sub-0`, `folder-sub-1`
- メール: `mail-0`, `mail-1`

インデックスと実データのマッピングは `self._subfolders` / `self.current_mails` で保持。

### App レベルのキーバインド制限

App の BINDINGS はフォーカスウィジェットがキーを消費しない場合のみ発火。
Vimキー（j/k/h/l）は `_main_list_focused()` で folder-list / mail-list にフォーカスがある場合のみ動作。

---

## 開発メモ

- `win32com` は管理者権限不要・IT制限に引っかからない
- `Restrict()` でサーバー側フィルタリングするので大量メールでも高速
- `mail.EntryID` をIDとして使い `GetItemFromID()` で直接取得可能
- 自宅開発時は `OUTLOOK_MOCK=1` でモックに切り替え
