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
  └── outlook-tui  （Textual TUIアプリ）
        └── 同じ OutlookBase 経由
```

### フォルダ構成

```
outlook-cli/
├── CLAUDE.md
├── pyproject.toml
├── mock_data.json         ← モックデータ（自宅開発用）
└── outlook_cli/
    ├── base.py            ← OutlookBase 抽象クラス
    ├── mock.py            ← OutlookMock
    ├── com.py             ← OutlookCOM（win32com）
    ├── cli.py             ← Click エントリポイント
    └── tui.py             ← Textual TUI
```

新しいメソッドを追加するときは **base.py・mock.py・com.py の3箇所**に実装する。

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

---

## ブリッジ構成（将来タスク：WSL2 / LoChaBot 連携）

```
WSL2 (LoChaBot / Claude Code)
    │  HTTP (localhost:5050)
    ▼
Windows Python サーバー (outlook_bridge.py)
    │  win32com
    ▼
Outlook COM
```

エンドポイント想定: `GET /unread/count` `/unread/summary` `/sent/today`

---

## 振り返りワークフロー（目標形態）

```
朝イチ  → unread-count     # 今日の負荷を把握
随時    → unread-summary   # Outlookを開かずに内容確認
夕方    → sent-today | LLM # 日次振り返りサマリー生成
```

将来的には LoChaBot の APScheduler Cog として自動実行し Discord に投げる。

---

## 次のアクション

1. `outlook_bridge.py` の実装 → HTTP化して LoChaBot から呼べるようにする
2. 会社PCで `OutlookCOM` の動作確認（`unread-summary` / `sent-today`）

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
