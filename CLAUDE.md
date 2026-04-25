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
| CPU | Intel Core i5-1145G7 |
| RAM | 8GB |
| Outlook | M365 Apps for Enterprise（クイック実行）|
| Python管理 | uv（winget経由） |
| ターミナル | PowerShell / Windows Terminal |
| エディタ | Neovim / Claude Code |

---

## 既存アーキテクチャ

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
├── CLAUDE.md              ← このファイル
├── DESIGN.md              ← 詳細設計仕様書
├── pyproject.toml
├── mock_data.json         ← モックデータ（自宅開発用）
└── outlook_cli/
    ├── __init__.py
    ├── base.py            ← OutlookBase 抽象クラス
    ├── mock.py            ← OutlookMock
    ├── com.py             ← OutlookCOM（win32com）
    ├── cli.py             ← Click エントリポイント
    └── tui.py             ← Textual TUI
```

### 環境切り替え

```powershell
# 自宅（モック）
$env:OUTLOOK_MOCK=1; .venv\Scripts\python.exe -m outlook_cli.tui

# 会社PC（COM）
.venv\Scripts\python.exe -m outlook_cli.tui
```

---

## 新規追加：Unread Skill 群

Outlookを開かずに未読状況を把握・活用するための軽量ツール。
Claude Code / OpenClaw の Skill として呼び出せる形を目指す。

### Skill 1: `outlook_unread_count`

**目的**: 未読数をフォルダ別に瞬時に把握する（超軽量）

```
入力: なし（オプションでフォルダ名指定）
出力: フォルダ別未読数サマリー（JSON or テキスト）

例:
  受信トレイ: 12件
  CCメール:    3件
  ---
  合計:       15件
```

**実装方針**:
- `OutlookCOM` の `GetDefaultFolder()` を使う
- `UnRead` プロパティでフィルタして件数カウント
- 既存の `OutlookBase` に `unread_count(folder)` メソッドを追加

---

### Skill 2: `outlook_unread_summary`

**目的**: 未読メールの内容をOutlookを開かずに把握・優先度判断

```
入力: limit（件数上限、デフォルト10）、folder（デフォルト inbox）
出力: 件名 / 送信者 / 受信日時 / 本文冒頭100文字 のリスト

例（JSON）:
[
  {
    "subject": "【重要】〇〇について",
    "from": "tanaka@example.com",
    "date": "2026-04-26 09:15",
    "preview": "お疲れ様です。先日の件について確認させてください..."
  },
  ...
]
```

**実装方針**:
- `OutlookCOM.list_mails()` の `unread=True` フィルタ版を追加
- 本文は `mail.Body[:100]` で冒頭抜粋
- `--output json` / `--output text` オプション対応

---

### Skill 3: `outlook_sent_summary`（既存機能の再整備）

**目的**: 当日の送信メールを抽出して「今日何をやったか」を記録

```
入力: date（デフォルト today）
出力: 送信件名・宛先・送信時刻のリスト

活用: 夕方に自動実行 → LLMに渡して日次サマリー生成
```

**実装方針**:
- 既存の送信メール抽出処理を `sent_today()` メソッドとして整理
- 日付フィルタを `[SentOn] >= 'MM/DD/YYYY'` で実装

---

## ブリッジ構成（WSL2環境での利用時）

会社PCがWindowsネイティブのため、WSL2 / LoChaBot から呼ぶ場合は：

```
WSL2 (LoChaBot / Claude Code)
    │  HTTP (localhost)
    ▼
Windows Python サーバー (outlook_bridge.py)
    │  win32com
    ▼
Outlook COM
```

### outlook_bridge.py（スケルトン）

```python
# Windows側で実行: python outlook_bridge.py
from flask import Flask, jsonify
from outlook_cli.com import OutlookCOM

app = Flask(__name__)
outlook = OutlookCOM()

@app.route("/unread/count")
def unread_count():
    return jsonify(outlook.unread_count())

@app.route("/unread/summary")
def unread_summary():
    return jsonify(outlook.unread_summary(limit=10))

@app.route("/sent/today")
def sent_today():
    return jsonify(outlook.sent_today())

if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5050)
```

---

## 振り返りワークフロー（目標形態）

```
朝イチ（自動 or 手動）
  → outlook_unread_count    # 今日の負荷を把握

作業中（随時）
  → outlook_unread_summary  # Outlookを開かずに内容確認

夕方（定時タスク or 手動）
  → outlook_sent_summary    # 今日やったことをログ
  → LLM（Ollama等）に渡す  # 日次振り返りサマリー生成

週次
  → 送信ログ + 未読推移 → 忙しさの原因分析
```

将来的には LoChaBot の APScheduler Cog として自動実行し、
Discord に投げる構成にする。

---

## 次のアクション

優先順で：

1. `OutlookBase` に以下のメソッドを追加定義
   - `unread_count(folder="inbox") -> dict`
   - `unread_summary(limit=10, folder="inbox") -> list`
   - `sent_today(date=None) -> list`

2. `OutlookMock` にモックデータを追加（自宅開発・テスト用）

3. `OutlookCOM` に実装（会社PCで動作確認）

4. CLIコマンドとして登録
   - `outlook unread count`
   - `outlook unread summary`
   - `outlook sent today`

5. （将来）`outlook_bridge.py` でHTTP化 → LoChaBot連携

---

## Textual 実装の注意点

### `ListView.clear()` は必ず `await` する

```python
# NG: 古いアイテムがDOMに残り DuplicateIds エラー
mail_list.clear()
mail_list.append(ListItem(..., id="mail-0"))

# OK
await mail_list.clear()
mail_list.append(ListItem(..., id="mail-0"))
```

### ウィジェットIDの制約

Textual のIDには英数字・アンダースコア・ハイフンのみ使用可能。
- 日本語フォルダ名 → インデックスで `folder-sub-0`, `folder-sub-1`
- メールID → `mail-0`, `mail-1`（連番）
- 連絡先候補 → `c-0`, `c-1`（連番）

インデックスと実データのマッピングは `self._subfolders` / `self.current_mails` / `self._candidates` で保持。

### App レベルのキーバインド制限

App の BINDINGS は、フォーカスしているウィジェットがキーを消費しない場合のみ発火する。
Vimキー（j/k/h/l/g/G）は `_main_list_focused()` で folder-list / mail-list-view にフォーカスがある場合のみ動作させる（ComposeScreen の Input に干渉しないよう）。

---

## 開発メモ

- `win32com` は管理者権限不要・既存Outlookを操作するだけなのでIT制限に引っかからない
- `Restrict()` でサーバー側フィルタリングするので大量メールでも高速
- `mail.EntryID` をIDとして使い `GetItemFromID()` で直接取得可能
- 自宅開発時は `OUTLOOK_MOCK=1` でモックに切り替え
