# Outlook CLI / TUI 設計仕様書

## 概要

会社PCのOutlook（M365 Apps for Enterprise）をCLI・TUIから操作するツール。
Outlookデスクトップをバックグラウンド常駐させ、GUIを一切触らずメール操作を行う。

---

## 背景・目的

- 会社PC（RAM 8GB）でOutlookのGUIが重い
- Thoriumブラウザ導入でEdgeを置き換え、節約したメモリでOllama（1.5B〜E2B級）を動かしたい
- メール操作はテキストベースで完結させ、Neovimターミナルから呼び出す

---

## 環境

| 項目 | 内容 |
|------|------|
| 会社PC OS | Windows 11 |
| Outlook | M365 Apps for Enterprise バージョン2512 ビルド19530.20226（クイック実行） |
| Python管理 | uv（winget経由でインストール） |
| ターミナル | PowerShell / Windows Terminal |
| エディタ | Neovim |

---

## アーキテクチャ

```
Neovim ターミナル
  │
  ├── outlook-cli （Click CLIコマンド）
  │     └── OutlookBase（抽象クラス）
  │           ├── OutlookMock  ← 自宅開発用（JSONファイル）
  │           └── OutlookCOM   ← 会社PC用（win32com）
  │
  └── outlook-tui （Textual TUIアプリ）
        └── 同じOutlookBase経由で操作
```

### 環境切り替え

```powershell
# 自宅（モック）
$env:OUTLOOK_MOCK=1; .venv\Scripts\python.exe -m outlook_cli.tui

# 会社PC（COM）
.venv\Scripts\python.exe -m outlook_cli.tui
```

---

## フォルダ構成

```
outlook-cli/
├── DESIGN.md              ← 本ファイル
├── pyproject.toml
├── mock_data.json         ← モックデータ（自宅開発用）
└── outlook_cli/
    ├── __init__.py
    ├── base.py            ← 抽象インターフェース
    ├── mock.py            ← モック実装
    ├── com.py             ← Outlook COM実装（会社PC用）
    ├── cli.py             ← CLIコマンド（Click）
    └── tui.py             ← TUIアプリ（Textual）
```

---

## CLIコマンド仕様

```powershell
# メール一覧
outlook list [--folder inbox|sent|drafts] [--limit 20] [--json-output]

# メール検索
outlook search <keyword> [--days 7] [--from <address>] [--json-output]

# メール本文表示
outlook read <mail_id> [--json-output]

# メール送信
outlook send --to <address> --subject <subject> --body <body>

# メール返信
outlook reply <mail_id> --body <body>
```

---

## TUI仕様

### 画面構成（3ペイン）

```
┌─ Outlook TUI ──────────────────────────────────────────────────┐
│ [N]新規  [R]返信  [F]転送  [/]検索  [U]未読切替  [Q]終了       │
├─────────────────┬──────────────────────────────────────────────┤
│ 受信トレイ (2)  │★ 2026-03-28  田中 一郎    来週の定例会議     │
│ 送信済み        │　 2026-03-27  鈴木 花子    Q1レポート提出     │
│ 下書き          │★ 2026-03-26  山田 太郎    承認依頼：経費精算 │
│ ゴミ箱          │　 2026-03-25  ITサポート   メンテナンス通知  │
│                 ├──────────────────────────────────────────────┤
│ プロジェクトA(1)│ 件名: 来週の定例会議について                  │
│ 社内連絡        │ From: 田中 一郎 <tanaka@company.com>          │
│                 │ ────────────────────────────────────────────  │
│                 │ お疲れ様です。来週月曜の定例会議ですが...      │
└─────────────────┴──────────────────────────────────────────────┘
```

### キーバインド

| キー | 動作 |
|------|------|
| `N` | 新規作成画面を開く |
| `R` | 返信画面を開く |
| `F` | 転送画面を開く |
| `U` | 未読/既読トグル |
| `D` | 削除 |
| `/` | 検索バー表示 |
| `Ctrl+R` | メール一覧更新 |
| `Tab` | フォルダペイン↔メール一覧ペイン切替 |
| `↑↓` | フォルダ選択 / メール選択 |
| `Enter` | フォルダ選択確定 / プレビュー表示 |
| `Esc` | 検索解除 / 画面を戻る |
| `Q` | 終了 |

### 新規作成・返信・転送画面

```
┌─ 新規メール ──────────────────────────────────────────┐
│ TO : [田中▌]                                          │
│      ┌──────────────────────────┐                    │
│      │▶ 田中 一郎 tanaka@co.jp  │  ← ドロップダウン  │
│      │  田中 花子 hanako@co.jp  │                    │
│      └──────────────────────────┘                    │
│ CC : [                         ]                      │
│ BCC: [                         ]                      │
│ 件名: [                        ]                      │
│ ────────────────────────────────────────────────────  │
│ 本文:                                                  │
│                                                        │
│ [Ctrl+Enter 送信]  [Esc キャンセル]                   │
└───────────────────────────────────────────────────────┘
```

- TOフィールドは2文字以上でインクリメンタルサーチ（社内連絡先）
- 上下キーで候補選択、Enterで確定
- セミコロン区切りで複数アドレス入力可能
- 返信時：TO自動入力・件名に「Re:」付与・本文に引用
- 転送時：件名に「Fw:」付与・本文に元メール内容

---

## 連絡先取得

| 環境 | 取得元 |
|------|--------|
| モック | mock_data.json の contacts配列 |
| 会社PC | OutlookのグローバルアドレスリストをCOM経由で取得 |

---

## 今後の予定

- [ ] ファイル添付機能（Ctrl+A でパス入力）
- [ ] 連絡先キャッシュの定期更新（contacts_cache.json）
- [ ] Neovimキーマップ設定
- [ ] fzf連携（メール一覧絞り込み）
