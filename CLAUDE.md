# Outlook CLI/TUI — Claude 作業メモ

## プロジェクト概要

会社PC（Windows 11）のOutlook M365をターミナルから操作するCLI/TUIツール。
詳細設計は `DESIGN.md` を参照。

---

## 起動方法

```powershell
# 自宅（モック環境）
$env:OUTLOOK_MOCK=1; uv run python -m outlook_cli.tui   # TUI
$env:OUTLOOK_MOCK=1; uv run python -m outlook_cli.cli list  # CLI

# 会社PC（Outlook COM接続）
uv run python -m outlook_cli.tui
```

---

## アーキテクチャ

```
OutlookBase（抽象クラス: base.py）
├── OutlookMock（mock.py）  ← OUTLOOK_MOCK=1 のとき / mock_data.json を読む
└── OutlookCOM（com.py）    ← 会社PCのみ / win32com 経由
```

- `get_client()` が環境変数を見てどちらかを返す（tui.py / cli.py 共通）
- 新しいメソッドを追加するときは **base.py・mock.py・com.py の3箇所**に実装する

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

## ファイル構成

```
outlook-cli/
├── CLAUDE.md              ← 本ファイル
├── DESIGN.md              ← 設計仕様書（今後の予定も含む）
├── SETUP.md               ← セットアップ手順
├── mock_data.json         ← モックデータ（自宅開発用）
├── pyproject.toml
└── outlook_cli/
    ├── base.py            ← 抽象インターフェース
    ├── mock.py            ← モック実装
    ├── com.py             ← Outlook COM実装
    ├── cli.py             ← CLIコマンド（Click）
    └── tui.py             ← TUIアプリ（Textual 8.x）
```

---

## 今後の実装予定

詳細は `DESIGN.md` の「今後の予定」セクション参照。

- ファイル添付機能（Ctrl+A でパス入力）
- 連絡先キャッシュ（contacts_cache.json）
- Neovimキーマップ設定
- fzf連携
