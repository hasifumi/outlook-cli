# 会社PC セットアップ手順

## 1. コピーするファイル

以下のフォルダ構成ごと会社PCに持ち込む（USBまたはGit経由）。

```
outlook-cli/                ← このフォルダごとコピー
├── pyproject.toml          ← 依存パッケージ定義（必須）
├── uv.lock                 ← バージョン固定ファイル（必須）
└── outlook_cli/            ← アプリ本体（必須）
    ├── __init__.py
    ├── base.py
    ├── cli.py
    ├── com.py
    ├── mock.py
    └── tui.py
```

**不要なファイル（コピーしなくてよい）：**

| ファイル | 理由 |
|----------|------|
| `mock_data.json` | モック用。会社PCでは使わない |
| `.venv/` | 会社PCで `uv sync` し直すため |
| `__pycache__/` | 自動生成されるため |

---

## 2. uvのインストール（会社PCに入っていない場合）

```powershell
winget install astral-sh.uv
```

インストール後、PowerShellを再起動して `uv --version` で確認。

---

## 3. 依存パッケージのインストール

```powershell
cd outlook-cli
uv sync
uv add pywin32
```

---

## 4. pywin32ポストインストール（必須）

これをやらないとOutlook COMが動かない。

```powershell
# pywin32_postinstall.py の場所を確認
Get-ChildItem -Recurse -Filter "pywin32_postinstall.py" .venv\

# 見つかったパスで実行（例）
.venv\Scripts\python.exe .venv\Scripts\pywin32_postinstall.py -install
```

---

## 5. PowerShellプロファイルにエイリアス登録

```powershell
notepad $PROFILE
```

以下を追記（`<ユーザー名>` は `echo $env:USERPROFILE` で確認）：

```powershell
function outlook-cli {
    & "C:\Users\<ユーザー名>\outlook-cli\.venv\Scripts\python.exe" -m outlook_cli.cli @args
}
function outlook-tui {
    & "C:\Users\<ユーザー名>\outlook-cli\.venv\Scripts\python.exe" -m outlook_cli.tui
}
```

保存後：

```powershell
. $PROFILE
```

---

## 6. Outlookを起動・ログイン済みにしておく

COM経由の操作にはOutlookプロセスが必要。
タスクトレイ常駐でもOK（設定方法：ファイル→オプション→詳細設定→「終了時にタスクバーの通知領域に格納する」をチェック）。

---

## 7. 動作確認

```powershell
# CLIテスト
outlook-cli list
outlook-cli search "承認" --days 30

# TUIテスト
outlook-tui
```

---

## トラブルシューティング

### `No module named 'outlook_cli'`

```powershell
Remove-Item -Recurse -Force .venv
uv sync
uv add pywin32
```

### `pywin32_postinstall.py が見つからない`

```powershell
Get-ChildItem -Recurse -Filter "pywin32_postinstall.py" .venv\
```

### Outlook COM接続エラー

Outlookが起動・ログイン済みか確認。起動していない場合はCOMが動かない。

### 「このプログラムがOutlookにアクセスしようとしています」ダイアログ

初回のみ表示される場合がある。「許可」を押してOK。
繰り返し出る場合は Outlook のセキュリティ設定（ファイル→オプション→セキュリティセンター→プログラムによるアクセス）を確認。
