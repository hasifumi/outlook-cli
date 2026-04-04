# セットアップ手順書

## 自宅PC（モック動作確認）

### 1. フォルダ構成確認

```
outlook-cli\
├── DESIGN.md
├── SETUP.md              ← 本ファイル
├── pyproject.toml
├── mock_data.json
└── outlook_cli\
    ├── __init__.py
    ├── base.py
    ├── cli.py
    ├── com.py
    ├── mock.py
    └── tui.py
```

### 2. 環境構築

```powershell
cd outlook-cli
Remove-Item -Recurse -Force .venv   # 既存の.venvがある場合
uv sync
```

### 3. PowerShellプロファイルにエイリアス登録

```powershell
notepad $PROFILE
```

以下を追記：

```powershell
function outlook-cli {
    $env:OUTLOOK_MOCK=1
    & "C:\Users\<ユーザー名>\project\outlook-cli\.venv\Scripts\python.exe" -m outlook_cli.cli @args
}
function outlook-tui {
    $env:OUTLOOK_MOCK=1
    & "C:\Users\<ユーザー名>\project\outlook-cli\.venv\Scripts\python.exe" -m outlook_cli.tui
}
```

※ `<ユーザー名>` は `echo $env:USERPROFILE` で確認

保存後：

```powershell
. $PROFILE
```

### 4. 動作確認

```powershell
# CLIテスト
outlook-cli list
outlook-cli search "承認" --days 30
outlook-cli read mock-003

# TUIテスト
outlook-tui
```

---

## 会社PC（Outlook COM動作確認）

### 1. ファイル持ち込み

USBまたはGit経由で`outlook-cli`フォルダを会社PCに配置。

### 2. 環境構築

```powershell
cd outlook-cli
uv sync
uv add pywin32
```

### 3. pywin32ポストインストール（必須）

```powershell
# ファイルの場所を確認
Get-ChildItem -Recurse -Filter "pywin32_postinstall.py" .venv\

# 実行（見つかったパスで）
.venv\Scripts\python.exe .venv\Scripts\pywin32_postinstall.py -install
```

### 4. PowerShellプロファイルにエイリアス登録

```powershell
notepad $PROFILE
```

以下を追記（OUTLOOK_MOCKなし）：

```powershell
function outlook-cli {
    & "C:\Users\<ユーザー名>\project\outlook-cli\.venv\Scripts\python.exe" -m outlook_cli.cli @args
}
function outlook-tui {
    & "C:\Users\<ユーザー名>\project\outlook-cli\.venv\Scripts\python.exe" -m outlook_cli.tui
}
```

保存後：

```powershell
. $PROFILE
```

### 5. Outlookを起動・ログイン済みにしておく

### 6. 動作確認

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
```

### `ImportError: OutlookCOMはWindowsのみ対応`

`$env:OUTLOOK_MOCK=1` が設定されているか確認。

### `pywin32_postinstall.py が見つからない`

```powershell
Get-ChildItem -Recurse -Filter "pywin32_postinstall.py" .venv\
```

### Outlook COM接続エラー

Outlookが起動・ログイン済みか確認。起動していない場合はCOMが動きません。
