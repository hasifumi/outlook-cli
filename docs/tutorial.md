# outlook-cli チュートリアル

Outlook を開かずに、メール・添付・予定をターミナルから扱う練習。
**第1部は自宅のモック（本物の Outlook は不要）**、**第2部は会社 PC の実機確認チェックリスト**。

コマンドの一覧は [README.md](../README.md)、設計は [CLAUDE.md](../CLAUDE.md) を参照。

## 準備

```powershell
cd outlook-cli
uv sync
$env:OUTLOOK_MOCK = 1          # モックを使う（会社PCでは設定しない）
```

以降は `outlook <コマンド>`（`uv sync` 後）。インストールしない場合は
`.venv\Scripts\python.exe -m outlook_cli.cli <コマンド>`。

> **モックの決まり**: データは `mock_data.json`（メール3通・予定など）から**実行のたびに読み直す**。
> 送信・移動・保存の効果は次のコマンドには残らない。コマンドが通ること・出力の形を確かめる練習用。

---

## 第1部 モックで練習する

### 1. 読む

```powershell
outlook unread-count            # フォルダ別の未読件数
outlook list --limit 3          # 一覧。左端の ID（mock-001 など）を以降で使う
outlook read mock-002           # 本文
outlook search 経費 --days 30   # 検索
```

### 2. 添付ファイルを扱う

`mock-002`（Q1レポート提出のお願い）には添付が1つある。

```powershell
outlook attachments mock-002                 # 添付ファイル名の一覧
outlook save-attachment mock-002 --all --dest C:\tmp\att
outlook save-attachment mock-002 --all --dest C:\tmp\att     # もう一度
```

- 保存先は `--dest`（必須。無ければ作る）。ファイル名か `--all` のどちらかが必要。
- **同じ名前は上書きしない**。2回目は `Q1レポート様式_1.xlsx` のように連番になる。
- 出力は `保存しました: <パス>`。`--json-output` で機械可読にもなる。

### 3. メールを送る・返す・転送する・下書きにする

```powershell
outlook draft --to a@example.jp --subject 試験 --body こんにちは   # 送らず下書きに保存（ID を返す）
outlook send --to a@example.jp --cc b@example.jp --subject 試験 --body こんにちは --attach C:\tmp\a.txt
outlook reply mock-003 --body 承認します --all      # --all で全員に返信
outlook forward mock-002 --to a@example.jp --body ご確認ください   # 元の添付も引き継ぐ
```

迷ったら **まず `draft`**。下書きで確認してから `open <ID>` で Outlook 本体で開いて送れる。

### 4. 整理する

```powershell
outlook folders --recursive                       # 受信トレイ配下のフォルダ
outlook move mock-001 --folder "プロジェクト"      # 移動。移動後の ID が返る（以降はそれを使う）
```

`--folder` は `親/子` の形でも指定できる。

### 5. 予定を扱う

```powershell
outlook cal today
outlook cal week --date 2026-11-10
outlook categories                                 # 使えるカテゴリ名
outlook cal create --subject 打合せ --start 2026-11-10T10:00 --end 2026-11-10T11:00 --category 赤カテゴリ
outlook cal update <event_id> --location 会議室B   # 指定した項目だけ変わる
outlook cal delete <event_id> --yes                # --yes が無いと実行しない
outlook find-slot --attendees "a@co.jp;b@co.jp" --duration 60
```

- `cal create` に **`--attendee` を付けると出席者に招集メールが飛ぶ**。練習では付けない。
- `cal update` は `--attendee` を省略すると出席者にも通知にも触れない。
- `cal delete` は出席者のいる会議だと取り消し通知が送られる。
- 共有カレンダーは `--mailbox <相手のアドレス>`（`cal today` / `cal week`）。

### 6. テストを流す

```powershell
$env:OUTLOOK_MOCK = 1
uv run pytest -q          # 148件前後。数秒で終わる
```

機能ごとのテスト: 添付・移動 `tests/test_attachments_move.py`、送信系 `tests/test_send_family.py`、
予定 `tests/test_calendar_ops.py`、下書き `tests/test_drafts.py`。
`tests/test_interface_parity.py` は、本物の Outlook 用実装（`com.py`）がモックと同じメソッド・引数名を
持つかを、Outlook 無しで調べる。**新しいメソッドを足したら `base.py` `mock.py` `com.py` の3つに書く。**

---

## 第2部 会社 PC の実機確認（COM 層）

モックでは確かめられない部分。**自分宛・テスト用の相手**でだけ試し、他人に届く操作は先に下書きや
自分宛で確かめる。`OUTLOOK_MOCK` は**設定しない**（`Remove-Item Env:OUTLOOK_MOCK`）。

| # | 確認 | コマンド | 期待する結果 | 結果 |
|---|---|---|---|---|
| 1 | 読み取り | `outlook unread-count` / `list --limit 3` | モックでなく実際のメールが出る | |
| 2 | 添付一覧 | `attachments <添付付きメールID>` | 日本語名も文字化けしない | |
| 3 | 添付保存 | `save-attachment <ID> --all --dest C:\tmp\att` | 日本語名のまま保存。2回目は `_1` 連番 | |
| 4 | フォルダ | `folders --recursive` | 子フォルダが `親/子` で出る | |
| 5 | 移動 | `move <ID> --folder <フォルダ>` | 移動後の ID が返り、その ID で `read` できる（移動で ID が変わる） | |
| 6 | 下書き | `draft --to <自分> --subject 試験 --body x --attach <ファイル>` | 下書きフォルダに入り、添付が付く | |
| 7 | 自分宛送信 | `send --to <自分> --cc <自分> --subject 試験 --body x --attach <ファイル>` | 届く。CC・添付も付く | |
| 8 | 返信 | `reply <ID> --body x --all`（自分宛のメールで） | 全員返信になる | |
| 9 | 転送 | `forward <ID> --to <自分>` | 元の添付が付く | |
| 10 | 予定の作成 | `cal create ... --category <名前>`（出席者なし） | `cal today/week` に出る | |
| 11 | 予定の更新 | `cal update <ID> --location x` | 場所だけ変わる。出席者・通知はそのまま | |
| 12 | 予定の削除 | `cal delete <ID> --yes` | 消える | |
| 13 | 共有カレンダー | `cal week --mailbox <共有されている人>` | その人の予定が出る | |
| 14 | 共有なし | `cal week --mailbox <共有されていない人>` | 分かるエラーになる（落ちない） | |
| 15 | 本体で開く | `open <ID>` | Outlook 本体でそのメールが開く | |

うまくいかなかった行は、**エラーの種類（例外名）と手順番号だけ**を控えて持ち帰る
（メール本文・社内の名前は持ち出さない）。直す場所は `outlook_cli/com.py`。

## 次に読むもの

- [README.md](../README.md) — コマンド表・MCP サーバー・日報
- [CLAUDE.md](../CLAUDE.md) — 構成・移植メモ・実機確認の状況
- [SETUP_COMPANY.md](../SETUP_COMPANY.md) — 会社 PC のセットアップ
