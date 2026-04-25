import json
import os

import click

from .mock import OutlookMock

# 環境変数で切り替え（会社PCではOUTLOOK_MOCK未設定）
def get_client():
    if os.getenv("OUTLOOK_MOCK"):
        return OutlookMock()
    else:
        from .com import OutlookCOM
        return OutlookCOM()


FOLDER_LABELS = {
    "inbox":  "受信トレイ",
    "sent":   "送信済み",
    "drafts": "下書き",
    "trash":  "ゴミ箱",
}


@click.group()
def cli():
    """Outlook CLI ツール"""
    pass


@cli.command()
@click.option("--folder", default="inbox", help="フォルダ名 (inbox/sent/drafts)")
@click.option("--limit", default=20, help="取得件数")
@click.option("--json-output", is_flag=True, help="JSON出力")
def list(folder, limit, json_output):
    """メール一覧表示"""
    client = get_client()
    mails = client.list_mails(folder=folder, limit=limit)
    if json_output:
        click.echo(json.dumps(mails, ensure_ascii=False, indent=2))
    else:
        for m in mails:
            unread = "★" if m.get("unread") else "　"
            click.echo(f"{unread} [{m['id']}] {m['date'][:10]}  {m['from']:<30}  {m['subject']}")


@cli.command()
@click.argument("keyword")
@click.option("--days", default=7, help="検索対象の日数")
@click.option("--from", "sender", default=None, help="送信者フィルタ")
@click.option("--json-output", is_flag=True, help="JSON出力")
def search(keyword, days, sender, json_output):
    """メール検索"""
    client = get_client()
    mails = client.search(keyword=keyword, days=days, sender=sender)
    if json_output:
        click.echo(json.dumps(mails, ensure_ascii=False, indent=2))
    else:
        if not mails:
            click.echo("該当メールなし")
            return
        for m in mails:
            click.echo(f"[{m['id']}] {m['date'][:10]}  {m['from']:<30}  {m['subject']}")


@cli.command()
@click.argument("mail_id")
@click.option("--json-output", is_flag=True, help="JSON出力")
def read(mail_id, json_output):
    """メール本文表示"""
    client = get_client()
    mail = client.read(mail_id)
    if json_output:
        click.echo(json.dumps(mail, ensure_ascii=False, indent=2))
    else:
        click.echo(f"件名  : {mail['subject']}")
        click.echo(f"送信者: {mail['from']}")
        click.echo(f"宛先  : {mail['to']}")
        click.echo(f"日時  : {mail['date']}")
        click.echo("-" * 40)
        click.echo(mail["body"])


@cli.command()
@click.option("--to", required=True, help="宛先メールアドレス")
@click.option("--subject", required=True, help="件名")
@click.option("--body", required=True, help="本文")
def send(to, subject, body):
    """メール送信"""
    client = get_client()
    client.send(to=to, subject=subject, body=body)
    click.echo("送信しました")


@cli.command()
@click.argument("mail_id")
@click.option("--body", required=True, help="返信本文")
def reply(mail_id, body):
    """メール返信"""
    client = get_client()
    client.reply(mail_id=mail_id, body=body)
    click.echo("返信しました")


@cli.command("unread-count")
@click.option("--folder", default=None, help="フォルダ名（省略時は全フォルダ）")
@click.option("--json-output", is_flag=True, help="JSON出力")
def unread_count(folder, json_output):
    """未読件数表示"""
    client = get_client()
    result = client.unread_count(folder=folder)
    if json_output:
        click.echo(json.dumps(result, ensure_ascii=False, indent=2))
        return
    total = result.pop("total")
    max_len = max((len(FOLDER_LABELS.get(k, k)) for k in result), default=0)
    for key, count in result.items():
        label = FOLDER_LABELS.get(key, key)
        click.echo(f"{label:{max_len}}: {count:>3}件")
    click.echo("---")
    click.echo(f"{'合計':{max_len}}: {total:>3}件")


@cli.command("unread-summary")
@click.option("--folder", default="inbox", help="フォルダ名（省略時は受信トレイ）")
@click.option("--limit", default=10, help="取得件数")
@click.option("--json-output", is_flag=True, help="JSON出力")
def unread_summary(folder, limit, json_output):
    """未読メールサマリー表示"""
    client = get_client()
    mails = client.unread_summary(limit=limit, folder=folder)
    if json_output:
        click.echo(json.dumps(mails, ensure_ascii=False, indent=2))
        return
    if not mails:
        click.echo("未読メールはありません")
        return
    for m in mails:
        click.echo(f"[{m['date'][:16]}] {m['from']:<30}  {m['subject']}")
        click.echo(f"  {m['preview']}")
        click.echo()


@cli.command("sent-today")
@click.option("--date", default=None, help="YYYY-MM-DD（省略時は今日）")
@click.option("--json-output", is_flag=True, help="JSON出力")
def sent_today(date, json_output):
    """当日の送信メール一覧表示"""
    client = get_client()
    mails = client.sent_today(date=date)
    if json_output:
        click.echo(json.dumps(mails, ensure_ascii=False, indent=2))
        return
    if not mails:
        click.echo("送信メールはありません")
        return
    for m in mails:
        click.echo(f"[{m['date'][11:16]}] {m.get('to', ''):<40}  {m['subject']}")


if __name__ == "__main__":
    cli()
