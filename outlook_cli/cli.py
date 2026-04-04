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


if __name__ == "__main__":
    cli()
