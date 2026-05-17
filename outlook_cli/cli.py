import json
import math
import os
from datetime import datetime, timedelta

import click

from .mock import OutlookMock

# 環境変数で切り替え（会社PCではOUTLOOK_MOCK未設定）
def get_client():
    if os.getenv("OUTLOOK_MOCK"):
        return OutlookMock()
    else:
        from .com import OutlookCOM
        return OutlookCOM()


WEEKDAY_JA = ["月", "火", "水", "木", "金", "土", "日"]

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


@cli.command("flagged")
@click.option("--folder", default="inbox", help="フォルダ名")
@click.option("--days", default=7, help="受信期間（日数）")
@click.option("--json-output", is_flag=True, help="JSON出力")
def flagged(folder, days, json_output):
    """重要フラグ／期限設定メールを検索"""
    client = get_client()
    mails = client.flagged_or_due(days=days, folder=folder)
    if json_output:
        click.echo(json.dumps(mails, ensure_ascii=False, indent=2))
        return
    if not mails:
        click.echo("該当メールなし")
        return
    for m in mails:
        flag_mark = "[F]" if m.get("flag_status") == 1 else "   "
        due = f" [期限: {m['due_date']}]" if m.get("due_date") else ""
        click.echo(f"{flag_mark} [{m['date'][:10]}] {m['from']:<30}  {m['subject']}{due}")


def _format_appointment(appt: dict) -> str:
    start_dt = datetime.fromisoformat(appt["start"])
    end_dt = datetime.fromisoformat(appt["end"])
    if appt.get("all_day"):
        time_str = "終日"
    else:
        time_str = f"{start_dt.strftime('%H:%M')}-{end_dt.strftime('%H:%M')}"
    loc = f"  @{appt['location']}" if appt.get("location") else ""
    return f"  {time_str:<13}  {appt['subject']}{loc}"


def _compute_free_slots(
    freebusy_map: dict,
    search_start: datetime,
    days: int,
    duration_min: int,
    work_start: int,
    work_end: int,
) -> list:
    slot_minutes = 30
    slots_per_duration = math.ceil(duration_min / slot_minutes)
    now = datetime.now()
    total_slots = days * 24 * 60 // slot_minutes
    candidates = []

    for i in range(total_slots - slots_per_duration + 1):
        slot_dt = search_start + timedelta(minutes=i * slot_minutes)
        if slot_dt < now:
            continue
        if slot_dt.weekday() >= 5:
            continue
        if not (work_start <= slot_dt.hour < work_end):
            continue
        end_dt = slot_dt + timedelta(minutes=duration_min)
        if end_dt.hour > work_end or (end_dt.hour == work_end and end_dt.minute > 0):
            continue

        block_ok = True
        tentative_emails = []
        for email, fb in freebusy_map.items():
            for j in range(slots_per_duration):
                idx = i + j
                if idx >= len(fb):
                    block_ok = False
                    break
                status = fb[idx]
                if status in ("2", "3"):
                    block_ok = False
                    break
                if status == "1" and email not in tentative_emails:
                    tentative_emails.append(email)
            if not block_ok:
                break

        if block_ok:
            candidates.append({
                "start":     slot_dt.isoformat()[:16],
                "end":       end_dt.isoformat()[:16],
                "tentative": tentative_emails,
            })

    candidates.sort(key=lambda c: len(c["tentative"]))
    return candidates[:5]


@cli.group()
def cal():
    """カレンダー参照"""
    pass


@cal.command("today")
@click.option("--date", default=None, help="対象日 YYYY-MM-DD（省略=今日）")
@click.option("--json-output", is_flag=True, help="JSON出力")
def cal_today(date, json_output):
    """今日の予定一覧"""
    client = get_client()
    target = datetime.fromisoformat(date).date() if date else datetime.now().date()
    start = datetime(target.year, target.month, target.day)
    end = start + timedelta(days=1)
    appts = client.get_calendar(start, end)
    if json_output:
        click.echo(json.dumps(appts, ensure_ascii=False, indent=2))
        return
    wd = WEEKDAY_JA[start.weekday()]
    click.echo(f"{target.strftime('%Y-%m-%d')} ({wd}) の予定")
    click.echo("─" * 38)
    if not appts:
        click.echo("  予定なし")
        return
    for appt in appts:
        click.echo(_format_appointment(appt))


@cal.command("week")
@click.option("--date", default=None, help="週内の任意日 YYYY-MM-DD（省略=今週）")
@click.option("--json-output", is_flag=True, help="JSON出力")
def cal_week(date, json_output):
    """今週（月〜金）の予定一覧"""
    client = get_client()
    base = datetime.fromisoformat(date).date() if date else datetime.now().date()
    monday = base - timedelta(days=base.weekday())
    start = datetime(monday.year, monday.month, monday.day)
    end = start + timedelta(days=5)
    appts = client.get_calendar(start, end)
    if json_output:
        click.echo(json.dumps(appts, ensure_ascii=False, indent=2))
        return
    week_str = f"{monday.strftime('%Y-%m-%d')} 週"
    click.echo(f"{week_str} の予定")
    click.echo("─" * 38)
    if not appts:
        click.echo("  予定なし")
        return
    current_day = None
    for appt in appts:
        day = appt["start"][:10]
        if day != current_day:
            current_day = day
            day_dt = datetime.fromisoformat(day)
            wd = WEEKDAY_JA[day_dt.weekday()]
            click.echo(f"\n{day} ({wd})")
        click.echo(_format_appointment(appt))


@cal.command("range")
def cal_range():
    """(未実装) 任意期間の予定"""
    raise click.ClickException("cal range は未実装です")


@cli.command("find-slot")
@click.option("--attendees", required=True, help="セミコロン区切りメールアドレス")
@click.option("--duration", default=60, help="必要な時間（分）")
@click.option("--days", default=5, help="何日先まで探すか")
@click.option("--work-start", default=9, help="業務開始時刻（時）")
@click.option("--work-end", default=18, help="業務終了時刻（時）")
@click.option("--json-output", is_flag=True, help="JSON出力")
def find_slot(attendees, duration, days, work_start, work_end, json_output):
    """複数人の空き時間候補を検索"""
    client = get_client()
    emails = [e.strip() for e in attendees.split(";") if e.strip()]
    search_start = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)

    freebusy_map = {}
    for email in emails:
        fb = client.get_freebusy(email, search_start, 30)
        if fb == "":
            click.echo(f"警告: {email} の空き情報を取得できませんでした（除外して探索）", err=True)
        else:
            freebusy_map[email] = fb

    if not freebusy_map:
        raise click.ClickException("空き情報を取得できる参加者がいません")

    candidates = _compute_free_slots(freebusy_map, search_start, days, duration, work_start, work_end)

    if json_output:
        click.echo(json.dumps(candidates, ensure_ascii=False, indent=2))
        return

    click.echo(f"空き時間候補（{duration}分）")
    click.echo("─" * 38)
    if not candidates:
        click.echo("候補が見つかりませんでした")
        return
    for i, c in enumerate(candidates, 1):
        start_dt = datetime.fromisoformat(c["start"])
        wd = WEEKDAY_JA[start_dt.weekday()]
        date_str = f"{c['start'][:10]} ({wd}) {c['start'][11:]}-{c['end'][11:]}"
        if c["tentative"]:
            tent_str = "  ※仮予定あり: " + ", ".join(c["tentative"])
            mark = "△"
        else:
            tent_str = ""
            mark = "○"
        click.echo(f"{i}. {mark}  {date_str}{tent_str}")


if __name__ == "__main__":
    cli()
