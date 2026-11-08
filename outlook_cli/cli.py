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


def _echo_body(mail: dict) -> None:
    """--with-body 指定時（body キーがある時）だけ本文を字下げ表示"""
    if "body" not in mail:
        return
    for line in (mail["body"] or "").splitlines():
        click.echo(f"    {line}")
    click.echo()


@click.group()
def cli():
    """Outlook CLI ツール"""
    pass


@cli.command()
@click.option("--folder", default="inbox", help="フォルダ名 (inbox/sent/drafts)")
@click.option("--limit", default=20, help="取得件数")
@click.option("--with-body", is_flag=True, help="本文全文を含める")
@click.option("--json-output", is_flag=True, help="JSON出力")
def list(folder, limit, with_body, json_output):
    """メール一覧表示"""
    client = get_client()
    mails = client.list_mails(folder=folder, limit=limit, include_body=with_body)
    if json_output:
        click.echo(json.dumps(mails, ensure_ascii=False, indent=2))
    else:
        for m in mails:
            unread = "★" if m.get("unread") else "　"
            click.echo(f"{unread} [{m['id']}] {m['date'][:10]}  {m['from']:<30}  {m['subject']}")
            _echo_body(m)


@cli.command()
@click.argument("keyword")
@click.option("--days", default=7, help="検索対象の日数")
@click.option("--from", "sender", default=None, help="送信者フィルタ")
@click.option("--with-body", is_flag=True, help="本文全文を含める")
@click.option("--json-output", is_flag=True, help="JSON出力")
def search(keyword, days, sender, with_body, json_output):
    """メール検索"""
    client = get_client()
    mails = client.search(keyword=keyword, days=days, sender=sender, include_body=with_body)
    if json_output:
        click.echo(json.dumps(mails, ensure_ascii=False, indent=2))
    else:
        if not mails:
            click.echo("該当メールなし")
            return
        for m in mails:
            click.echo(f"[{m['id']}] {m['date'][:10]}  {m['from']:<30}  {m['subject']}")
            _echo_body(m)


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


# 添付するファイル。存在しないパスは送信前に弾く（COM に渡すため絶対パスにする）
attach_option = click.option(
    "--attach", "attach", multiple=True,
    type=click.Path(exists=True, dir_okay=False, resolve_path=True),
    help="添付ファイル（複数指定可）",
)


@cli.command()
@click.option("--to", required=True, help="宛先メールアドレス")
@click.option("--subject", required=True, help="件名")
@click.option("--body", required=True, help="本文")
@click.option("--cc", default="", help="Cc（セミコロン区切り）")
@attach_option
def send(to, subject, body, cc, attach):
    """メール送信"""
    client = get_client()
    client.send(to=to, subject=subject, body=body, cc=cc, attachments=[*attach] or None)
    click.echo("送信しました")


@cli.command()
@click.argument("mail_id")
@click.option("--body", required=True, help="返信本文")
@click.option("--all", "reply_all", is_flag=True, help="全員に返信")
@attach_option
def reply(mail_id, body, reply_all, attach):
    """メール返信"""
    client = get_client()
    client.reply(mail_id=mail_id, body=body, reply_all=reply_all, attachments=[*attach] or None)
    click.echo("返信しました")


@cli.command()
@click.argument("mail_id")
@click.option("--to", required=True, help="転送先メールアドレス")
@click.option("--body", default="", help="転送時に添える本文")
@attach_option
def forward(mail_id, to, body, attach):
    """メール転送（元メールの添付も引き継ぐ）"""
    client = get_client()
    try:
        client.forward(mail_id=mail_id, to=to, body=body, attachments=[*attach] or None)
    except ValueError as e:
        raise click.ClickException(str(e))
    click.echo("転送しました")


@cli.command()
@click.option("--to", required=True, help="宛先メールアドレス")
@click.option("--subject", required=True, help="件名")
@click.option("--body", required=True, help="本文")
@click.option("--cc", default="", help="Cc（セミコロン区切り）")
@attach_option
@click.option("--json-output", is_flag=True, help="JSON出力")
def draft(to, subject, body, cc, attach, json_output):
    """送信せずに下書きフォルダへ保存"""
    client = get_client()
    draft_id = client.save_draft(to=to, subject=subject, body=body, cc=cc, attachments=[*attach] or None)
    if json_output:
        click.echo(json.dumps({"id": draft_id}, ensure_ascii=False, indent=2))
        return
    click.echo(f"下書きに保存しました: {draft_id}")


@cli.command("open")
@click.argument("mail_id", required=False)
def open_cmd(mail_id):
    """Outlook 本体でメールを開く（ID 省略時は受信トレイ）"""
    get_client().open_mail(mail_id=mail_id)


@cli.command()
@click.argument("mail_id")
@click.option("--json-output", is_flag=True, help="JSON出力")
def attachments(mail_id, json_output):
    """メールの添付ファイル名を一覧表示"""
    client = get_client()
    try:
        names = client.read(mail_id).get("attachments", [])
    except ValueError as e:
        raise click.ClickException(str(e))
    if json_output:
        click.echo(json.dumps(names, ensure_ascii=False, indent=2))
        return
    if not names:
        click.echo("添付ファイルなし")
        return
    for name in names:
        click.echo(name)


@cli.command("save-attachment")
@click.argument("mail_id")
@click.argument("filename", required=False)
@click.option("--all", "save_all", is_flag=True, help="すべての添付ファイルを保存")
@click.option("--dest", required=True, type=click.Path(file_okay=False), help="保存先フォルダ（無ければ作成）")
@click.option("--json-output", is_flag=True, help="JSON出力")
def save_attachment(mail_id, filename, save_all, dest, json_output):
    """添付ファイルを保存（同名ファイルは上書きせず _1, _2 と連番）"""
    if filename and save_all:
        raise click.UsageError("ファイル名と --all は同時に指定できません")
    if not filename and not save_all:
        raise click.UsageError("添付ファイル名か --all を指定してください")
    client = get_client()
    dest_abs = os.path.abspath(dest)
    try:
        targets = client.read(mail_id).get("attachments", []) if save_all else [filename]
        saved = [
            {"filename": name, "path": client.save_mail_attachment(mail_id, name, dest_abs)}
            for name in targets
        ]
    except ValueError as e:
        raise click.ClickException(str(e))
    if json_output:
        click.echo(json.dumps(saved, ensure_ascii=False, indent=2))
        return
    if not saved:
        click.echo("添付ファイルなし")
        return
    for s in saved:
        click.echo(f"保存しました: {s['path']}")


@cli.command()
@click.argument("mail_id")
@click.option("--folder", required=True, help="移動先フォルダ（inbox/sent/drafts/trash または受信トレイ配下の 親/子）")
@click.option("--json-output", is_flag=True, help="JSON出力")
def move(mail_id, folder, json_output):
    """メールを別フォルダへ移動"""
    client = get_client()
    try:
        new_id = client.move_mail(mail_id, folder)
    except ValueError as e:
        raise click.ClickException(str(e))
    if json_output:
        click.echo(json.dumps({"id": new_id, "folder": folder}, ensure_ascii=False, indent=2))
        return
    click.echo(f"移動しました: {folder}")


@cli.command()
@click.option("--recursive", is_flag=True, help="孫フォルダ以下も 親/子 形式で含める")
@click.option("--json-output", is_flag=True, help="JSON出力")
def folders(recursive, json_output):
    """受信トレイ配下のフォルダ一覧"""
    client = get_client()
    names = client.list_subfolders(recursive=recursive)
    if json_output:
        click.echo(json.dumps(names, ensure_ascii=False, indent=2))
        return
    if not names:
        click.echo("サブフォルダなし")
        return
    for name in names:
        click.echo(name)


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


mailbox_option = click.option(
    "--mailbox", default=None,
    help="他の人の予定表を見る（メールアドレス。相手が共有している場合のみ）",
)


@cal.command("today")
@click.option("--date", default=None, help="対象日 YYYY-MM-DD（省略=今日）")
@mailbox_option
@click.option("--json-output", is_flag=True, help="JSON出力")
def cal_today(date, mailbox, json_output):
    """今日の予定一覧"""
    client = get_client()
    target = datetime.fromisoformat(date).date() if date else datetime.now().date()
    start = datetime(target.year, target.month, target.day)
    end = start + timedelta(days=1)
    appts = client.get_calendar(start, end, mailbox=mailbox)
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
@mailbox_option
@click.option("--json-output", is_flag=True, help="JSON出力")
def cal_week(date, mailbox, json_output):
    """今週（月〜金）の予定一覧"""
    client = get_client()
    base = datetime.fromisoformat(date).date() if date else datetime.now().date()
    monday = base - timedelta(days=base.weekday())
    start = datetime(monday.year, monday.month, monday.day)
    end = start + timedelta(days=5)
    appts = client.get_calendar(start, end, mailbox=mailbox)
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


DATETIME_TYPE = click.DateTime(formats=[
    "%Y-%m-%dT%H:%M", "%Y-%m-%d %H:%M", "%Y-%m-%dT%H:%M:%S", "%Y-%m-%d %H:%M:%S",
])
BUSY_STATUS_HELP = "0=空き 1=仮 2=予定あり 3=外出中 4=他の場所"


@cal.command("create")
@click.option("--subject", required=True, help="件名")
@click.option("--start", "start", required=True, type=DATETIME_TYPE, help="開始 YYYY-MM-DDTHH:MM")
@click.option("--end", "end", required=True, type=DATETIME_TYPE, help="終了 YYYY-MM-DDTHH:MM")
@click.option("--attendee", "attendees", multiple=True,
              help="出席者のメールアドレス（複数指定可）。指定すると会議招集メールが送信される")
@click.option("--location", default="", help="場所")
@click.option("--body", default="", help="本文")
@click.option("--busy-status", type=click.IntRange(0, 4), default=2, help=BUSY_STATUS_HELP)
@click.option("--reminder", "reminder_minutes", type=int, default=15,
              help="何分前に通知するか（負の値で通知なし）")
@click.option("--category", default="", help="カテゴリ名（`categories` で一覧）")
@attach_option
@click.option("--json-output", is_flag=True, help="JSON出力")
def cal_create(subject, start, end, attendees, location, body, busy_status,
               reminder_minutes, category, attach, json_output):
    """予定を作成（--attendee を付けると出席者に招集メールが送信される）"""
    if end <= start:
        raise click.UsageError("--end は --start より後にしてください")
    client = get_client()
    event_id = client.create_event(
        subject=subject, start=start, end=end, attendees=[*attendees] or None,
        location=location, body=body, busy_status=busy_status,
        reminder_minutes=reminder_minutes, category=category, attachments=[*attach] or None,
    )
    if json_output:
        click.echo(json.dumps({"id": event_id}, ensure_ascii=False, indent=2))
        return
    click.echo(f"予定を作成しました: {event_id}")


@cal.command("update")
@click.argument("event_id")
@click.option("--subject", default=None, help="件名")
@click.option("--start", "start", default=None, type=DATETIME_TYPE,
              help="開始 YYYY-MM-DDTHH:MM（--end 省略時は元の長さを保つ）")
@click.option("--end", "end", default=None, type=DATETIME_TYPE, help="終了 YYYY-MM-DDTHH:MM")
@click.option("--attendee", "attendees", multiple=True,
              help="出席者を置き換えて更新通知を送る。省略すると出席者は変更せず通知も送らない")
@click.option("--location", default=None, help="場所")
@click.option("--body", default=None, help="本文")
@click.option("--busy-status", type=click.IntRange(0, 4), default=None, help=BUSY_STATUS_HELP)
@click.option("--reminder", "reminder_minutes", type=int, default=None,
              help="何分前に通知するか（負の値で通知なし）")
@click.option("--category", default=None, help="カテゴリ名")
@attach_option
def cal_update(event_id, subject, start, end, attendees, location, body, busy_status,
               reminder_minutes, category, attach):
    """予定を更新（指定した項目だけを変更し、ほかは今の値のまま）"""
    client = get_client()
    try:
        current = client.get_event(event_id)
    except ValueError as e:
        raise click.ClickException(str(e))
    cur_start = datetime.fromisoformat(current["start"])
    cur_end = datetime.fromisoformat(current["end"])
    new_start = start or cur_start
    if end:
        new_end = end
    elif start:
        new_end = start + (cur_end - cur_start)    # 開始だけ動かしたら長さを保つ
    else:
        new_end = cur_end
    if new_end <= new_start:
        raise click.UsageError("終了は開始より後にしてください")
    try:
        client.update_event(
            event_id=event_id,
            subject=current["subject"] if subject is None else subject,
            start=new_start, end=new_end,
            attendees=[*attendees] or None,
            location=current["location"] if location is None else location,
            body=current["body"] if body is None else body,
            busy_status=current["busy_status"] if busy_status is None else busy_status,
            reminder_minutes=current["reminder_minutes"] if reminder_minutes is None else reminder_minutes,
            category=current["category"] if category is None else category,
            attachments=[*attach] or None,
        )
    except ValueError as e:
        raise click.ClickException(str(e))
    click.echo("予定を更新しました")


@cal.command("delete")
@click.argument("event_id")
@click.option("--yes", is_flag=True, help="削除を実行する（出席者がいる会議では取り消し通知が送られる）")
def cal_delete(event_id, yes):
    """予定を削除"""
    if not yes:
        raise click.UsageError("予定を削除します。実行するには --yes を付けてください")
    client = get_client()
    try:
        client.delete_event(event_id)
    except ValueError as e:
        raise click.ClickException(str(e))
    click.echo("予定を削除しました")


@cli.command()
@click.option("--json-output", is_flag=True, help="JSON出力")
def categories(json_output):
    """使えるカテゴリ名の一覧"""
    names = get_client().get_categories()
    if json_output:
        click.echo(json.dumps(names, ensure_ascii=False, indent=2))
        return
    for name in names:
        click.echo(name)


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
