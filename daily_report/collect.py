import json
import os
import sqlite3
from datetime import date, datetime, timedelta
from pathlib import Path

import requests

WEEKDAY_JA = ["月", "火", "水", "木", "金", "土", "日"]


def collect_calendar(client, target: date) -> str:
    today_start = datetime(target.year, target.month, target.day)
    today_end = today_start + timedelta(days=1)
    tomorrow_start = today_end
    tomorrow_end = tomorrow_start + timedelta(days=1)

    today_events = client.get_calendar(today_start, today_end)
    tomorrow_events = client.get_calendar(tomorrow_start, tomorrow_end)

    tomorrow = target + timedelta(days=1)
    lines = [
        f"=== 今日の予定 ({target.isoformat()} {WEEKDAY_JA[target.weekday()]}) ===",
    ]
    lines += _format_events(today_events)
    lines += [
        f"\n=== 明日の予定 ({tomorrow.isoformat()} {WEEKDAY_JA[tomorrow.weekday()]}) ===",
    ]
    lines += _format_events(tomorrow_events)
    return "\n".join(lines)


def _format_events(events: list) -> list[str]:
    if not events:
        return ["  予定なし"]
    result = []
    for e in events:
        if e.get("all_day"):
            time_str = "終日"
        else:
            time_str = f"{e['start'][11:16]}-{e['end'][11:16]}"
        loc = f" @{e['location']}" if e.get("location") else ""
        result.append(f"  {time_str}  {e['subject']}{loc}")
    return result


def collect_sent_mail(client, target: date) -> str:
    mails = client.sent_today(date=target.isoformat())
    lines = [f"=== 送信メール ({target.isoformat()}) ==="]
    if mails:
        for m in mails:
            time_str = m["date"][11:16] if len(m["date"]) > 10 else ""
            lines.append(f"  [{time_str}] 宛先: {m.get('to', '')}  件名: {m['subject']}")
    else:
        lines.append("  送信なし")
    return "\n".join(lines)


def collect_activitywatch(target: date) -> str:
    try:
        return _aw_from_rest(target)
    except Exception:
        pass
    try:
        return _aw_from_sqlite(target)
    except Exception:
        return ""


def _aw_from_rest(target: date) -> str:
    base = "http://localhost:5600/api/0"
    resp = requests.get(f"{base}/buckets", timeout=3)
    resp.raise_for_status()
    buckets = resp.json()
    window_bucket = next(
        (b["id"] for b in buckets if "watcher-window" in b["id"]), None
    )
    if not window_bucket:
        raise ValueError("window bucket not found")

    resp = requests.get(
        f"{base}/buckets/{window_bucket}/events",
        params={"start": f"{target.isoformat()}T00:00:00", "end": f"{target.isoformat()}T23:59:59"},
        timeout=5,
    )
    resp.raise_for_status()
    return _format_aw_events(resp.json(), target)


def _aw_from_sqlite(target: date) -> str:
    db_path = os.environ.get("ACTIVITYWATCH_DB") or str(
        Path(os.environ.get("LOCALAPPDATA", ""))
        / "activitywatch/activitywatch/aw-server/peewee-sqlite.v2.db"
    )
    con = sqlite3.connect(db_path)
    try:
        cur = con.execute(
            """
            SELECT e.datastr, e.duration
            FROM eventmodel e
            JOIN bucketmodel b ON e.bucketrow_id = b.id
            WHERE b.bucketid LIKE 'aw-watcher-window_%'
              AND e.timestamp >= ? AND e.timestamp < ?
            """,
            (f"{target.isoformat()}T00:00:00", f"{target.isoformat()}T23:59:59"),
        )
        events = [{"data": json.loads(r[0]), "duration": r[1]} for r in cur.fetchall()]
    finally:
        con.close()
    return _format_aw_events(events, target)


def _format_aw_events(events: list, target: date) -> str:
    totals: dict[str, float] = {}
    for e in events:
        app = e["data"].get("app", "Unknown")
        totals[app] = totals.get(app, 0) + e.get("duration", 0)

    lines = [f"=== 作業ログ ({target.isoformat()}) ==="]
    if not totals:
        lines.append("  データなし")
        return "\n".join(lines)

    for app, secs in sorted(totals.items(), key=lambda x: x[1], reverse=True)[:10]:
        h, m = divmod(int(secs) // 60, 60)
        lines.append(f"  {app:<30} {h}h{m:02d}m")
    return "\n".join(lines)


def collect_mattermost() -> str:
    return ""
