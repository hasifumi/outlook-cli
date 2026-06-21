import os
from datetime import date

import click
from dotenv import load_dotenv

load_dotenv()


def _get_client():
    if os.getenv("OUTLOOK_MOCK"):
        from outlook_cli.mock import OutlookMock
        return OutlookMock()
    from outlook_cli.com import OutlookCOM
    return OutlookCOM()


@click.command()
@click.option("--output-dir", default=None, help="出力先ディレクトリ（省略時は .env の DAILY_REPORT_OUTPUT_DIR）")
@click.option("--date", "target_date", default=None, help="対象日 YYYY-MM-DD（省略時は今日）")
@click.option("--no-notify", is_flag=True, help="Windows通知を抑制")
def main(output_dir, target_date, no_notify):
    """日報を自動生成して保存する"""
    from daily_report.collect import collect_calendar, collect_sent_mail, collect_activitywatch, collect_mattermost
    from daily_report.llm import ensure_api_server, call_llm
    from daily_report.report import save_report, notify_done

    target = date.fromisoformat(target_date) if target_date else date.today()
    out_dir = output_dir or os.environ.get("DAILY_REPORT_OUTPUT_DIR", ".")
    api_endpoint = os.environ.get("DAILY_REPORT_API_ENDPOINT", "http://localhost:8089/v1/chat/completions")
    health_url = os.environ.get("DAILY_REPORT_API_HEALTH_URL", "http://localhost:8089/v1/models")
    start_cmd = os.environ.get("DAILY_REPORT_API_START_CMD", "")

    click.echo("APIサーバを確認中...")
    ensure_api_server(health_url, start_cmd)

    click.echo("データ収集中...")
    client = _get_client()
    cal = collect_calendar(client, target)
    comm = collect_sent_mail(client, target)
    mm = collect_mattermost()
    comm_full = comm + ("\n\n" + mm if mm else "")
    aw = collect_activitywatch(target)

    click.echo("LLMで日報生成中...")
    text = call_llm(api_endpoint, cal, comm_full, aw)

    path = save_report(out_dir, text, target)
    click.echo(f"保存しました: {path}")

    if not no_notify:
        notify_done(path)


if __name__ == "__main__":
    main()
