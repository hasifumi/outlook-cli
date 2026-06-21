from datetime import date
from pathlib import Path


def save_report(output_dir: str, text: str, target: date) -> Path:
    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    path = out / f"{target.isoformat()}.md"
    path.write_text(text, encoding="utf-8")
    return path


def notify_done(path: Path) -> None:
    try:
        from win10toast import ToastNotifier
        ToastNotifier().show_toast(
            "日報作成完了",
            f"{path.name} を保存しました",
            duration=5,
            threaded=True,
        )
    except Exception:
        pass
