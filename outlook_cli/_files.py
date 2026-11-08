"""添付ファイル保存用のファイル名ユーティリティ（Outlook に依存しない純粋関数）"""
import re
from pathlib import Path

# Windows のファイル名に使えない文字と制御文字
_ILLEGAL = re.compile(r'[\\/:*?"<>|\x00-\x1f]')


def safe_filename(name: str) -> str:
    """ファイル名として使えない文字を "_" に置き換える（空・"." のみなら "attachment"）"""
    cleaned = _ILLEGAL.sub("_", name or "").strip().rstrip(". ")
    return cleaned or "attachment"


def unique_path(dest_dir, filename: str) -> Path:
    """dest_dir 内で未使用のパスを返す。同名があれば a.xlsx -> a_1.xlsx -> a_2.xlsx と連番を付ける"""
    dest = Path(dest_dir)
    candidate = dest / filename
    if not candidate.exists():
        return candidate
    stem, suffix = candidate.stem, candidate.suffix
    n = 1
    while True:
        candidate = dest / f"{stem}_{n}{suffix}"
        if not candidate.exists():
            return candidate
        n += 1
