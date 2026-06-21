import subprocess
import time

import requests

SYSTEM_PROMPT = """\
あなたは日報作成アシスタントです。
以下のデータをもとに日本語で日報を作成してください。

## やったこと
- （箇条書き）

## 明日やること
- （明日の予定・継続タスク）

## 所感
（1〜3文）

余計な説明は不要。日報本文のみ出力してください。
"""


def ensure_api_server(health_url: str, start_cmd: str, timeout: int = 30) -> None:
    try:
        r = requests.get(health_url, timeout=3)
        if r.ok:
            return
    except Exception:
        pass

    if not start_cmd:
        raise RuntimeError(f"APIサーバに接続できません: {health_url}")

    flags = getattr(subprocess, "DETACHED_PROCESS", 0)
    subprocess.Popen(["powershell", "-Command", start_cmd], creationflags=flags)

    deadline = time.time() + timeout
    while time.time() < deadline:
        time.sleep(2)
        try:
            r = requests.get(health_url, timeout=3)
            if r.ok:
                return
        except Exception:
            pass

    raise TimeoutError(f"APIサーバが{timeout}秒以内に起動しませんでした: {health_url}")


def call_llm(endpoint: str, cal: str, comm: str, aw: str) -> str:
    payload = {
        "model": "Qwen3.6-27B-Q4_K_M.gguf",
        "messages": [
            {"role": "system", "content": SYSTEM_PROMPT},
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": cal},
                    {"type": "text", "text": comm},
                    {"type": "text", "text": aw},
                ],
            },
        ],
    }
    resp = requests.post(endpoint, json=payload, timeout=120)
    resp.raise_for_status()
    return resp.json()["choices"][0]["message"]["content"]
