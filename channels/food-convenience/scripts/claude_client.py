"""claude -p ヘッドレスモード呼び出しの共通ヘルパー（generate-themes.py / generate-script.py 共用）。"""

import json
import logging
import os
import re
import subprocess
import sys
from pathlib import Path

# タスクスケジューラ実行時はPATHが通っていないことがあるため、
# npmのグローバルインストール先をフルパスで既定値にする（%APPDATA%\npm\claude.cmd）。
_appdata = os.environ.get("APPDATA")
DEFAULT_CLAUDE_CMD = str(Path(_appdata) / "npm" / "claude.cmd") if _appdata else "claude"


def setup_logging(log_file: Path, logger_name: str) -> logging.Logger:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:
            pass

    log_file.parent.mkdir(parents=True, exist_ok=True)
    logger = logging.getLogger(logger_name)
    logger.setLevel(logging.INFO)
    formatter = logging.Formatter("%(asctime)s [%(levelname)s] %(message)s")

    file_handler = logging.FileHandler(log_file, encoding="utf-8")
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)

    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)
    logger.addHandler(stream_handler)
    return logger


def decode_safely(data: bytes | None) -> str:
    if not data:
        return ""
    for encoding in ("utf-8", "cp932"):
        try:
            return data.decode(encoding)
        except UnicodeDecodeError:
            continue
    return data.decode("utf-8", errors="replace")


def call_claude(prompt: str, timeout: int, claude_cmd: str = DEFAULT_CLAUDE_CMD) -> str:
    command = f'"{claude_cmd}" -p --output-format json'
    try:
        proc = subprocess.run(
            command,
            input=prompt.encode("utf-8"),
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            shell=True,
            timeout=timeout,
        )
    except subprocess.TimeoutExpired as exc:
        raise RuntimeError(f"claudeの応答がタイムアウトしました（{timeout}秒）。") from exc

    stdout_text = decode_safely(proc.stdout)
    stderr_text = decode_safely(proc.stderr)

    if proc.returncode != 0:
        raise RuntimeError(
            f"claudeがエラー終了しました（code={proc.returncode}, claude_cmd={claude_cmd}）。"
            f"stderr:\n{stderr_text[:2000]}\nstdout:\n{stdout_text[:2000]}"
        )

    try:
        outer = json.loads(stdout_text)
    except json.JSONDecodeError as exc:
        raise RuntimeError(
            f"claudeの出力（JSON）を解析できませんでした。stdout:\n{stdout_text[:2000]}"
        ) from exc

    if outer.get("is_error"):
        raise RuntimeError(f"claudeがエラーを返しました: {outer}")

    result_text = outer.get("result", "")
    if not result_text:
        raise RuntimeError(f"claudeの出力にresultフィールドがありませんでした: {outer}")

    return result_text


def strip_code_fence(text: str) -> str:
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```[a-zA-Z]*\n?", "", text)
        text = re.sub(r"\n?```$", "", text)
    return text.strip()
