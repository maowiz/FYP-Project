import os
import sys
import datetime


_LEVELS = {"DEBUG": 10, "INFO": 20, "WARN": 30, "ERROR": 40}
_CURRENT_LEVEL = _LEVELS.get(os.environ.get("ASSISTANT_LOG_LEVEL", "INFO").upper(), 20)


def set_level(level_name: str) -> None:
    global _CURRENT_LEVEL
    _CURRENT_LEVEL = _LEVELS.get(level_name.upper(), 20)


def _should_log(level: int) -> bool:
    return level >= _CURRENT_LEVEL


def _fmt(level_name: str, message: str) -> str:
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    return f"[{ts}] {level_name}: {message}"


def debug(message: str) -> None:
    if _should_log(_LEVELS["DEBUG"]):
        print(_fmt("DEBUG", message))


def info(message: str) -> None:
    if _should_log(_LEVELS["INFO"]):
        print(_fmt("INFO", message))


def warn(message: str) -> None:
    if _should_log(_LEVELS["WARN"]):
        print(_fmt("WARN", message))


def error(message: str) -> None:
    if _should_log(_LEVELS["ERROR"]):
        print(_fmt("ERROR", message), file=sys.stderr)


