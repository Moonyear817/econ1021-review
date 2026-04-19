import json
from datetime import datetime
from pathlib import Path
from typing import Any


def now_iso() -> str:
    return datetime.now().isoformat(timespec="seconds")


def default_progress() -> dict[str, Any]:
    return {
        "attempts": [],
        "stats": {},
        "wrong_book": {},
    }


def load_progress(path: Path) -> dict[str, Any]:
    if not path.exists():
        return default_progress()
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return default_progress()


def save_progress(path: Path, data: dict[str, Any]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")


def record_attempt(
    progress: dict[str, Any],
    *,
    qid: str,
    selected: str,
    correct_answer: str,
    stem: str,
    options: dict[str, str],
    explanation_zh: str,
    explanation_en: str,
    lecture_title: str,
    mode: str,
) -> dict[str, Any]:
    is_correct = selected == correct_answer

    attempt = {
        "time": now_iso(),
        "qid": qid,
        "selected": selected,
        "correct_answer": correct_answer,
        "is_correct": is_correct,
        "lecture_title": lecture_title,
        "mode": mode,
    }
    progress.setdefault("attempts", []).append(attempt)

    stats = progress.setdefault("stats", {}).setdefault(qid, {"correct": 0, "wrong": 0})
    if is_correct:
        stats["correct"] += 1
    else:
        stats["wrong"] += 1

    wrong_book = progress.setdefault("wrong_book", {})
    if is_correct:
        if qid in wrong_book:
            del wrong_book[qid]
    else:
        wrong_book[qid] = {
            "qid": qid,
            "last_wrong_time": now_iso(),
            "stem": stem,
            "options": options,
            "correct_answer": correct_answer,
            "your_answer": selected,
            "explanation_zh": explanation_zh,
            "explanation_en": explanation_en,
            "lecture_title": lecture_title,
        }

    return progress
