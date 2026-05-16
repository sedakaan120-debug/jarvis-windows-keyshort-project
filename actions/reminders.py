"""Windows reminders/tasks - Outlook Tasks COM entegrasyonu."""

from __future__ import annotations

import datetime as dt
import re


TR_WEEKDAYS = ["Pazartesi", "Sali", "Carsamba", "Persembe", "Cuma", "Cumartesi", "Pazar"]
TR_MONTHS = ["", "Ocak", "Subat", "Mart", "Nisan", "Mayis", "Haziran", "Temmuz", "Agustos", "Eylul", "Ekim", "Kasim", "Aralik"]


def _outlook():
    try:
        import win32com.client

        return win32com.client.Dispatch("Outlook.Application")
    except Exception as exc:
        raise RuntimeError(
            "Windows animsaticilari icin Microsoft Outlook ve pywin32 gerekir. "
            "Kurulum: pip install pywin32"
        ) from exc


def _tasks_folder():
    return _outlook().Session.GetDefaultFolder(13)


def _normalize_query(query: str) -> tuple[str, int]:
    q = (query or "").strip().lower()
    if any(token in q for token in ("bugun", "today")):
        return "today", 8
    if any(token in q for token in ("geciken", "gecmis", "overdue")):
        return "overdue", 8
    if any(token in q for token in ("siradaki", "sıradaki", "next")):
        return "next", 1
    if any(token in q for token in ("hepsi", "tum", "tüm", "all", "listele")):
        return "all", 10
    return "upcoming", 8


def _normalize_due_iso(due_iso: str) -> tuple[dt.datetime | None, bool]:
    raw = (due_iso or "").strip()
    if not raw:
        return None, False

    candidates = (
        ("%Y-%m-%dT%H:%M:%S", False),
        ("%Y-%m-%dT%H:%M", False),
        ("%Y-%m-%d %H:%M:%S", False),
        ("%Y-%m-%d %H:%M", False),
        ("%d.%m.%Y %H:%M", False),
        ("%Y-%m-%d", True),
        ("%d.%m.%Y", True),
    )

    if raw.endswith("Z"):
        parsed = dt.datetime.fromisoformat(raw.replace("Z", "+00:00")).replace(tzinfo=None)
        return parsed, False
    if re.match(r"^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}", raw):
        return dt.datetime.fromisoformat(raw).replace(tzinfo=None), False

    for fmt, is_all_day in candidates:
        try:
            return dt.datetime.strptime(raw, fmt), is_all_day
        except ValueError:
            continue

    raise ValueError("Animsatici tarihi gecersiz. due_iso icin 'YYYY-MM-DD' veya 'YYYY-MM-DDTHH:MM' kullan.")


def _task_due(item) -> dt.datetime | None:
    due = getattr(item, "DueDate", None)
    if not due:
        return None
    try:
        return due.replace(tzinfo=None)
    except Exception:
        try:
            return dt.datetime.fromtimestamp(due.timestamp()).replace(tzinfo=None)
        except Exception:
            return None


def _day_label(when: dt.datetime, now: dt.datetime) -> str:
    if when.date() == now.date():
        return "bugun"
    if when.date() == now.date() + dt.timedelta(days=1):
        return "yarin"
    return f"{when.day} {TR_MONTHS[when.month]} {TR_WEEKDAYS[when.weekday()]}"


def _format_task(item, now: dt.datetime) -> str:
    due = _task_due(item)
    when = f"{_day_label(due, now)} {due.strftime('%H:%M')}" if due else "zaman atanmamis"
    title = str(getattr(item, "Subject", "") or "Adsiz animsatici")
    return f"{when} - {title}"


def _load_open_tasks() -> list:
    items = []
    for item in _tasks_folder().Items:
        try:
            if not bool(getattr(item, "Complete", False)):
                items.append(item)
        except Exception:
            continue
    items.sort(key=lambda item: (_task_due(item) is None, _task_due(item) or dt.datetime.max, str(getattr(item, "Subject", "")).lower()))
    return items


def get_reminders(query: str = "upcoming", limit: int = 8, list_name: str = "") -> str:
    mode, default_limit = _normalize_query(query)
    limit = max(1, min(20, int(limit or default_limit)))

    try:
        now = dt.datetime.now()
        today = now.date()
        tasks = _load_open_tasks()
    except Exception as exc:
        return f"Animsaticilar okunamadi: {exc}"

    if mode == "today":
        tasks = [task for task in tasks if (_task_due(task) and _task_due(task).date() == today)]
        empty = "Bugun icin animsatici gorunmuyor."
        header = f"Bugun icin {{count}} animsatici buldum:"
    elif mode == "overdue":
        tasks = [task for task in tasks if (_task_due(task) and _task_due(task).date() < today)]
        empty = "Geciken animsatici gorunmuyor."
        header = f"Gecikmis {{count}} animsatici buldum:"
    elif mode == "next":
        tasks = tasks[:1]
        empty = "Siradaki animsaticiyi bulamadim."
        header = ""
    elif mode == "all":
        empty = "Kayitli acik animsatici gorunmuyor."
        header = f"Acik {{count}} animsatici buldum:"
    else:
        tasks = [task for task in tasks if not _task_due(task) or _task_due(task).date() >= today]
        empty = "Yaklasan animsatici gorunmuyor."
        header = f"Yaklasan {{count}} animsatici buldum:"

    if not tasks:
        return empty
    if mode == "next":
        return f"Siradaki animsatici: {_format_task(tasks[0], now)}."

    selected = tasks[:limit]
    lines = [header.format(count=len(selected))]
    lines.extend(f"- {_format_task(task, now)}" for task in selected)
    return "\n".join(lines)


def add_reminder(
    title: str,
    due_iso: str = "",
    notes: str = "",
    list_name: str = "",
    priority: str = "",
    all_day: bool = False,
) -> str:
    if not title or not title.strip():
        return "Animsatici basligi bos olamaz."

    try:
        due, inferred_all_day = _normalize_due_iso(due_iso) if due_iso else (None, False)
        item = _outlook().CreateItem(3)
        item.Subject = title.strip()
        item.Body = (notes or "").strip()
        if due:
            item.DueDate = due.date() if (all_day or inferred_all_day) else due
            item.ReminderSet = True
            item.ReminderTime = due
        if str(priority or "").lower() in {"high", "yuksek", "yüksek", "1"}:
            item.Importance = 2
        item.Save()
        when = due.strftime("%d.%m.%Y %H:%M") if due else "zaman atanmamis"
        return f"Animsatici eklendi: {when} - {title.strip()}"
    except Exception as exc:
        return f"Animsatici eklenemedi: {exc}"
