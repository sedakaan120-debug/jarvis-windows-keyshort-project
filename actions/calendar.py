"""Windows takvim araci - Outlook COM varsa takvim okuma/ekleme/silme yapar."""

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
            "Windows takvim erisimi icin Microsoft Outlook ve pywin32 gerekir. "
            "Kurulum: pip install pywin32"
        ) from exc


def _calendar_folder():
    app = _outlook()
    return app.Session.GetDefaultFolder(9)


def _month_start(value: dt.datetime) -> dt.datetime:
    return value.replace(day=1, hour=0, minute=0, second=0, microsecond=0)


def _add_months(value: dt.datetime, months: int) -> dt.datetime:
    total = (value.year * 12 + (value.month - 1)) + months
    year = total // 12
    month = total % 12 + 1
    return value.replace(year=year, month=month, day=1)


def _normalize_query(query: str) -> dict:
    q = (query or "today").strip().lower()
    now = dt.datetime.now()
    today_start = now.replace(hour=0, minute=0, second=0, microsecond=0)

    if "gelecek ay" in q or "onumuzdeki ay" in q or "next month" in q:
        start = _add_months(_month_start(now), 1)
        end = _add_months(start, 1)
        return {"start": start, "end": end, "kind": "range", "header": "Gelecek ay icin {count} etkinlik buldum:", "empty": "Gelecek ay takviminde etkinlik gorunmuyor.", "default_limit": 24}
    if "bu ay" in q or "this month" in q:
        start = _month_start(now)
        end = _add_months(start, 1)
        return {"start": start, "end": end, "kind": "range", "header": "Bu ay icin {count} etkinlik buldum:", "empty": "Bu ay takviminde etkinlik gorunmuyor.", "default_limit": 24}

    day_match = re.search(r"(\d+)\s*(g[uü]n|gun|day|days)", q)
    if day_match:
        days = max(1, min(365, int(day_match.group(1))))
        return {"start": today_start, "end": today_start + dt.timedelta(days=days), "kind": "range", "header": f"Onumuzdeki {days} gun icin {{count}} etkinlik buldum:", "empty": f"Onumuzdeki {days} gunde takviminde etkinlik gorunmuyor.", "default_limit": min(60, max(8, days * 2))}

    if any(token in q for token in ("yarin", "tomorrow")):
        start = today_start + dt.timedelta(days=1)
        return {"start": start, "end": start + dt.timedelta(days=1), "kind": "range", "header": "Yarin icin {count} etkinlik buldum:", "empty": "Yarin takviminde etkinlik gorunmuyor.", "default_limit": 6}
    if any(token in q for token in ("hafta", "week", "7 gun")):
        return {"start": today_start, "end": today_start + dt.timedelta(days=7), "kind": "range", "header": "Onumuzdeki 7 gun icin {count} etkinlik buldum:", "empty": "Onumuzdeki 7 gunde takviminde etkinlik gorunmuyor.", "default_limit": 10}
    if any(token in q for token in ("siradaki", "sıradaki", "sonraki", "next")):
        return {"start": now, "end": now + dt.timedelta(days=365), "kind": "next", "header": "", "empty": "Siradaki takvim etkinligini bulamadim.", "default_limit": 1}
    if any(token in q for token in ("ajanda", "agenda", "yaklasan", "yaklaşan", "upcoming")):
        return {"start": now, "end": now + dt.timedelta(days=30), "kind": "agenda", "header": "Yaklasan ajandanda {count} etkinlik var:", "empty": "Yaklasan takvim etkinligi gorunmuyor.", "default_limit": 8}

    return {"start": today_start, "end": today_start + dt.timedelta(days=1), "kind": "range", "header": "Bugun icin {count} etkinlik buldum:", "empty": "Bugun takviminde etkinlik gorunmuyor.", "default_limit": 6}


def _restrict(items, start: dt.datetime, end: dt.datetime):
    items.IncludeRecurrences = True
    items.Sort("[Start]")
    start_s = start.strftime("%m/%d/%Y %I:%M %p")
    end_s = end.strftime("%m/%d/%Y %I:%M %p")
    return items.Restrict(f"[Start] < '{end_s}' AND [End] > '{start_s}'")


def _event_from_item(item) -> dict:
    start = item.Start
    end = item.End
    if not isinstance(start, dt.datetime):
        start = dt.datetime.fromtimestamp(start.timestamp())
    if not isinstance(end, dt.datetime):
        end = dt.datetime.fromtimestamp(end.timestamp())
    return {
        "start": start.replace(tzinfo=None),
        "end": end.replace(tzinfo=None),
        "title": str(item.Subject or "Adsiz etkinlik"),
        "location": str(item.Location or ""),
        "all_day": bool(item.AllDayEvent),
    }


def _day_label(when: dt.datetime, now: dt.datetime) -> str:
    if when.date() == now.date():
        return "bugun"
    if when.date() == now.date() + dt.timedelta(days=1):
        return "yarin"
    return f"{when.day} {TR_MONTHS[when.month]} {TR_WEEKDAYS[when.weekday()]}"


def _format_event_line(event: dict, now: dt.datetime) -> str:
    prefix = _day_label(event["start"], now)
    if event["all_day"]:
        time_part = f"{prefix} tum gun"
    else:
        time_part = f"{prefix} {event['start'].strftime('%H:%M')}-{event['end'].strftime('%H:%M')}"
    pieces = [f"{time_part} - {event['title']}"]
    if event["location"]:
        pieces.append(f"@ {event['location']}")
    return " ".join(pieces)


def get_calendar_events(query: str = "today", limit: int = 6) -> str:
    window = _normalize_query(query)
    limit = max(1, min(60, int(limit or window["default_limit"])))
    try:
        restricted = _restrict(_calendar_folder().Items, window["start"], window["end"])
        events = []
        for item in restricted:
            try:
                events.append(_event_from_item(item))
            except Exception:
                continue
    except Exception as exc:
        return f"Takvim okunamadi: {exc}"

    events.sort(key=lambda event: (event["start"], event["title"].lower()))
    if not events:
        return window["empty"]

    now = dt.datetime.now()
    if window["kind"] == "next":
        return f"Siradaki etkinlik: {_format_event_line(events[0], now)}."

    selected = events[:limit]
    lines = [str(window["header"]).format(count=len(selected))]
    lines.extend(f"- {_format_event_line(event, now)}" for event in selected)
    return "\n".join(lines)


def _parse_iso(value: str) -> dt.datetime:
    return dt.datetime.fromisoformat(value.replace("Z", "+00:00")).replace(tzinfo=None)


def add_calendar_event(
    title: str,
    start_iso: str,
    end_iso: str = "",
    notes: str = "",
    location: str = "",
    calendar_name: str = "",
    all_day: bool = False,
) -> str:
    title = (title or "").strip()
    start_iso = (start_iso or "").strip()
    if not title:
        return "Takvime eklemek icin etkinlik basligi gerekli."
    if not start_iso:
        return "Takvime eklemek icin baslangic tarihi gerekli."

    try:
        app = _outlook()
        item = app.CreateItem(1)
        start = _parse_iso(start_iso)
        end = _parse_iso(end_iso) if end_iso else start + dt.timedelta(hours=1)
        item.Subject = title
        item.Start = start
        item.End = end
        item.Body = notes or ""
        item.Location = location or ""
        item.AllDayEvent = bool(all_day)
        item.Save()
        return f"Takvime eklendi: {_format_event_line({'start': start, 'end': end, 'title': title, 'location': location, 'all_day': bool(all_day)}, dt.datetime.now())}."
    except Exception as exc:
        return f"Takvim etkinligi eklenemedi: {exc}"


def delete_calendar_event(
    title: str,
    start_iso: str = "",
    calendar_name: str = "",
    delete_all_matches: bool = False,
) -> str:
    title = (title or "").strip()
    if not title:
        return "Takvimden silmek icin etkinlik basligi gerekli."

    try:
        start = _parse_iso(start_iso) if start_iso else dt.datetime.now() - dt.timedelta(days=30)
        end = start + dt.timedelta(days=1) if start_iso else dt.datetime.now() + dt.timedelta(days=365)
        restricted = _restrict(_calendar_folder().Items, start, end)
        matches = []
        needle = title.casefold()
        for item in restricted:
            subject = str(getattr(item, "Subject", "") or "")
            if needle in subject.casefold():
                matches.append(item)
        if not matches:
            return "Silinecek etkinlik bulunamadi."
        if len(matches) > 1 and not delete_all_matches:
            previews = [_format_event_line(_event_from_item(item), dt.datetime.now()) for item in matches[:3]]
            return "Birden fazla etkinlik eslesti. Eslesenler: " + " | ".join(previews)
        deleted_preview = _format_event_line(_event_from_item(matches[0]), dt.datetime.now())
        for item in matches if delete_all_matches else matches[:1]:
            item.Delete()
        return f"Takvimden silindi: {deleted_preview}."
    except Exception as exc:
        return f"Takvim etkinligi silinemedi: {exc}"
