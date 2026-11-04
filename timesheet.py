"""Read a WhenIWork timesheet export and apply clock-in / clock-out rules."""

from __future__ import annotations

import csv
import io
import zipfile
from collections import Counter
from dataclasses import dataclass
from datetime import date, datetime, time, timedelta
from xml.etree import ElementTree as ET

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}
EXCEL_EPOCH = datetime(1899, 12, 30)

HEADER_ALIASES = {
    "first name": "first",
    "last name": "last",
    "employee id": "employee_id",
    "date": "date",
    "start time": "start",
    "end time": "end",
    "position": "position",
    "schedule": "schedule",
    "clock out note": "out_note",
}


class TimesheetError(Exception):
    """The file is missing or not a timesheet we can read."""


@dataclass(frozen=True)
class Rule:
    clock_in_before: time | None
    clock_out_after: time | None
    minimum_days: int
    schedule: str = ""
    exclude_note: str = ""
    period: str = "all"

    def sentence(self) -> str:
        parts: list[str] = []
        if self.clock_in_before is not None:
            parts.append(f"clocked in before {format_time(self.clock_in_before)}")
        if self.clock_out_after is not None:
            parts.append(f"clocked out after {format_time(self.clock_out_after)}")
        if not parts:
            what = "had a clock entry"
        elif len(parts) == 1:
            what = parts[0]
        else:
            what = f"{parts[0]} and {parts[1]}"
        day_word = "day" if self.minimum_days == 1 else "days"
        where = f" on the {self.schedule} schedule" if self.schedule else ""
        sentence = (
            f"People who {what}{where} "
            f"on at least {self.minimum_days} {day_word} in this report."
        )
        if self.exclude_note:
            sentence += f' Anyone whose clock out note contains "{self.exclude_note}" is excluded.'
        return sentence


def format_time(value: time) -> str:
    hour = value.hour % 12 or 12
    suffix = "AM" if value.hour < 12 else "PM"
    return f"{hour}:{value.minute:02d} {suffix}"


def format_date(value: datetime) -> str:
    return f"{value.strftime('%a')}, {value.strftime('%b')} {value.day}"


def parse_clock(value: str | None) -> time | None:
    """Parse an HTML time input. Blank means that half of the rule is off."""
    if value is None or not str(value).strip():
        return None
    text = str(value).strip()
    try:
        parsed = datetime.strptime(text, "%H:%M")
    except ValueError as error:
        raise TimesheetError("Use a clock time like 7:45 AM.") from error
    return parsed.time().replace(second=0, microsecond=0)


def read_timesheet(data: bytes, filename: str) -> list[dict]:
    name = (filename or "").lower()
    if name.endswith(".csv"):
        return _punches_from_rows(_rows_from_csv(data))
    if name.endswith(".xlsx"):
        return _punches_from_rows(_rows_from_xlsx(data))
    raise TimesheetError("Upload the WhenIWork timesheet as an Excel file (.xlsx).")


def period_bounds(period: str, today: date) -> tuple[date, date] | None:
    """Inclusive dates for a period. Weeks are Monday through Friday."""
    if period == "all":
        return None
    if period == "this-month":
        start = today.replace(day=1)
        return start, _month_end(start)
    if period == "last-month":
        end = today.replace(day=1) - timedelta(days=1)
        return end.replace(day=1), end
    if period in {"this-week", "last-week"}:
        monday = today - timedelta(days=today.weekday())
        if period == "last-week":
            monday -= timedelta(days=7)
        return monday, monday + timedelta(days=4)
    raise TimesheetError("Choose a period from the list.")


def _month_end(first: date) -> date:
    if first.month == 12:
        following = first.replace(year=first.year + 1, month=1)
    else:
        following = first.replace(month=first.month + 1)
    return following - timedelta(days=1)


def evaluate(punches: list[dict], rule: Rule, today: date | None = None) -> dict:
    if rule.minimum_days < 1:
        raise TimesheetError("The number of days needs to be at least 1.")
    if rule.clock_in_before is None and rule.clock_out_after is None:
        raise TimesheetError("Set a clock-in time, a clock-out time, or both.")

    selected = punches
    if rule.schedule:
        selected = [punch for punch in punches if punch["schedule"] == rule.schedule]
    window = period_bounds(rule.period, today or date.today())
    if window is not None:
        start, end = window
        selected = [punch for punch in selected if start <= punch["start"].date() <= end]
        if not selected:
            raise TimesheetError("No clock entries fall in that period.")
    excluded_ids = _excluded_people(selected, rule.exclude_note)
    note_needle = " ".join(rule.exclude_note.split()).casefold()

    grouped: dict[tuple[str, datetime.date], dict] = {}
    for punch in selected:
        key = (punch["person_id"], punch["start"].date())
        day = grouped.get(key)
        if day is None:
            day = {
                "person_id": punch["person_id"],
                "employee_id": punch["employee_id"],
                "first": punch["first"],
                "last": punch["last"],
                "date": punch["start"].date(),
                "clock_in": punch["start"],
                "clock_out": punch["end"],
                "punches": 0,
                "positions": [],
                "out_notes": [],
                "excluded_by_note": False,
            }
            grouped[key] = day
        day["punches"] += 1
        if punch["position"]:
            day["positions"].append(punch["position"])
        note = punch.get("out_note", "")
        if note and note not in day["out_notes"]:
            day["out_notes"].append(note)
        if note_needle and note_needle in note.casefold():
            day["excluded_by_note"] = True
        if punch["start"] < day["clock_in"]:
            day["clock_in"] = punch["start"]
        if punch["end"] is not None and (
            day["clock_out"] is None or punch["end"] > day["clock_out"]
        ):
            day["clock_out"] = punch["end"]

    people: dict[str, dict] = {}
    for day in grouped.values():
        person = people.get(day["person_id"])
        if person is None:
            person = {
                "person_id": day["person_id"],
                "employee_id": day["employee_id"],
                "first": day["first"],
                "last": day["last"],
                "name": f"{day['first']} {day['last']}".strip(),
                "positions": [],
                "days_worked": 0,
                "worked": [],
                "matching": [],
            }
            people[day["person_id"]] = person
        person["days_worked"] += 1
        person["worked"].append(day)
        person["positions"].extend(day["positions"])
        day["met_rule"] = _day_matches(day, rule)
        if day["met_rule"]:
            person["matching"].append(day)

    schedules = sorted({punch["schedule"] for punch in punches if punch["schedule"]})
    dates = [punch["start"].date() for punch in selected]
    matched = []
    others = []
    excluded = []
    for person in people.values():
        if person["person_id"] in excluded_ids:
            excluded.append(_person_record(person, include_all=True))
            continue
        record = _person_record(person, include_all=len(person["matching"]) < rule.minimum_days)
        if len(person["matching"]) >= rule.minimum_days:
            matched.append(record)
        else:
            others.append(record)
    excluded.sort(key=lambda item: (-item["matchingDays"], item["last"], item["first"]))
    matched.sort(key=lambda item: (-item["matchingDays"], item["last"], item["first"]))
    others.sort(key=lambda item: (-item["matchingDays"], item["last"], item["first"]))
    no_match = [item for item in others if item["matchingDays"] == 0]
    short = [item for item in others if item["matchingDays"] > 0]

    return {
        "sentence": rule.sentence(),
        "range": {
            "start": min(dates).isoformat() if dates else "",
            "end": max(dates).isoformat() if dates else "",
            "label": _range_label(dates),
        },
        "punchCount": len(selected),
        "dayCount": len(grouped),
        "peopleCount": len({punch["person_id"] for punch in selected}),
        "schedules": schedules,
        "matched": matched,
        "short": short,
        "noMatch": no_match,
        "excludeNote": rule.exclude_note,
        "excluded": excluded,
    }


def _excluded_people(punches: list[dict], note: str) -> set[str]:
    needle = " ".join(note.split()).casefold()
    if not needle:
        return set()
    if not punches or not punches[0].get("has_out_note_column"):
        raise TimesheetError("This timesheet has no Clock Out Note column.")
    excluded = set()
    for punch in punches:
        if needle in punch["out_note"].casefold():
            excluded.add(punch["person_id"])
    return excluded


def _day_matches(day: dict, rule: Rule) -> bool:
    if rule.clock_in_before is not None:
        threshold = datetime.combine(day["date"], rule.clock_in_before)
        if day["clock_in"] >= threshold:
            return False
    if rule.clock_out_after is not None:
        if day["clock_out"] is None:
            return False
        threshold = datetime.combine(day["date"], rule.clock_out_after)
        if day["clock_out"] <= threshold:
            return False
    return True


def _person_record(person: dict, include_all: bool = False) -> dict:
    shown = person["worked"] if include_all else person["matching"]
    days = sorted(shown, key=lambda day: day["date"])
    position = ""
    if person["positions"]:
        position = Counter(person["positions"]).most_common(1)[0][0]
    return {
        "name": person["name"],
        "first": person["first"],
        "last": person["last"],
        "employeeId": person["employee_id"],
        "position": position,
        "daysWorked": person["days_worked"],
        "matchingDays": len(person["matching"]),
        "days": [
            {
                "date": day["date"].isoformat(),
                "dateLabel": format_date(day["clock_in"]),
                "clockIn": format_time(day["clock_in"].time()),
                "clockOut": _clock_out_label(day),
                "clockOutNote": " · ".join(day["out_notes"]),
                "entries": day["punches"],
                "metRule": bool(day.get("met_rule")),
                "excludedByNote": bool(day.get("excluded_by_note")),
            }
            for day in days
        ],
    }


def _clock_out_label(day: dict) -> str:
    if day["clock_out"] is None:
        return "Still clocked in"
    label = format_time(day["clock_out"].time())
    if day["clock_out"].date() != day["date"]:
        label += f" ({format_date(day['clock_out'])})"
    return label


def _range_label(dates: list) -> str:
    if not dates:
        return "No dated clock entries"
    start, end = min(dates), max(dates)
    if start == end:
        return format_date(datetime.combine(start, time()))
    if start.year == end.year and start.month == end.month:
        return f"{start.strftime('%b')} {start.day} – {end.day}, {start.year}"
    if start.year == end.year:
        return f"{start.strftime('%b')} {start.day} – {end.strftime('%b')} {end.day}, {start.year}"
    return f"{format_date(datetime.combine(start, time()))} – {format_date(datetime.combine(end, time()))}"


def _punches_from_rows(rows: list[dict[str, str]]) -> list[dict]:
    if not rows:
        raise TimesheetError("That file has no rows.")

    header_index = None
    columns: dict[str, str] = {}
    for index, row in enumerate(rows):
        mapped = {}
        for letter, value in row.items():
            key = HEADER_ALIASES.get(_normalize_header(value))
            if key:
                mapped[key] = letter
        if {"first", "last", "start", "end"} <= set(mapped):
            header_index = index
            columns = mapped
            break
    if header_index is None:
        raise TimesheetError(
            "This file does not look like a WhenIWork timesheet. "
            "It needs First Name, Last Name, Start Time, and End Time columns."
        )

    punches = []
    for row in rows[header_index + 1 :]:
        first = _cell(row, columns, "first")
        last = _cell(row, columns, "last")
        if not first and not last:
            continue
        start = _cell_datetime(_cell(row, columns, "start"), _cell(row, columns, "date"))
        end = _cell_datetime(_cell(row, columns, "end"), _cell(row, columns, "date"))
        if start is None:
            continue
        employee_id = _clean_id(_cell(row, columns, "employee_id"))
        person_id = employee_id or _normalize_header(f"{first} {last}")
        punches.append(
            {
                "person_id": person_id,
                "employee_id": employee_id,
                "first": first,
                "last": last,
                "start": start,
                "end": end,
                "position": _cell(row, columns, "position"),
                "schedule": _cell(row, columns, "schedule"),
                "out_note": _cell(row, columns, "out_note"),
                "has_out_note_column": "out_note" in columns,
            }
        )
    if not punches:
        raise TimesheetError("The timesheet has column headers, but no clock entries.")
    return punches


def _cell(row: dict[str, str], columns: dict[str, str], key: str) -> str:
    letter = columns.get(key)
    if not letter:
        return ""
    return (row.get(letter) or "").strip()


def _normalize_header(value: str) -> str:
    return " ".join(value.replace("\n", " ").strip().lower().split())


def _clean_id(value: str) -> str:
    if value.endswith(".0") and value[:-2].isdigit():
        return value[:-2]
    return value


def _cell_datetime(value: str, date_hint: str = "") -> datetime | None:
    parsed = _parse_datetime(value)
    if parsed is None:
        return None
    if parsed.year <= 1900:
        hint = _parse_datetime(date_hint)
        if hint is None or hint.year <= 1900:
            return None
        parsed = datetime.combine(hint.date(), parsed.time())
    return parsed


def _parse_datetime(value: str) -> datetime | None:
    text = (value or "").strip()
    if not text:
        return None
    try:
        serial = float(text)
    except ValueError:
        serial = None
    if serial is not None:
        if serial <= 0:
            return None
        return _from_excel(serial)
    for fmt in (
        "%m/%d/%Y %I:%M %p",
        "%m/%d/%Y %I:%M:%S %p",
        "%m/%d/%Y %H:%M",
        "%m/%d/%Y %H:%M:%S",
        "%Y-%m-%d %H:%M:%S",
        "%Y-%m-%d %H:%M",
        "%m/%d/%Y",
        "%Y-%m-%d",
    ):
        try:
            return datetime.strptime(text, fmt)
        except ValueError:
            continue
    return None


def _from_excel(serial: float) -> datetime:
    """Convert an Excel serial to a clock time rounded the way a timesheet displays it."""
    moment = EXCEL_EPOCH + timedelta(days=serial)
    if moment.microsecond >= 500_000:
        moment += timedelta(seconds=1)
    moment = moment.replace(microsecond=0)
    if moment.second >= 30:
        moment += timedelta(minutes=1)
    return moment.replace(second=0)


def _rows_from_csv(data: bytes) -> list[dict[str, str]]:
    text = data.decode("utf-8-sig", errors="replace")
    reader = csv.reader(io.StringIO(text))
    rows = []
    for raw in reader:
        rows.append({_column_letter(index + 1): value for index, value in enumerate(raw)})
    return rows


def _rows_from_xlsx(data: bytes) -> list[dict[str, str]]:
    try:
        archive = zipfile.ZipFile(io.BytesIO(data))
    except zipfile.BadZipFile as error:
        raise TimesheetError("That Excel file could not be opened.") from error

    with archive:
        names = set(archive.namelist())
        strings = _shared_strings(archive, names)
        sheet_names = _worksheet_names(archive, names)
        for sheet_name in sheet_names:
            rows = _sheet_rows(archive, sheet_name, strings)
            if _looks_like_timesheet(rows):
                return rows
        if not sheet_names:
            raise TimesheetError("That Excel file has no sheets.")
        return _sheet_rows(archive, sheet_names[0], strings)


def _looks_like_timesheet(rows: list[dict[str, str]]) -> bool:
    for row in rows[:15]:
        headers = {_normalize_header(value) for value in row.values()}
        if {"first name", "last name", "start time", "end time"} <= headers:
            return True
    return False


def _shared_strings(archive: zipfile.ZipFile, names: set[str]) -> list[str]:
    if "xl/sharedStrings.xml" not in names:
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings = []
    for item in root.findall("m:si", NS):
        texts = [
            node.text or ""
            for node in item.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
        ]
        strings.append("".join(texts))
    return strings


def _worksheet_names(archive: zipfile.ZipFile, names: set[str]) -> list[str]:
    rels_name = "xl/_rels/workbook.xml.rels"
    if "xl/workbook.xml" not in names or rels_name not in names:
        return sorted(name for name in names if name.startswith("xl/worksheets/sheet"))
    rels = ET.fromstring(archive.read(rels_name))
    targets = {}
    for rel in rels:
        rel_id = rel.attrib.get("Id")
        target = rel.attrib.get("Target", "")
        if target.startswith("/"):
            path = target.lstrip("/")
        elif target.startswith("xl/"):
            path = target
        else:
            path = "xl/" + target
        targets[rel_id] = path
    workbook = ET.fromstring(archive.read("xl/workbook.xml"))
    sheets = []
    for sheet in workbook.findall("m:sheets/m:sheet", NS):
        rel_id = sheet.attrib.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id")
        path = targets.get(rel_id)
        if path and path in names:
            sheets.append(path)
    return sheets


def _sheet_rows(archive: zipfile.ZipFile, sheet_name: str, strings: list[str]) -> list[dict[str, str]]:
    root = ET.fromstring(archive.read(sheet_name))
    rows = []
    for row in root.findall("m:sheetData/m:row", NS):
        values: dict[str, str] = {}
        for cell in row.findall("m:c", NS):
            ref = cell.attrib.get("r", "")
            letter = "".join(character for character in ref if character.isalpha())
            values[letter] = _cell_text(cell, strings)
        if any(value.strip() for value in values.values()):
            rows.append(values)
    return rows


def _cell_text(cell: ET.Element, strings: list[str]) -> str:
    kind = cell.attrib.get("t")
    if kind == "inlineStr":
        inline = cell.find("m:is", NS)
        if inline is None:
            return ""
        return "".join(
            node.text or ""
            for node in inline.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t")
        )
    value = cell.find("m:v", NS)
    if value is None or value.text is None:
        return ""
    if kind == "s":
        index = int(value.text)
        return strings[index] if 0 <= index < len(strings) else ""
    return value.text


def _column_letter(number: int) -> str:
    letters = ""
    while number:
        number, remainder = divmod(number - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters
