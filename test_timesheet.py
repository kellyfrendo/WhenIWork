import io
import unittest
import zipfile
from datetime import date, datetime
from xml.sax.saxutils import escape

from timesheet import Rule, TimesheetError, evaluate, period_bounds, read_timesheet


def excel_serial(moment: datetime) -> str:
    delta = moment - datetime(1899, 12, 30)
    return str(delta.total_seconds() / 86400)


def workbook(rows: list[list[str]]) -> bytes:
    shared: list[str] = []
    index: dict[str, int] = {}

    def string_id(value: str) -> int:
        if value not in index:
            index[value] = len(shared)
            shared.append(value)
        return index[value]

    body = []
    for row_number, row in enumerate(rows, start=1):
        cells = []
        for column, value in enumerate(row, start=1):
            letter = chr(64 + column)
            ref = f"{letter}{row_number}"
            if row_number == 1 or column <= 3 or column >= 7:
                cells.append(f'<c r="{ref}" t="s"><v>{string_id(value)}</v></c>')
            elif value:
                cells.append(f'<c r="{ref}"><v>{value}</v></c>')
        body.append(f'<row r="{row_number}">{"".join(cells)}</row>')

    shared_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        + "".join(f"<si><t>{escape(value)}</t></si>" for value in shared)
        + "</sst>"
    )
    sheet_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
        f"<sheetData>{''.join(body)}</sheetData></worksheet>"
    )
    workbook_xml = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="Windsor" sheetId="1" r:id="rId1"/></sheets></workbook>'
    )
    rels = (
        '<?xml version="1.0" encoding="UTF-8"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="worksheet" Target="worksheets/sheet1.xml"/>'
        "</Relationships>"
    )
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w") as archive:
        archive.writestr("xl/sharedStrings.xml", shared_xml)
        archive.writestr("xl/workbook.xml", workbook_xml)
        archive.writestr("xl/_rels/workbook.xml.rels", rels)
        archive.writestr("xl/worksheets/sheet1.xml", sheet_xml)
    return output.getvalue()


class TimesheetTests(unittest.TestCase):
    def test_lunch_split_counts_as_one_full_day(self):
        data = workbook(
            [
                ["First Name", "Last Name", "Employee ID", "Date", "Start Time", "End Time", "Position", "Schedule"],
                ["Ada", "Lovelace", "10", "", excel_serial(datetime(2026, 10, 1, 7, 30)), excel_serial(datetime(2026, 10, 1, 12, 0)), "Teacher", "Windsor"],
                ["Ada", "Lovelace", "10", "", excel_serial(datetime(2026, 10, 1, 12, 30)), excel_serial(datetime(2026, 10, 1, 16, 0)), "Teacher", "Windsor"],
                ["Ada", "Lovelace", "10", "", excel_serial(datetime(2026, 10, 2, 7, 44)), excel_serial(datetime(2026, 10, 2, 15, 46)), "Teacher", "Windsor"],
                ["Grace", "Hopper", "11", "", excel_serial(datetime(2026, 10, 1, 8, 0)), excel_serial(datetime(2026, 10, 1, 17, 0)), "Teacher", "Windsor"],
                ["Grace", "Hopper", "11", "", excel_serial(datetime(2026, 10, 2, 7, 45)), excel_serial(datetime(2026, 10, 2, 15, 45)), "Teacher", "Windsor"],
                ["Edith", "Clarke", "12", "", excel_serial(datetime(2026, 10, 3, 7, 10)), excel_serial(datetime(2026, 10, 3, 16, 5)), "Teacher", "Windsor"],
                ["Edith", "Clarke", "12", "", excel_serial(datetime(2026, 10, 4, 8, 0)), excel_serial(datetime(2026, 10, 4, 16, 0)), "Teacher", "Windsor"],
            ]
        )
        punches = read_timesheet(data, "timesheets.xlsx")
        result = evaluate(
            punches,
            Rule(clock_in_before=datetime.strptime("07:45", "%H:%M").time(), clock_out_after=datetime.strptime("15:45", "%H:%M").time(), minimum_days=2),
        )
        self.assertEqual([person["name"] for person in result["matched"]], ["Ada Lovelace"])
        self.assertEqual(result["matched"][0]["matchingDays"], 2)
        self.assertEqual(result["matched"][0]["days"][0]["clockIn"], "7:30 AM")
        self.assertEqual(result["matched"][0]["days"][0]["clockOut"], "4:00 PM")
        self.assertEqual(result["matched"][0]["days"][0]["entries"], 2)
        self.assertEqual(result["short"][0]["name"], "Edith Clarke")
        self.assertEqual(result["short"][0]["matchingDays"], 1)
        self.assertEqual([person["name"] for person in result["noMatch"]], ["Grace Hopper"])
        self.assertEqual(len(result["noMatch"][0]["days"]), 2)
        self.assertEqual([day["clockIn"] for day in result["short"][0]["days"]], ["7:10 AM", "8:00 AM"])
        self.assertEqual([day["metRule"] for day in result["short"][0]["days"]], [True, False])

    def test_clock_out_note_excludes_the_person(self):
        early = excel_serial(datetime(2026, 10, 1, 7, 30))
        late = excel_serial(datetime(2026, 10, 1, 16, 0))
        data = workbook(
            [
                ["First Name", "Last Name", "Employee ID", "Date", "Start Time", "End Time", "Position", "Schedule", "Clock Out Note"],
                ["Ada", "Lovelace", "10", "", early, late, "Teacher", "Windsor", "Out for lunch"],
                ["Ada", "Lovelace", "10", "", excel_serial(datetime(2026, 10, 2, 7, 30)), excel_serial(datetime(2026, 10, 2, 16, 0)), "Teacher", "Windsor", ""],
                ["Grace", "Hopper", "11", "", early, late, "Teacher", "Windsor", ""],
                ["Edith", "Clarke", "12", "", early, late, "Teacher", "Windsor", "Doctor visit"],
            ]
        )
        punches = read_timesheet(data, "timesheets.xlsx")
        rule = Rule(
            clock_in_before=datetime.strptime("07:45", "%H:%M").time(),
            clock_out_after=datetime.strptime("15:45", "%H:%M").time(),
            minimum_days=1,
            exclude_note="LUNCH",
        )
        result = evaluate(punches, rule)
        self.assertEqual([person["name"] for person in result["matched"]], ["Edith Clarke", "Grace Hopper"])
        self.assertEqual([person["name"] for person in result["excluded"]], ["Ada Lovelace"])
        self.assertEqual(result["excluded"][0]["days"][0]["clockOutNote"], "Out for lunch")
        self.assertEqual(
            [day["excludedByNote"] for day in result["excluded"][0]["days"]],
            [True, False],
        )
        self.assertIn('contains "LUNCH"', result["sentence"])

        kept = evaluate(punches, Rule(rule.clock_in_before, rule.clock_out_after, 1))
        self.assertEqual(len(kept["matched"]), 3)
        self.assertEqual(kept["excluded"], [])
        notes = {person["name"]: person["days"][0]["clockOutNote"] for person in kept["matched"]}
        self.assertEqual(notes["Ada Lovelace"], "Out for lunch")
        self.assertEqual(notes["Edith Clarke"], "Doctor visit")
        self.assertEqual(notes["Grace Hopper"], "")

    def test_manager_note_excludes_the_person(self):
        early = excel_serial(datetime(2026, 10, 1, 7, 30))
        late = excel_serial(datetime(2026, 10, 1, 16, 0))
        data = workbook(
            [
                ["First Name", "Last Name", "Employee ID", "Date", "Start Time", "End Time", "Position", "Schedule", "Clock Out Note", "Manager Note"],
                ["Ada", "Lovelace", "10", "", early, late, "Teacher", "Windsor", "Out for lunch", ""],
                ["Grace", "Hopper", "11", "", early, late, "Teacher", "Windsor", "", "See HR"],
                ["Grace", "Hopper", "11", "", excel_serial(datetime(2026, 10, 2, 7, 30)), excel_serial(datetime(2026, 10, 2, 16, 0)), "Teacher", "Windsor", "", ""],
                ["Edith", "Clarke", "12", "", early, late, "Teacher", "Windsor", "Doctor visit", "On site"],
            ]
        )
        punches = read_timesheet(data, "timesheets.xlsx")
        result = evaluate(
            punches,
            Rule(
                clock_in_before=datetime.strptime("07:45", "%H:%M").time(),
                clock_out_after=datetime.strptime("15:45", "%H:%M").time(),
                minimum_days=1,
                exclude_note="HR",
            ),
        )
        self.assertEqual([person["name"] for person in result["matched"]], ["Edith Clarke", "Ada Lovelace"])
        self.assertEqual([person["name"] for person in result["excluded"]], ["Grace Hopper"])
        self.assertEqual(
            [(day["clockOutNote"], day["excludedByNote"]) for day in result["excluded"][0]["days"]],
            [("See HR", True), ("", False)],
        )
        edith = next(person for person in result["matched"] if person["name"] == "Edith Clarke")
        self.assertEqual(edith["days"][0]["clockOutNote"], "Clock out: Doctor visit · Manager: On site")

        same = workbook(
            [
                ["First Name", "Last Name", "Employee ID", "Date", "Start Time", "End Time", "Position", "Schedule", "Clock Out Note", "Manager Note"],
                ["Ida", "Wells", "14", "", early, late, "Teacher", "Windsor", "HR", "HR"],
            ]
        )
        shown = evaluate(
            read_timesheet(same, "timesheets.xlsx"),
            Rule(datetime.strptime("07:45", "%H:%M").time(), datetime.strptime("15:45", "%H:%M").time(), 1),
        )
        self.assertEqual(shown["matched"][0]["days"][0]["clockOutNote"], "Clock out: HR · Manager: HR")

    def test_period_keeps_only_the_chosen_dates(self):
        today = date(2026, 11, 4)
        self.assertIsNone(period_bounds("all", today))
        self.assertEqual(period_bounds("last-week", today), (date(2026, 10, 26), date(2026, 10, 30)))
        self.assertEqual(period_bounds("this-week", today), (date(2026, 11, 2), date(2026, 11, 6)))
        self.assertEqual(period_bounds("this-month", today), (date(2026, 11, 1), date(2026, 11, 30)))
        self.assertEqual(period_bounds("last-month", today), (date(2026, 10, 1), date(2026, 10, 31)))
        self.assertEqual(period_bounds("last-month", date(2026, 1, 15)), (date(2025, 12, 1), date(2025, 12, 31)))
        self.assertEqual(period_bounds("last-week", date(2026, 11, 2)), (date(2026, 10, 26), date(2026, 10, 30)))
        self.assertEqual(period_bounds("this-week", date(2026, 11, 8)), (date(2026, 11, 2), date(2026, 11, 6)))

        def row(first, day, note=""):
            start = excel_serial(datetime(day.year, day.month, day.day, 7, 30))
            end = excel_serial(datetime(day.year, day.month, day.day, 16, 0))
            employee = "10" if first == "Ada" else "12"
            return [first, "Example", employee, "", start, end, "Teacher", "Windsor", note]

        data = workbook(
            [
                ["First Name", "Last Name", "Employee ID", "Date", "Start Time", "End Time", "Position", "Schedule", "Clock Out Note"],
                row("Ada", date(2026, 10, 26)),
                row("Ada", date(2026, 10, 27)),
                row("Ada", date(2026, 10, 31)),
                row("Ada", date(2026, 11, 2), "Out for lunch"),
                row("Edith", date(2026, 11, 4)),
            ]
        )
        punches = read_timesheet(data, "timesheets.xlsx")
        rule = Rule(
            clock_in_before=datetime.strptime("07:45", "%H:%M").time(),
            clock_out_after=datetime.strptime("15:45", "%H:%M").time(),
            minimum_days=1,
            exclude_note="lunch",
        )

        def names(period):
            result = evaluate(punches, Rule(
                rule.clock_in_before, rule.clock_out_after, 1, "", rule.exclude_note, period
            ), today)
            return [person["name"] for person in result["matched"]]

        self.assertEqual(names("last-week"), ["Ada Example"])
        self.assertEqual(names("this-week"), ["Edith Example"])
        self.assertEqual(names("last-month"), ["Ada Example"])
        self.assertEqual(names("this-month"), ["Edith Example"])
        self.assertEqual(names("all"), ["Edith Example"])

        with self.assertRaises(TimesheetError):
            evaluate(punches, Rule(rule.clock_in_before, rule.clock_out_after, 1, "", "", "next-week"), today)

    def test_rejects_a_file_without_clock_columns(self):
        data = workbook([["Name", "Hours"], ["Ada", "8"]])
        with self.assertRaises(TimesheetError):
            read_timesheet(data, "notes.xlsx")


if __name__ == "__main__":
    unittest.main()
