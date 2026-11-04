# Timesheet Filter

A small page for a WhenIWork timesheet export. Upload the Excel file, set a clock-in and clock-out rule, and see who met it. The check runs in the browser. The timesheet is not sent anywhere else.

The example rule is: clocked in before 7:45 AM and clocked out after 3:45 PM on at least 10 days in the report.

## Run

```bash
python3 app.py
```

Open http://127.0.0.1:8765

The timesheet stays on your computer. Nothing is uploaded to another service.

## How a day is counted

Each row in the WhenIWork export is one clock entry. If someone clocks out for lunch and clocks back in, that day is two rows. Clock check joins those rows: the earliest start is the clock-in, and the latest end is the clock-out.

“Before 7:45 AM” does not include 7:45. “After 3:45 PM” does not include 3:45. Times are read to the nearest minute, which is how the timesheet shows them.

## The export

Use the WhenIWork Timesheets Excel export. The sheet needs these columns:

- First Name
- Last Name
- Start Time
- End Time

Employee ID, Position, and Schedule are used when they are present. A `.csv` export with the same column names also works.

## Tests

```bash
python3 test_timesheet.py
```
