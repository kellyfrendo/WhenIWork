const SPREADSHEET_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const OFFICE_REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";

const HEADER_ALIASES = {
  "first name": "first",
  "last name": "last",
  "employee id": "employee_id",
  date: "date",
  "start time": "start",
  "end time": "end",
  position: "position",
  schedule: "schedule",
  "clock out note": "out_note",
};

const WEEKDAYS = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

class TimesheetError extends Error {
  constructor(message) {
    super(message);
    this.name = "TimesheetError";
  }
}

function checkTimesheet(buffer, filename, rule) {
  const name = String(filename || "").toLowerCase();
  const minimumDays = Number(rule.minimumDays);
  if (!Number.isInteger(minimumDays) || minimumDays < 1) {
    return Promise.reject(new TimesheetError("Enter a whole number of days."));
  }
  const clockInBefore = parseClock(rule.clockInBefore);
  const clockOutAfter = parseClock(rule.clockOutAfter);
  if (clockInBefore == null && clockOutAfter == null) {
    return Promise.reject(new TimesheetError("Set a clock-in time, a clock-out time, or both."));
  }
  const parsed = {
    clockInBefore,
    clockOutAfter,
    minimumDays,
    schedule: String(rule.schedule || "").trim(),
    excludeNote: String(rule.excludeNote || "").trim(),
    period: String(rule.period || "last-week").trim(),
  };

  const read = name.endsWith(".csv")
    ? Promise.resolve(rowsFromCsv(buffer))
    : name.endsWith(".xlsx")
      ? rowsFromXlsx(buffer)
      : Promise.reject(new TimesheetError("Upload the WhenIWork timesheet as an Excel file (.xlsx)."));

  return read.then((rows) => {
    const punches = punchesFromRows(rows);
    const schedules = [...new Set(punches.map((punch) => punch.schedule).filter(Boolean))].sort();
    if (parsed.schedule && !schedules.includes(parsed.schedule)) {
      throw new TimesheetError("That schedule is not in this timesheet.");
    }
    const result = evaluatePunches(punches, parsed, todayISO());
    result.schedules = schedules;
    return result;
  });
}

function parseClock(value) {
  const text = String(value || "").trim();
  if (!text) return null;
  const match = /^(\d{1,2}):(\d{2})$/.exec(text);
  if (!match) throw new TimesheetError("Use a clock time like 7:45 AM.");
  const hour = Number(match[1]);
  const minute = Number(match[2]);
  if (hour > 23 || minute > 59) throw new TimesheetError("Use a clock time like 7:45 AM.");
  return hour * 60 + minute;
}

function evaluatePunches(punches, rule, today) {
  let selected = punches;
  if (rule.schedule) selected = selected.filter((punch) => punch.schedule === rule.schedule);
  const window = periodBounds(rule.period, today);
  if (window) {
    selected = selected.filter((punch) => punch.start.iso >= window.start && punch.start.iso <= window.end);
    if (!selected.length) throw new TimesheetError("No clock entries fall in that period.");
  }

  const needle = rule.excludeNote.replace(/\s+/g, " ").trim().toLowerCase();
  const excludedIds = excludedPeople(selected, needle);
  const grouped = new Map();
  selected.forEach((punch) => {
    const key = `${punch.personId}|${punch.start.iso}`;
    let day = grouped.get(key);
    if (!day) {
      day = {
        personId: punch.personId,
        employeeId: punch.employeeId,
        first: punch.first,
        last: punch.last,
        date: punch.start.iso,
        clockIn: punch.start,
        clockOut: punch.end,
        punches: 0,
        positions: [],
        outNotes: [],
        excludedByNote: false,
      };
      grouped.set(key, day);
    }
    day.punches += 1;
    if (punch.position) day.positions.push(punch.position);
    if (punch.outNote && !day.outNotes.includes(punch.outNote)) day.outNotes.push(punch.outNote);
    if (needle && punch.outNote.toLowerCase().includes(needle)) day.excludedByNote = true;
    if (compareMoments(punch.start, day.clockIn) < 0) day.clockIn = punch.start;
    if (punch.end && (!day.clockOut || compareMoments(punch.end, day.clockOut) > 0)) day.clockOut = punch.end;
  });

  const people = new Map();
  grouped.forEach((day) => {
    let person = people.get(day.personId);
    if (!person) {
      person = {
        personId: day.personId,
        employeeId: day.employeeId,
        first: day.first,
        last: day.last,
        name: `${day.first} ${day.last}`.trim(),
        positions: [],
        daysWorked: 0,
        worked: [],
        matching: [],
      };
      people.set(day.personId, person);
    }
    person.daysWorked += 1;
    person.worked.push(day);
    person.positions.push(...day.positions);
    day.metRule = dayMatches(day, rule);
    if (day.metRule) person.matching.push(day);
  });

  const matched = [];
  const others = [];
  const excluded = [];
  people.forEach((person) => {
    if (excludedIds.has(person.personId)) {
      excluded.push(personRecord(person, true));
      return;
    }
    const record = personRecord(person, person.matching.length < rule.minimumDays);
    if (person.matching.length >= rule.minimumDays) matched.push(record);
    else others.push(record);
  });
  const byPerson = (a, b) => b.matchingDays - a.matchingDays || a.last.localeCompare(b.last) || a.first.localeCompare(b.first);
  excluded.sort(byPerson);
  matched.sort(byPerson);
  others.sort(byPerson);
  return {
    matched,
    short: others.filter((item) => item.matchingDays > 0),
    noMatch: others.filter((item) => item.matchingDays === 0),
    excluded,
    excludeNote: rule.excludeNote,
    schedules: [],
  };
}

function excludedPeople(punches, needle) {
  if (!needle) return new Set();
  if (!punches.length || !punches[0].hasOutNoteColumn) {
    throw new TimesheetError("This timesheet has no Clock Out Note column.");
  }
  const excluded = new Set();
  punches.forEach((punch) => {
    if (punch.outNote.toLowerCase().includes(needle)) excluded.add(punch.personId);
  });
  return excluded;
}

function dayMatches(day, rule) {
  if (rule.clockInBefore != null && day.clockIn.minutes >= rule.clockInBefore) return false;
  if (rule.clockOutAfter != null) {
    if (!day.clockOut) return false;
    if (day.clockOut.iso < day.date) return false;
    if (day.clockOut.iso === day.date && day.clockOut.minutes <= rule.clockOutAfter) return false;
  }
  return true;
}

function personRecord(person, includeAll) {
  const shown = (includeAll ? person.worked : person.matching).slice().sort((a, b) => a.date.localeCompare(b.date));
  const counts = new Map();
  person.positions.forEach((position) => counts.set(position, (counts.get(position) || 0) + 1));
  let position = "";
  let best = 0;
  counts.forEach((count, name) => {
    if (count > best) {
      best = count;
      position = name;
    }
  });
  return {
    name: person.name,
    first: person.first,
    last: person.last,
    employeeId: person.employeeId,
    position,
    daysWorked: person.daysWorked,
    matchingDays: person.matching.length,
    days: shown.map((day) => ({
      date: day.date,
      dateLabel: formatDate(day.clockIn),
      clockIn: formatTime(day.clockIn.minutes),
      clockOut: clockOutLabel(day),
      clockOutNote: day.outNotes.join(" · "),
      entries: day.punches,
      metRule: Boolean(day.metRule),
      excludedByNote: Boolean(day.excludedByNote),
    })),
  };
}

function clockOutLabel(day) {
  if (!day.clockOut) return "Still clocked in";
  let label = formatTime(day.clockOut.minutes);
  if (day.clockOut.iso !== day.date) label += ` (${formatDate(day.clockOut)})`;
  return label;
}

function formatTime(minutes) {
  let hour = Math.floor(minutes / 60);
  const minute = minutes % 60;
  const suffix = hour >= 12 ? "PM" : "AM";
  hour = hour % 12 || 12;
  return `${hour}:${String(minute).padStart(2, "0")} ${suffix}`;
}

function formatDate(moment) {
  const [year, month, day] = moment.iso.split("-").map(Number);
  const weekday = WEEKDAYS[new Date(year, month - 1, day).getDay()];
  return `${weekday}, ${MONTHS[month - 1]} ${day}`;
}

function periodBounds(period, today) {
  if (period === "all") return null;
  if (period === "this-month") {
    const start = `${today.slice(0, 8)}01`;
    return { start, end: monthEnd(start) };
  }
  if (period === "last-month") {
    const end = addDaysISO(`${today.slice(0, 8)}01`, -1);
    return { start: `${end.slice(0, 8)}01`, end };
  }
  if (period === "this-week" || period === "last-week") {
    let monday = addDaysISO(today, -weekdayMonday(today));
    if (period === "last-week") monday = addDaysISO(monday, -7);
    return { start: monday, end: addDaysISO(monday, 4) };
  }
  throw new TimesheetError("Choose a period from the list.");
}

function monthEnd(first) {
  const [year, month] = first.split("-").map(Number);
  return todayISO(new Date(year, month, 0));
}

function weekdayMonday(iso) {
  const [year, month, day] = iso.split("-").map(Number);
  return (new Date(year, month - 1, day).getDay() + 6) % 7;
}

function addDaysISO(iso, days) {
  const [year, month, day] = iso.split("-").map(Number);
  const date = new Date(year, month - 1, day);
  date.setDate(date.getDate() + days);
  return todayISO(date);
}

function todayISO(date) {
  const value = date || new Date();
  const month = String(value.getMonth() + 1).padStart(2, "0");
  const day = String(value.getDate()).padStart(2, "0");
  return `${value.getFullYear()}-${month}-${day}`;
}

function compareMoments(a, b) {
  if (a.iso !== b.iso) return a.iso < b.iso ? -1 : 1;
  return a.minutes - b.minutes;
}

function punchesFromRows(rows) {
  if (!rows.length) throw new TimesheetError("That file has no rows.");
  let headerIndex = -1;
  let columns = {};
  rows.some((row, index) => {
    const mapped = {};
    Object.entries(row).forEach(([letter, value]) => {
      const key = HEADER_ALIASES[normalizeHeader(value)];
      if (key) mapped[key] = letter;
    });
    if (["first", "last", "start", "end"].every((key) => mapped[key])) {
      headerIndex = index;
      columns = mapped;
      return true;
    }
    return false;
  });
  if (headerIndex < 0) {
    throw new TimesheetError(
      "This file does not look like a WhenIWork timesheet. It needs First Name, Last Name, Start Time, and End Time columns."
    );
  }

  const punches = [];
  rows.slice(headerIndex + 1).forEach((row) => {
    const first = cell(row, columns, "first");
    const last = cell(row, columns, "last");
    if (!first && !last) return;
    const start = cellDatetime(cell(row, columns, "start"), cell(row, columns, "date"));
    const end = cellDatetime(cell(row, columns, "end"), cell(row, columns, "date"));
    if (!start) return;
    const employeeId = cleanId(cell(row, columns, "employee_id"));
    punches.push({
      personId: employeeId || normalizeHeader(`${first} ${last}`),
      employeeId,
      first,
      last,
      start,
      end,
      position: cell(row, columns, "position"),
      schedule: cell(row, columns, "schedule"),
      outNote: cell(row, columns, "out_note"),
      hasOutNoteColumn: Object.prototype.hasOwnProperty.call(columns, "out_note"),
    });
  });
  if (!punches.length) throw new TimesheetError("The timesheet has column headers, but no clock entries.");
  return punches;
}

function cell(row, columns, key) {
  const letter = columns[key];
  if (!letter) return "";
  return String(row[letter] || "").trim();
}

function normalizeHeader(value) {
  return String(value || "").replace(/\n/g, " ").trim().toLowerCase().split(/\s+/).filter(Boolean).join(" ");
}

function cleanId(value) {
  if (value.endsWith(".0") && /^\d+$/.test(value.slice(0, -2))) return value.slice(0, -2);
  return value;
}

function cellDatetime(value, dateHint) {
  const parsed = parseDatetime(value);
  if (!parsed) return null;
  if (Number(parsed.iso.slice(0, 4)) <= 1900) {
    const hint = parseDatetime(dateHint);
    if (!hint || Number(hint.iso.slice(0, 4)) <= 1900) return null;
    return { iso: hint.iso, minutes: parsed.minutes };
  }
  return parsed;
}

function parseDatetime(value) {
  const text = String(value || "").trim();
  if (!text) return null;
  if (/^-?\d+(\.\d+)?$/.test(text)) {
    const serial = Number(text);
    if (serial <= 0) return null;
    return fromExcel(serial);
  }
  const ampm = /^(\d{1,2})\/(\d{1,2})\/(\d{4}) (\d{1,2}):(\d{2})(?::(\d{2}))? (AM|PM)$/i.exec(text);
  if (ampm) {
    let hour = Number(ampm[4]) % 12;
    if (ampm[7].toUpperCase() === "PM") hour += 12;
    return moment(Number(ampm[3]), Number(ampm[1]), Number(ampm[2]), hour, Number(ampm[5]));
  }
  const monthFirst = /^(\d{1,2})\/(\d{1,2})\/(\d{4})(?: (\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(text);
  if (monthFirst) {
    return moment(
      Number(monthFirst[3]),
      Number(monthFirst[1]),
      Number(monthFirst[2]),
      Number(monthFirst[4] || 0),
      Number(monthFirst[5] || 0)
    );
  }
  const yearFirst = /^(\d{4})-(\d{2})-(\d{2})(?: (\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(text);
  if (yearFirst) {
    return moment(
      Number(yearFirst[1]),
      Number(yearFirst[2]),
      Number(yearFirst[3]),
      Number(yearFirst[4] || 0),
      Number(yearFirst[5] || 0)
    );
  }
  return null;
}

function moment(year, month, day, hour, minute) {
  const monthText = String(month).padStart(2, "0");
  const dayText = String(day).padStart(2, "0");
  return { iso: `${year}-${monthText}-${dayText}`, minutes: hour * 60 + minute };
}

function fromExcel(serial) {
  const secondsFloat = serial * 86400;
  let wholeSeconds = Math.floor(secondsFloat + 1e-8);
  let micros = Math.round((secondsFloat - wholeSeconds) * 1e6);
  if (micros >= 1000000) {
    wholeSeconds += Math.floor(micros / 1000000);
    micros %= 1000000;
  }
  if (micros < 0) micros = 0;
  if (micros >= 500000) wholeSeconds += 1;
  let minutes = Math.floor(wholeSeconds / 60);
  if (wholeSeconds % 60 >= 30) minutes += 1;
  const epoch = Date.UTC(1899, 11, 30);
  const stamp = new Date(epoch + minutes * 60000);
  return {
    iso: `${stamp.getUTCFullYear()}-${String(stamp.getUTCMonth() + 1).padStart(2, "0")}-${String(stamp.getUTCDate()).padStart(2, "0")}`,
    minutes: stamp.getUTCHours() * 60 + stamp.getUTCMinutes(),
  };
}

function rowsFromCsv(buffer) {
  const text = new TextDecoder("utf-8").decode(buffer).replace(/^\uFEFF/, "");
  return parseCsv(text).map((values) => {
    const row = {};
    values.forEach((value, index) => {
      row[columnLetter(index + 1)] = value;
    });
    return row;
  });
}

function parseCsv(text) {
  const rows = [];
  let row = [];
  let cell = "";
  let quoted = false;
  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    if (quoted) {
      if (character === '"') {
        if (text[index + 1] === '"') {
          cell += '"';
          index += 1;
        } else quoted = false;
      } else cell += character;
    } else if (character === '"') quoted = true;
    else if (character === ",") {
      row.push(cell);
      cell = "";
    } else if (character === "\n" || character === "\r") {
      if (character === "\r" && text[index + 1] === "\n") index += 1;
      row.push(cell);
      cell = "";
      rows.push(row);
      row = [];
    } else cell += character;
  }
  if (cell.length || row.length) {
    row.push(cell);
    rows.push(row);
  }
  return rows;
}

function rowsFromXlsx(buffer) {
  return readZip(buffer).then((files) => {
    const names = new Set(files.keys());
    const strings = files.has("xl/sharedStrings.xml") ? sharedStrings(parseXml(files.get("xl/sharedStrings.xml"))) : [];
    const sheets = worksheetNames(files, names);
    if (!sheets.length) throw new TimesheetError("That Excel file has no sheets.");
    for (let index = 0; index < sheets.length; index += 1) {
      const rows = sheetRows(parseXml(files.get(sheets[index])), strings);
      if (looksLikeTimesheet(rows)) return rows;
    }
    return sheetRows(parseXml(files.get(sheets[0])), strings);
  }).catch((error) => {
    if (error instanceof TimesheetError) throw error;
    throw new TimesheetError("That Excel file could not be opened.");
  });
}

function looksLikeTimesheet(rows) {
  return rows.slice(0, 15).some((row) => {
    const headers = new Set(Object.values(row).map(normalizeHeader));
    return ["first name", "last name", "start time", "end time"].every((name) => headers.has(name));
  });
}

function sharedStrings(doc) {
  return [...doc.getElementsByTagNameNS(SPREADSHEET_NS, "si")].map((item) =>
    [...item.getElementsByTagNameNS(SPREADSHEET_NS, "t")].map((node) => node.textContent || "").join("")
  );
}

function worksheetNames(files, names) {
  const relsName = "xl/_rels/workbook.xml.rels";
  if (!names.has("xl/workbook.xml") || !names.has(relsName)) {
    return [...names].filter((name) => name.startsWith("xl/worksheets/sheet")).sort();
  }
  const targets = {};
  [...parseXml(files.get(relsName)).getElementsByTagNameNS(PACKAGE_REL_NS, "Relationship")].forEach((rel) => {
    const relId = rel.getAttribute("Id");
    const target = rel.getAttribute("Target") || "";
    let path = target;
    if (target.startsWith("/")) path = target.replace(/^\/+/, "");
    else if (!target.startsWith("xl/")) path = `xl/${target}`;
    targets[relId] = path;
  });
  const sheets = [];
  [...parseXml(files.get("xl/workbook.xml")).getElementsByTagNameNS(SPREADSHEET_NS, "sheet")].forEach((sheet) => {
    const path = targets[sheet.getAttributeNS(OFFICE_REL_NS, "id")];
    if (path && names.has(path)) sheets.push(path);
  });
  return sheets;
}

function sheetRows(doc, strings) {
  const rows = [];
  [...doc.getElementsByTagNameNS(SPREADSHEET_NS, "row")].forEach((row) => {
    const values = {};
    [...row.children].forEach((cell) => {
      if (cell.localName !== "c") return;
      const letter = (cell.getAttribute("r") || "").replace(/[^A-Za-z]/g, "");
      values[letter] = cellText(cell, strings);
    });
    if (Object.values(values).some((value) => String(value).trim())) rows.push(values);
  });
  return rows;
}

function cellText(cell, strings) {
  const kind = cell.getAttribute("t");
  if (kind === "inlineStr") {
    const inline = [...cell.children].find((node) => node.localName === "is");
    if (!inline) return "";
    return [...inline.getElementsByTagNameNS(SPREADSHEET_NS, "t")].map((node) => node.textContent || "").join("");
  }
  const value = [...cell.children].find((node) => node.localName === "v");
  if (!value || value.textContent == null) return "";
  if (kind === "s") {
    const index = Number(value.textContent);
    return index >= 0 && index < strings.length ? strings[index] : "";
  }
  return value.textContent;
}

function parseXml(bytes) {
  const text = new TextDecoder("utf-8").decode(bytes);
  const doc = new DOMParser().parseFromString(text, "application/xml");
  if (doc.querySelector("parsererror")) throw new TimesheetError("That Excel file could not be opened.");
  return doc;
}

function columnLetter(number) {
  let letters = "";
  let current = number;
  while (current) {
    const remainder = (current - 1) % 26;
    letters = String.fromCharCode(65 + remainder) + letters;
    current = Math.floor((current - 1) / 26);
  }
  return letters;
}

function readZip(buffer) {
  const bytes = buffer instanceof Uint8Array ? buffer : new Uint8Array(buffer);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let eocd = -1;
  const scan = Math.max(0, bytes.length - 22 - 65535);
  for (let index = bytes.length - 22; index >= scan; index -= 1) {
    if (view.getUint32(index, true) === 0x06054b50) {
      eocd = index;
      break;
    }
  }
  if (eocd < 0) return Promise.reject(new TimesheetError("That Excel file could not be opened."));
  const count = view.getUint16(eocd + 10, true);
  let offset = view.getUint32(eocd + 16, true);
  const entries = [];
  for (let index = 0; index < count; index += 1) {
    if (view.getUint32(offset, true) !== 0x02014b50) break;
    const method = view.getUint16(offset + 10, true);
    const compressedSize = view.getUint32(offset + 20, true);
    const nameLength = view.getUint16(offset + 28, true);
    const extraLength = view.getUint16(offset + 30, true);
    const commentLength = view.getUint16(offset + 32, true);
    const localOffset = view.getUint32(offset + 42, true);
    const name = new TextDecoder().decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
    const localNameLength = view.getUint16(localOffset + 26, true);
    const localExtraLength = view.getUint16(localOffset + 28, true);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    entries.push({
      name,
      method,
      compressed: bytes.subarray(dataStart, dataStart + compressedSize),
    });
    offset += 46 + nameLength + extraLength + commentLength;
  }
  return Promise.all(entries.map((entry) => inflateEntry(entry))).then((files) => {
    const archive = new Map();
    files.forEach((file) => archive.set(file.name, file.bytes));
    return archive;
  });
}

function inflateEntry(entry) {
  if (entry.method === 0) return Promise.resolve({ name: entry.name, bytes: entry.compressed });
  if (entry.method !== 8) return Promise.reject(new TimesheetError("That Excel file could not be opened."));
  const stream = new Blob([entry.compressed]).stream().pipeThrough(new DecompressionStream("deflate-raw"));
  return new Response(stream).arrayBuffer().then((buffer) => ({ name: entry.name, bytes: new Uint8Array(buffer) }));
}
