const form = document.getElementById("check-form");
const fileInput = document.getElementById("file");
const drop = document.getElementById("drop");
const fileLabel = document.getElementById("file-label");
const period = document.getElementById("period");
const clockIn = document.getElementById("clock-in");
const clockOut = document.getElementById("clock-out");
const minimum = document.getElementById("minimum");
const scheduleField = document.getElementById("schedule-field");
const schedule = document.getElementById("schedule");
const excludeNote = document.getElementById("exclude-note");
const excludeText = document.getElementById("exclude-text");
const error = document.getElementById("form-error");
const submit = document.getElementById("submit");
const results = document.getElementById("results");

const saved = loadRule();
if (saved) {
  period.value = saved.period || "last-week";
  clockIn.value = saved.clockInBefore || "";
  clockOut.value = saved.clockOutAfter || "";
  minimum.value = saved.minimumDays || "10";
  excludeNote.checked = Boolean(saved.excludeNote);
  excludeText.value = saved.excludeText || "";
}
excludeText.disabled = !excludeNote.checked;

fileInput.addEventListener("change", () => {
  const file = fileInput.files && fileInput.files[0];
  fileLabel.textContent = file ? file.name : "Choose a timesheet";
  drop.classList.toggle("is-ready", Boolean(file));
  hideError();
});

["dragover", "dragenter"].forEach((name) => {
  drop.addEventListener(name, (event) => {
    event.preventDefault();
    drop.classList.add("is-over");
  });
});
["dragleave", "drop"].forEach((name) => {
  drop.addEventListener(name, () => drop.classList.remove("is-over"));
});

excludeNote.addEventListener("change", () => {
  excludeText.disabled = !excludeNote.checked;
  if (excludeNote.checked) excludeText.focus();
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  hideError();
  const file = fileInput.files && fileInput.files[0];
  if (!file) {
    showError("Choose a timesheet first.");
    return;
  }
  if (!clockIn.value && !clockOut.value) {
    showError("Set a clock-in time, a clock-out time, or both.");
    return;
  }
  const minimumDays = Number(minimum.value);
  if (!Number.isInteger(minimumDays) || minimumDays < 1) {
    showError("Enter a whole number of days, at least 1.");
    return;
  }
  const note = excludeText.value.trim();
  if (excludeNote.checked && !note) {
    showError("Enter the clock out note to exclude.");
    excludeText.focus();
    return;
  }

  const rule = {
    period: period.value,
    clockInBefore: clockIn.value,
    clockOutAfter: clockOut.value,
    minimumDays,
    schedule: schedule.value,
    excludeNote: excludeNote.checked ? note : "",
    excludeText: excludeText.value,
  };
  localStorage.setItem("clock-check-rule", JSON.stringify(rule));
  submit.disabled = true;
  submit.textContent = "Checking…";

  try {
    const payload = await checkTimesheet(await file.arrayBuffer(), file.name, rule);
    fillSchedules(payload.schedules || []);
    renderResults(payload);
  } catch (cause) {
    const message = cause instanceof TimesheetError
      ? cause.message
      : "Clock check could not read that file.";
    showError(message);
    if (message === "No clock entries fall in that period.") clearResults();
  } finally {
    submit.disabled = false;
    submit.textContent = "Show Results";
  }
});

function fillSchedules(schedules) {
  const current = schedule.value;
  schedule.replaceChildren(new Option("All schedules", ""));
  schedules.forEach((name) => schedule.append(new Option(name, name)));
  if (schedules.includes(current)) schedule.value = current;
  scheduleField.hidden = schedules.length < 2;
}

function clearResults() {
  results.hidden = true;
  results.replaceChildren();
}

function renderResults(payload) {
  results.hidden = false;
  results.replaceChildren();

  const head = el("div", "results__head");
  head.append(el("h2", null, `${payload.matched.length} ${payload.matched.length === 1 ? "person" : "people"}`));
  if (payload.matched.length) {
    const button = el("button", "download", "Download list");
    button.type = "button";
    button.addEventListener("click", () => downloadCsv(payload));
    head.append(button);
  }
  results.append(head);

  results.append(personSection("Meet The Rule", payload.matched));

  if (payload.short.length) {
    results.append(personSection("Short Of The Rule", payload.short, true));
  }

  if (payload.excluded && payload.excluded.length) {
    results.append(personSection("Excluded By Note", payload.excluded, false, true));
  }

  if (payload.noMatch.length) {
    results.append(personSection("No Days Meet The Rule", payload.noMatch));
  }
  results.scrollIntoView({ behavior: "smooth", block: "start" });
}

function personSection(title, people, markMisses, markExcluded) {
  const block = el("div", "short");
  const head = el("div", "short__head");
  head.append(el("h2", null, title));
  block.append(head);
  if (!people.length) {
    block.append(el("p", "empty", "Nobody in this report reached that many days."));
    return block;
  }
  people.forEach((person) => block.append(personBlock(person, markMisses, markExcluded)));
  return block;
}

function personBlock(person, markMisses, markExcluded) {
  const wrap = el("article", "person");
  const button = el("button", "person__button");
  button.type = "button";
  button.setAttribute("aria-expanded", "false");
  const name = el("span");
  name.append(el("span", "person__name", person.name));
  const role = [person.position, person.employeeId].filter(Boolean).join(" · ");
  if (role) name.append(el("span", "person__role", role));
  const count = el("span", "person__count", String(person.matchingDays));
  count.append(el("small", null, person.matchingDays === 1 ? "day" : "days"));
  button.append(name, count, el("span", "person__role", `${person.daysWorked} worked`));

  const table = document.createElement("table");
  table.className = "days-table";
  table.hidden = true;
  const thead = document.createElement("thead");
  const header = document.createElement("tr");
  ["Date", "Clock in", "Clock out", "Entries", "Clock out note"].forEach((label) => {
    header.append(el("th", null, label));
  });
  thead.append(header);
  const tbody = document.createElement("tbody");
  person.days.forEach((day) => {
    const row = document.createElement("tr");
    if (isStillClockedInToday(day)) row.className = "day-open";
    else if ((markMisses && !day.metRule) || (markExcluded && day.excludedByNote)) row.className = "day-miss";
    const entries = day.entries > 1 ? `${day.entries} clock entries` : "1 clock entry";
    [day.dateLabel, day.clockIn, day.clockOut, entries, day.clockOutNote || ""].forEach((value) => {
      row.append(el("td", null, value));
    });
    tbody.append(row);
  });
  table.append(thead, tbody);

  button.addEventListener("click", () => {
    const open = table.hidden;
    table.hidden = !open;
    button.setAttribute("aria-expanded", open ? "true" : "false");
  });
  wrap.append(button, table);
  return wrap;
}

function isStillClockedInToday(day) {
  if (day.clockOut !== "Still clocked in") return false;
  const now = new Date();
  const month = String(now.getMonth() + 1).padStart(2, "0");
  const date = String(now.getDate()).padStart(2, "0");
  return day.date === `${now.getFullYear()}-${month}-${date}`;
}

function downloadCsv(payload) {
  const lines = [["Name", "Employee ID", "Position", "Matching days", "Days worked"]];
  payload.matched.forEach((person) => {
    lines.push([person.name, person.employeeId, person.position, person.matchingDays, person.daysWorked]);
  });
  lines.push([]);
  lines.push(["Name", "Date", "Clock in", "Clock out", "Clock entries", "Clock out note"]);
  payload.matched.forEach((person) => {
    person.days.forEach((day) => {
      lines.push([person.name, day.date, day.clockIn, day.clockOut, day.entries, day.clockOutNote || ""]);
    });
  });
  const csv = lines
    .map((row) => row.map((value) => `"${String(value ?? "").replaceAll('"', '""')}"`).join(","))
    .join("\n");
  const blob = new Blob([csv], { type: "text/csv" });
  const link = document.createElement("a");
  link.href = URL.createObjectURL(blob);
  link.download = "clock-check.csv";
  link.click();
  URL.revokeObjectURL(link.href);
}

function loadRule() {
  try {
    return JSON.parse(localStorage.getItem("clock-check-rule") || "null");
  } catch (error) {
    return null;
  }
}

function showError(message) {
  error.hidden = false;
  error.textContent = message;
}

function hideError() {
  error.hidden = true;
  error.textContent = "";
}

function el(tag, className, text) {
  const node = document.createElement(tag);
  if (className) node.className = className;
  if (text != null) node.append(document.createTextNode(text));
  return node;
}
