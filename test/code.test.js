/**
 * Tests for the Apps Script in src/Code.gs.
 *
 * Apps Script code cannot be imported directly (it relies on Google-provided
 * globals and the .gs extension), so it is loaded as a module after stubbing
 * those globals.
 */
import assert from "node:assert/strict";
import fs from "node:fs";
import { test } from "node:test";

/* --------------------------------------------------- Google service stubs */

const TIMEZONE = "Europe/Lisbon";
const scriptProperties = {};

globalThis.Utilities = {
  formatDate(date, timeZone, pattern) {
    if (pattern === "HH:mm") {
      return new Intl.DateTimeFormat("en-GB", {
        timeZone,
        hour: "2-digit",
        minute: "2-digit",
        hour12: false,
      }).format(date);
    }
    if (pattern === "yyyy-MM-dd") {
      return new Intl.DateTimeFormat("en-CA", {
        timeZone,
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
      }).format(date);
    }
    throw new Error(`unstubbed pattern: ${pattern}`);
  },
};

globalThis.PropertiesService = {
  getScriptProperties: () => ({
    getProperty: (key) => (key in scriptProperties ? scriptProperties[key] : null),
  }),
};

const display = (value) =>
  value instanceof Date
    ? new Intl.DateTimeFormat("en-CA", {
        timeZone: TIMEZONE,
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
      }).format(value)
    : String(value);

class FakeRange {
  constructor(sheet, row, col, numRows, numCols) {
    Object.assign(this, { sheet, row, col, numRows, numCols });
  }
  // Sheet coordinates are 1-based; grid indices are 0-based.
  #each(fn) {
    for (let r = 0; r < this.numRows; r++) {
      for (let c = 0; c < this.numCols; c++) fn(this.row - 1 + r, this.col - 1 + c);
    }
    return this;
  }
  setValues(values) {
    this.sheet.calls.push(["setValues", this.row, this.numRows]);
    values.forEach((line, i) =>
      line.forEach((v, j) => (this.sheet.grid[this.row - 1 + i][this.col - 1 + j] = v)),
    );
    return this;
  }
  setNumberFormat(f) {
    this.sheet.calls.push(["setNumberFormat", f, this.row, this.col, this.numRows]);
    return this;
  }
  setFontWeight(w) {
    this.sheet.calls.push(["setFontWeight", w, this.row, this.numRows]);
    return this;
  }
  setBackground(c) {
    this.sheet.calls.push(["setBackground", c, this.row, this.numRows]);
    return this;
  }
  setFontColor(c) {
    this.sheet.calls.push(["setFontColor", c]);
    return this;
  }
  setDataValidation(rule) {
    validations.push({ row: this.row, col: this.col, rows: this.numRows, rule });
    return this;
  }
  clearContent() {
    this.#each((r, c) => (this.sheet.grid[r][c] = ""));
    return this;
  }
  getValues() {
    // Raw values, preserving types (a Date stays a Date).
    return Array.from({ length: this.numRows }, (_, i) =>
      Array.from({ length: this.numCols }, (_, j) => this.sheet.grid[this.row - 1 + i][this.col - 1 + j]),
    );
  }
  getDisplayValues() {
    // Sheets returns formatted text for the whole range, not the raw values.
    return Array.from({ length: this.numRows }, (_, i) =>
      Array.from({ length: this.numCols }, (_, j) =>
        display(this.sheet.grid[this.row - 1 + i][this.col - 1 + j]),
      ),
    );
  }
}

class FakeSheet {
  constructor(name, maxRows = 40, maxCols = 12) {
    this.name = name;
    this.maxRows = maxRows;
    this.maxCols = maxCols;
    this.grid = Array.from({ length: maxRows }, () => Array(maxCols).fill(""));
    this.calls = [];
    this.frozen = 0;
  }
  getRange(a, b, c, d) {
    if (typeof a === "string") return { setNumberFormat: (f) => this.setColumnFormat(a, f) };
    return new FakeRange(this, a, b, c, d);
  }
  setColumnFormat(a1, f) {
    this.calls.push(["colFormat", a1, f]);
  }
  getMaxRows() {
    return this.maxRows;
  }
  getLastRow() {
    for (let i = this.grid.length - 1; i >= 0; i--) {
      if (this.grid[i].some((v) => v !== "")) return i + 1;
    }
    return 0;
  }
  setFrozenRows(n) {
    this.frozen = n;
    return this;
  }
  setColumnWidth() {
    return this;
  }
  clear() {
    this.grid.forEach((row) => row.fill(""));
    this.calls.push(["clear"]);
    return this;
  }
  insertRowsAfter(at, howMany) {
    this.calls.push(["insertRowsAfter", at, howMany]);
    for (let i = 0; i < howMany; i++) this.grid.splice(at, 0, Array(this.maxCols).fill(""));
    this.maxRows += howMany;
  }
}

const sheet = new FakeSheet("Events");

const validations = [];

class FakeDataValidation {
  constructor() {
    this.values = null;
    this.showDropdown = false;
    this.allowInvalid = true;
  }
  requireValueInList(values, showDropdown) {
    this.values = values;
    this.showDropdown = showDropdown;
    return this;
  }
  setAllowInvalid(allow) {
    this.allowInvalid = allow;
    return this;
  }
  build() {
    return this;
  }
}

globalThis.SpreadsheetApp = {
  newDataValidation: () => new FakeDataValidation(),
  getActive: () => ({
    getSheetByName: (name) => (name === "Events" ? sheet : null),
    insertSheet: () => {
      throw new Error("should reuse the existing sheet");
    },
    getSpreadsheetTimeZone: () => TIMEZONE,
    toast: (msg) => {
      sheet.toast = msg;
    },
  }),
  getUi: () => ({
    alert: (msg) => {
      sheet.alert = msg;
    },
    Button: { OK: "OK" },
    ButtonSet: { OK_CANCEL: "OK_CANCEL" },
    prompt: () => ({ getSelectedButton: () => "CANCEL", getResponseText: () => "" }),
    createMenu: () => ({
      addItem() {
        return this;
      },
      addSeparator() {
        return this;
      },
      addToUi() {
        return this;
      },
    }),
  }),
};

const fakeEvents = [
  { title: "Team lunch", start: new Date("2026-10-15T12:00:00Z") },
  { title: "Pending|Beta Ltd", start: new Date("2026-10-15T09:00:00Z") },
  { title: "Weekly sync", start: new Date("2026-10-16T10:00:00Z") },
  { title: "Confirmed|Acme Corp", start: new Date("2026-10-15T14:00:00Z") },
  { title: "Broken: ", start: new Date("2026-10-15T16:00:00Z") },
  // Far future so the expected status cannot depend on when the suite runs.
  { title: "Agendada | Treinamento Gamma", start: new Date("2099-01-05T09:00:00Z") },
];

globalThis.CalendarApp = {
  getCalendarById: () => ({
    getName: () => "Work",
    getEvents: () =>
      fakeEvents.map((e) => ({
        getTitle: () => e.title,
        getStartTime: () => e.start,
        getEndTime: () => e.start,
        isAllDayEvent: () => false,
      })),
  }),
};

/* ------------------------------------------------------------ load module */

const source =
  fs.readFileSync(new URL("../src/Code.gs", import.meta.url), "utf8") +
  "\nexport { toRow, splitTitle, normalizeEvent, toDate, callIf, startOfDay, addDays, parseDate, " +
  "prepareSheet, writeRows, runSync, getSettings, syncWindow, onOpen, HEADERS };\n";

const script = await import(`data:text/javascript,${encodeURIComponent(source)}`);

/* ----------------------------------------------------------------- tests */

test("the sheet has the documented columns", () => {
  assert.deepEqual(script.HEADERS, [
    "Date",
    "Start",
    "Client",
    "Status",
    "Client Situation",
    "Questions",
    "Next Step",
    "Next Step Due Date",
  ]);
});

test("the due date column is formatted as a date", () => {
  validations.length = 0;
  script.prepareSheet();

  const formatted = sheet.calls.filter((c) => c[0] === "setNumberFormat" && c[1] === "yyyy-mm-dd");
  assert.ok(formatted.length >= 1, "a date format is applied");
  assert.ok(
    formatted.some((c) => c[3] === 8),
    "applied to the Next Step Due Date column",
  );
});

test("the default window is 30 days back and 30 days ahead", () => {
  const settings = script.getSettings();
  assert.equal(settings.daysBehind, 30);
  assert.equal(settings.daysAhead, 30);

  const { start, end } = script.syncWindow(settings, new Date("2026-10-15T13:37:00Z"));
  assert.equal(start.toISOString().slice(0, 10), "2026-09-15", "30 days back");
  assert.equal(end.toISOString().slice(0, 10), "2026-11-14", "30 days ahead");
  assert.equal(start.getHours(), 0, "starts at midnight");
});

test("the window honours saved properties and ignores bad values", () => {
  scriptProperties.DAYS_BEHIND = "7";
  scriptProperties.DAYS_AHEAD = "0";
  let settings = script.getSettings();
  assert.equal(settings.daysBehind, 7);
  assert.equal(settings.daysAhead, 0, "zero days ahead is a valid setting");

  const { start, end } = script.syncWindow(settings, new Date("2026-10-15T13:37:00Z"));
  assert.equal(start.toISOString().slice(0, 10), "2026-10-08");
  assert.equal(end.toISOString().slice(0, 10), "2026-10-15", "zero ahead means today only");

  scriptProperties.DAYS_AHEAD = "not-a-number";
  assert.equal(script.getSettings().daysAhead, 30, "falls back to the default");

  delete scriptProperties.DAYS_BEHIND;
  delete scriptProperties.DAYS_AHEAD;
});

test("the menu no longer offers the 30 day shortcut", () => {
  const items = [];
  const menu = {
    addItem(label, fn) {
      items.push([label, fn]);
      return menu;
    },
    addSeparator() {
      items.push(["-", null]);
      return menu;
    },
    addToUi() {
      return menu;
    },
  };
  const uiStub = globalThis.SpreadsheetApp;
  globalThis.SpreadsheetApp = {
    ...uiStub,
    getUi: () => ({ createMenu: () => menu }),
  };
  script.onOpen();
  globalThis.SpreadsheetApp = uiStub;

  assert.deepEqual(items.map((i) => i[0]), ["Sync now", "-", "Settings…"]);
});

test("Status and Next Step both get dropdowns", () => {
  validations.length = 0;
  script.prepareSheet();

  assert.equal(validations.length, 2, "one dropdown per column");

  const status = validations.find((v) => v.col === 4);
  assert.ok(status, "applied to the Status column");
  assert.deepEqual(status.rule.values, ["Agendado", "No-Show", "Realizado", "Reagendado"]);
  assert.equal(status.rule.showDropdown, true, "shown as a dropdown menu");
  assert.equal(status.rule.allowInvalid, true, "guidance only, so a sync never fails");
  assert.ok(status.rows > 1, "covers the data rows");

  const next = validations.find((v) => v.col === 7);
  assert.ok(next, "applied to the Next Step column");
  assert.deepEqual(next.rule.values, ["WhatsApp", "Ligação", "E-mail"]);
  assert.equal(next.rule.allowInvalid, false, "other values are rejected");
});

test("a training title is Realizado once past and Agendado while future", () => {
  const NOW = new Date("2026-10-15T12:00:00Z");
  const rowFor = (title, start) =>
    script.normalizeEvent(
      { getTitle: () => title, getStartTime: () => start, isAllDayEvent: () => false },
      NOW,
    );

  const past = rowFor("Cancelada | Acme Treinamento", new Date("2026-10-15T09:00:00Z"));
  assert.equal(past.status, "Realizado", "a training that already started is Realizado");
  assert.equal(past.client, "Acme Treinamento", "client still parsed from the pipe");

  const future = rowFor("Agendada | Treinamento Beta", new Date("2026-10-20T09:00:00Z"));
  assert.equal(future.status, "Agendado", "a training still to come is Agendado");

  assert.equal(
    rowFor("Agendada | treinamento", new Date("2026-10-15T11:59:00Z")).status,
    "Realizado",
    "case insensitive, one minute before now",
  );
  assert.equal(
    rowFor("Treinamento", new Date("2026-10-15T12:00:00Z")).status,
    "Realizado",
    "starting exactly now counts as already begun",
  );

  const untouched = rowFor("No-Show | Acme", new Date("2026-10-15T09:00:00Z"));
  assert.equal(untouched.status, "No-Show", "titles without the marker keep their own status");
  assert.equal(
    rowFor("Reagendado | Beta Ltd", new Date("2026-10-20T09:00:00Z")).status,
    "Reagendado",
  );

  assert.equal(
    script.normalizeEvent(
      { getTitle: () => "Treinamento", getStartTime: () => new Date(), isAllDayEvent: () => false },
      NOW,
    ).matchesPattern,
    false,
    "still excluded from the sheet: no client to go with it",
  );
});

test("toRow returns only the synced columns", () => {
  const row = script.toRow({
    getTitle: () => "Confirmed | Acme Corp",
    getStartTime: () => new Date("2026-10-15T09:00:00Z"),
    isAllDayEvent: () => false,
  });
  assert.equal(row.length, 4, "manual columns are appended by runSync, not by toRow");
});

test("splitTitle extracts status and client", () => {
  assert.deepEqual(script.splitTitle("Confirmed|Acme Corp"), {
    status: "Confirmed",
    client: "Acme Corp",
  });
  assert.deepEqual(
    script.splitTitle("  Follow-up |  Acme| EU office  "),
    { status: "Follow-up", client: "Acme| EU office" },
    "trims whitespace and splits on the first pipe only",
  );
});

test("splitTitle accepts spaces around the pipe", () => {
  assert.deepEqual(script.splitTitle("Confirmed | Acme Corp"), {
    status: "Confirmed",
    client: "Acme Corp",
  });
  assert.deepEqual(
    script.splitTitle("Pending | Beta Ltd"),
    { status: "Pending", client: "Beta Ltd" },
  );
  assert.deepEqual(
    script.splitTitle("Follow-up | Acme | EU office"),
    { status: "Follow-up", client: "Acme | EU office" },
    "only the first pipe separates, and the spaces are trimmed",
  );
  assert.deepEqual(script.splitTitle("Done |Acme"), { status: "Done", client: "Acme" });
  assert.deepEqual(script.splitTitle("Done| Acme"), { status: "Done", client: "Acme" });
});

test("splitTitle rejects titles without the pattern", () => {
  for (const title of ["Team lunch", "Confirmed|", "|Acme", " | ", "", null]) {
    assert.equal(script.splitTitle(title), null, `"${title}" should not match`);
  }
  assert.equal(
    script.splitTitle("Confirmed: Acme Corp"),
    null,
    "the old colon format no longer matches",
  );
});

test("toRow maps a timed event", () => {
  const row = script.toRow({
    getTitle: () => "Confirmed|Acme Corp",
    getStartTime: () => new Date("2026-10-15T09:00:00Z"),
    getEndTime: () => new Date("2026-10-15T10:30:00Z"),
    isAllDayEvent: () => false,
  });

  assert.equal(row.length, 4);
  assert.equal(row[0].toISOString(), "2026-10-15T09:00:00.000Z");
  assert.equal(row[1], "10:00", "start time in the spreadsheet timezone (UTC+1)");
  assert.equal(row[2], "Acme Corp");
  assert.equal(row[3], "Confirmed");
});

test("toRow leaves the start time blank for all-day events", () => {
  const row = script.toRow({
    getTitle: () => "Pending|Beta Ltd",
    getStartTime: () => new Date("2026-10-16T00:00:00Z"),
    getEndTime: () => new Date("2026-10-17T00:00:00Z"),
    isAllDayEvent: () => true,
  });

  assert.equal(row[1], "", "no clock time");
  assert.equal(row[2], "Beta Ltd");
  assert.equal(row[3], "Pending");
});

test("the TIMEZONE property overrides the spreadsheet timezone", () => {
  const event = {
    getTitle: () => "Done|Acme",
    getStartTime: () => new Date("2026-10-15T09:00:00Z"),
    isAllDayEvent: () => false,
  };

  scriptProperties.TIMEZONE = "UTC";
  assert.equal(script.toRow(event)[1], "09:00");
  delete scriptProperties.TIMEZONE;
  assert.equal(script.toRow(event)[1], "10:00");
});

test("toRow accepts REST API JSON and wrapper objects", () => {
  const json = script.toRow({
    summary: "Won|API event",
    start: { dateTime: "2026-10-15T09:00:00Z" },
    end: { dateTime: "2026-10-15T10:00:00Z" },
  });
  assert.equal(json[2], "API event");
  assert.equal(json[3], "Won");

  const allDayJson = script.toRow({ summary: "N/A|Acme", start: { date: "2026-10-16" } });
  assert.equal(allDayJson[1], "", "all-day detected from start.date");

  const wrapped = script.toRow({
    event: {
      getTitle: () => "Done|Wrapped",
      getStartTime: () => new Date("2026-10-15T09:00:00Z"),
      isAllDayEvent: () => false,
    },
  });
  assert.equal(wrapped[3], "Done");
});

test("toRow drops unusable events instead of throwing", () => {
  assert.equal(script.toRow(null), null);
  assert.equal(script.toRow({}), null);
  assert.equal(script.toRow({ summary: "no dates" }), null);
  assert.equal(script.toRow({ start: { dateTime: "not-a-date" } }), null);
});

test("unstructured titles produce empty client and status", () => {
  const row = script.toRow({
    getTitle: () => "Team lunch",
    getStartTime: () => new Date("2026-10-15T09:00:00Z"),
    isAllDayEvent: () => false,
  });
  assert.equal(row[2], null);
  assert.equal(row[3], null);
});

test("callIf only calls getters that exist", () => {
  assert.equal(script.callIf({}, "getTitle", "fallback"), "fallback");
  assert.equal(script.callIf(null, "getTitle", "fallback"), "fallback");
  assert.equal(
    script.callIf(
      {
        getTitle: () => {
          throw new Error("boom");
        },
      },
      "getTitle",
      "safe",
    ),
    "safe",
    "swallows getter errors",
  );
});

test("date helpers behave", () => {
  assert.equal(script.toDate({ date: "2026-10-16" }).getDate(), 16);
  assert.equal(script.toDate(null), null);
  assert.equal(script.startOfDay(new Date("2026-10-15T23:59:00Z")).getHours(), 0);
  assert.equal(
    script.addDays(new Date("2026-10-15T12:00:00Z"), 5).toISOString(),
    "2026-10-20T12:00:00.000Z",
  );
  assert.equal(script.parseDate("2026-10-15").getMonth(), 9);
  assert.equal(script.parseDate("not-a-date"), null);
});

test("prepareSheet writes and formats the header row", () => {
  assert.equal(script.prepareSheet(), sheet);
  assert.deepEqual(sheet.grid[0].slice(0, script.HEADERS.length), script.HEADERS);
  assert.equal(sheet.frozen, 1);
  assert.ok(sheet.calls.some((c) => c[0] === "clear"), "sheet cleared before writing");
});

test("writeRows replaces previous contents", () => {
  const rows = [
    [new Date("2026-10-15T09:00:00Z"), "10:00", "Acme", "Done"],
    [new Date("2026-10-16T09:00:00Z"), "10:00", "Beta", "Done"],
  ];

  script.writeRows(sheet, rows);
  assert.equal(sheet.grid[1][2], "Acme");
  assert.equal(sheet.getLastRow(), 3);

  script.writeRows(sheet, [rows[0]]);
  assert.equal(sheet.getLastRow(), 2, "stale row cleared");
  assert.equal(sheet.grid[2][2], "", "old value wiped");
  assert.equal(
    sheet.calls.some((c) => c[0] === "deleteRows"),
    false,
    "rows are cleared, never deleted",
  );

  script.writeRows(sheet, []);
  assert.equal(sheet.getLastRow(), 1, "empty sync leaves only the header");
});

test("writeRows grows a sheet trimmed by an earlier sync", () => {
  sheet.maxRows = 2;
  sheet.grid.length = 2;

  script.writeRows(sheet, [
    [new Date(), "09:00", "Acme", "Done"],
    [new Date(), "10:00", "Beta", "Done"],
  ]);

  assert.equal(sheet.maxRows, 3, "sheet grew to fit the rows");
  assert.equal(sheet.grid[2][2], "Beta", "and the data is really there");
});

test("runSync keeps only titles matching the pattern", () => {
  sheet.calls.length = 0;
  const written = script.runSync(
    new Date("2026-10-15T00:00:00Z"),
    new Date("2026-10-22T00:00:00Z"),
    { calendarIds: ["primary"], daysAhead: 14 },
  );

  assert.equal(written, 3, "only the events matching the pattern were kept");
  assert.match(sheet.toast, /Skipped 3 event\(s\)/, "skipped count is reported");

  const kept = [sheet.grid[1], sheet.grid[2], sheet.grid[3]];
  assert.deepEqual(
    kept.map((row) => row[1]),
    ["10:00", "15:00", "09:00"],
    "sorted by start time",
  );
  assert.deepEqual(
    kept.map((row) => row[2]),
    ["Beta Ltd", "Acme Corp", "Treinamento Gamma"],
    "client column",
  );
  assert.deepEqual(
    kept.map((row) => row[3]),
    ["Pending", "Confirmed", "Agendado"],
    "a future Treinamento is Agendado end to end",
  );

  assert.deepEqual(
    [
      sheet.grid[1][4],
      sheet.grid[1][5],
      sheet.grid[1][6],
      sheet.grid[2][4],
      sheet.grid[2][5],
      sheet.grid[2][6],
      sheet.grid[3][4],
      sheet.grid[3][5],
      sheet.grid[3][6],
    ],
    ["", "", "", "", "", "", "", "", ""],
    "all manual columns start empty",
  );

  assert.equal(
    sheet.calls.some((c) => c[0] === "setBackground" && c[1] === "#e8f0fe"),
    false,
    "no same-day highlighting",
  );
  assert.equal(
    sheet.calls.some((c) => c[0] === "setFontWeight" && c[1] === "bold" && c[2] > 1),
    false,
    "data rows are never bolded",
  );
});

test("hand-filled entries survive a re-sync", () => {
  // Simulate the user typing into the manual columns.
  sheet.grid[1][4] = "Waiting on budget";
  sheet.grid[1][5] = "Who signs off?";
  sheet.grid[2][4] = "Chased twice";
  sheet.grid[2][5] = "";

  const written = script.runSync(
    new Date("2026-10-15T00:00:00Z"),
    new Date("2026-10-22T00:00:00Z"),
    { calendarIds: ["primary"], daysAhead: 14 },
  );

  assert.equal(written, 3);
  assert.equal(sheet.grid[1][4], "Waiting on budget", "situation kept for the same event");
  assert.equal(sheet.grid[1][5], "Who signs off?", "questions kept for the same event");
  assert.equal(sheet.grid[2][4], "Chased twice", "kept even when only one column was filled");
  assert.equal(sheet.grid[2][5], "", "the unfilled column stays empty");
});

test("a chosen Next Step survives a re-sync", () => {
  sheet.grid[1][4] = "";
  sheet.grid[1][5] = "";
  sheet.grid[2][4] = "";
  sheet.grid[2][5] = "";
  sheet.grid[1][6] = "Ligação";
  sheet.grid[2][6] = "E-mail";

  script.runSync(
    new Date("2026-10-15T00:00:00Z"),
    new Date("2026-10-22T00:00:00Z"),
    { calendarIds: ["primary"], daysAhead: 14 },
  );

  assert.equal(sheet.grid[1][6], "Ligação", "next step kept alongside the other manual columns");
  assert.equal(sheet.grid[2][6], "E-mail");
});

test("a due date survives a re-sync as a real date", () => {
  const due = new Date("2026-10-20T00:00:00Z");
  sheet.grid[1][7] = due;
  sheet.grid[2][7] = "";

  script.runSync(
    new Date("2026-10-15T00:00:00Z"),
    new Date("2026-10-22T00:00:00Z"),
    { calendarIds: ["primary"], daysAhead: 14 },
  );

  assert.ok(
    sheet.grid[1][7] instanceof Date,
    "kept as a Date, not converted to display text",
  );
  assert.equal(sheet.grid[1][7].getTime(), due.getTime(), "same date");
  assert.equal(sheet.grid[2][7], "", "an empty due date stays empty");
});

test("a changed title drops the old manual entry", () => {
  sheet.grid[1][4] = "stale note";
  sheet.grid[1][5] = "";

  // Same events, but the title now yields a different client.
  fakeEvents[1].title = "Pending | Gamma";
  script.runSync(
    new Date("2026-10-15T00:00:00Z"),
    new Date("2026-10-22T00:00:00Z"),
    { calendarIds: ["primary"], daysAhead: 14 },
  );
  fakeEvents[1].title = "Pending|Beta Ltd";

  const firstRow = sheet.grid[1];
  assert.equal(
    firstRow[4],
    "",
    "an entry is only carried over when date, time, client and status all still match",
  );
});