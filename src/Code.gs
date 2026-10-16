/**
 * agenda-sync — Google Calendar to Google Sheets
 *
 * Paste into Extensions > Apps Script of your spreadsheet.
 * Requires the "Google Calendar API (readonly)" and "current spreadsheet" scopes.
 */

const SETTINGS = {
  /** Sheet that receives the events. Created if missing. */
  SHEET_NAME: 'Events',
  /** Property holding a comma-separated list of calendar ids (blank = primary). */
  CALENDAR_IDS_KEY: 'CALENDAR_IDS',
  /** Property holding how many days ahead to sync. */
  DAYS_AHEAD_KEY: 'DAYS_AHEAD',
  /** Property holding how many days back to sync. */
  DAYS_BEHIND_KEY: 'DAYS_BEHIND',
  DEFAULT_DAYS_AHEAD: 30,
  DEFAULT_DAYS_BEHIND: 30,
};

const HEADERS = [
  'Date',
  'Start',
  'Client',
  'Status',
  'Client Situation',
  'Questions',
  'Next Step',
  'Next Step Due Date',
];

/** Trailing columns the script never fills in; their contents are preserved. */
const MANUAL_COLUMNS = ['Client Situation', 'Questions', 'Next Step', 'Next Step Due Date'];

/** Manual column holding a date; formatted as one after every sync. */
const DUE_DATE_COLUMN = 'Next Step Due Date';

/** Choices offered by the Next Step dropdown. */
const NEXT_STEP_COLUMN = 'Next Step';
const NEXT_STEP_OPTIONS = ['WhatsApp', 'Ligação', 'E-mail'];

/** Choices offered by the Status dropdown. */
const STATUS_COLUMN = 'Status';
const STATUS_OPTIONS = ['Agendado', 'No-Show', 'Realizado', 'Reagendado'];

/**
 * A title containing this word describes a training session, whose status
 * depends on whether it has already started.
 */
const TRAINING_MARKER = 'Treinamento';
const TRAINING_PAST_STATUS = 'Realizado';
const TRAINING_FUTURE_STATUS = 'Agendado';

/* ------------------------------------------------------------------ menus */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Calendar Sync')
    .addItem('Sync now', 'syncEvents')
    .addSeparator()
    .addItem('Settings…', 'showSettings')
    .addToUi();
}

function onInstall() {
  onOpen();
}

/* --------------------------------------------------------------- settings */

function getSettings() {
  const properties = PropertiesService.getScriptProperties();

  const rawCalendars = properties.getProperty(SETTINGS.CALENDAR_IDS_KEY) || '';
  const calendars = rawCalendars
    .split(',')
    .map((value) => value.trim())
    .filter(Boolean);

  return {
    calendarIds: calendars.length ? calendars : ['primary'],
    daysAhead: readDays(properties, SETTINGS.DAYS_AHEAD_KEY, SETTINGS.DEFAULT_DAYS_AHEAD),
    daysBehind: readDays(properties, SETTINGS.DAYS_BEHIND_KEY, SETTINGS.DEFAULT_DAYS_BEHIND),
  };
}

/** Validates a number of days typed into the settings prompt. */
function toDays(value, fallback) {
  const days = Number(value);
  return Number.isFinite(days) && days >= 0 ? Math.floor(days) : fallback;
}

/**
 * Reads a day count from script properties.
 *
 * Unlike `Number(value) || fallback`, this keeps a deliberate 0, which is a
 * valid setting meaning "today only".
 */
function readDays(properties, key, fallback) {
  const raw = properties.getProperty(key);
  if (raw === null || String(raw).trim() === '') return fallback;
  return toDays(raw, fallback);
}

function showSettings() {
  const settings = getSettings();
  const ui = SpreadsheetApp.getUi();

  const response = ui.prompt(
    'Calendar Sync settings',
    'Calendar ids (comma separated, use "primary" for your own calendar):\n' +
      settings.calendarIds.join(', ') +
      '\n\nDays back to sync:\n' +
      settings.daysBehind +
      '\n\nDays ahead to sync:\n' +
      settings.daysAhead,
    ui.ButtonSet.OK_CANCEL,
  );

  if (response.getSelectedButton() !== ui.Button.OK) return;

  const lines = response.getResponseText().split('\n');
  const calendarIds = (lines.shift() || '').split(',').map((v) => v.trim()).filter(Boolean);
  const daysBehind = toDays((lines.shift() || '').trim(), SETTINGS.DEFAULT_DAYS_BEHIND);
  const daysAhead = toDays(lines.join(' ').trim(), SETTINGS.DEFAULT_DAYS_AHEAD);

  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(SETTINGS.CALENDAR_IDS_KEY, calendarIds.join(','));
  properties.setProperty(SETTINGS.DAYS_BEHIND_KEY, String(daysBehind));
  properties.setProperty(SETTINGS.DAYS_AHEAD_KEY, String(daysAhead));

  ui.alert('Saved. Run "Sync now" to apply.');
}

/* ------------------------------------------------------------------- sync */

/** The window a standard sync covers: some days back and ahead of today. */
function syncWindow(settings, today) {
  const midnight = startOfDay(today);
  return {
    start: addDays(midnight, -settings.daysBehind),
    end: addDays(midnight, settings.daysAhead),
  };
}

/** Syncs the configured window: some days back and ahead of today. */
function syncEvents() {
  const settings = getSettings();
  const window = syncWindow(settings, new Date());
  return runSync(window.start, window.end, settings);
}

/**
 * Custom entry point for an Apps Script trigger.
 * Set windowStart/windowEnd as script properties, e.g. "2026-10-15".
 */
function syncCustomRange() {
  const properties = PropertiesService.getScriptProperties();
  const settings = getSettings();
  const start = startOfDay(parseDate(properties.getProperty('WINDOW_START')) || new Date());
  const end = parseDate(properties.getProperty('WINDOW_END')) || addDays(start, settings.daysAhead);

  return runSync(start, addDays(end, 1), settings);
}

/**
 * Collects events from every configured calendar and writes them to the sheet.
 *
 * @param {Date} start Inclusive start of the window.
 * @param {Date} end   Exclusive end of the window.
 * @param {Object} settings Result of getSettings().
 */
function runSync(start, end, settings) {
  const ui = SpreadsheetApp.getUi();
  const errors = [];
  const events = [];
  const now = new Date();
  const seenCalendars = {};
  let skippedTitles = 0;

  for (const calendarId of settings.calendarIds) {
    if (seenCalendars[calendarId]) continue;
    seenCalendars[calendarId] = true;

    try {
      const calendar = CalendarApp.getCalendarById(calendarId);
      const found = calendar.getEvents(start, end);
      found.forEach(function (event) {
        const normalized = normalizeEvent(event, now);
        if (!normalized) return;

        // Only keep events whose title follows "{status} | {client_name}".
        if (!normalized.matchesPattern) {
          skippedTitles++;
          return;
        }

        events.push(normalized);
      });
    } catch (error) {
      errors.push(calendarId + ': ' + error.message);
    }
  }

  events.sort(function (a, b) {
    const diff = a.start.getTime() - b.start.getTime();
    return diff !== 0 ? diff : String(a.title).localeCompare(String(b.title));
  });

  // Read before prepareSheet(), which clears the sheet.
  const previous = readManualValues(
    SpreadsheetApp.getActive().getSheetByName(SETTINGS.SHEET_NAME)
  );

  const sheet = prepareSheet();
  const timeZone = getTimeZone();

  const rows = events.map(function (event) {
    const row = toRow(event);
    const key = manualKey(
      Utilities.formatDate(row[0], timeZone, 'yyyy-MM-dd'),
      row[1],
      row[2],
      row[3]
    );
    // Carry over anything typed by hand for this same event.
    return row.concat(previous[key] || MANUAL_COLUMNS.map(function () {
      return '';
    }));
  });

  writeRows(sheet, rows);

  const summary =
    'Synced ' +
    rows.length +
    ' event(s) from ' +
    new Date(start).toLocaleDateString() +
    ' to ' +
    new Date(end).toLocaleDateString() +
    '.' +
    (skippedTitles ? '\nSkipped ' + skippedTitles + ' event(s) without a "status | client" title.' : '');

  if (errors.length) {
    ui.alert(summary + '\n\nSkipped calendars:\n' + errors.join('\n'));
  } else {
    SpreadsheetApp.getActive().toast(summary, 'Calendar Sync', 5);
  }

  return rows.length;
}

/**
 * Calls a getter only when the object actually provides it.
 *
 * Calendar events expose methods such as getAttendees(), but plain JSON from
 * the Calendar REST API (or a wrapper object) does not, so every optional
 * field is read defensively.
 */
function callIf(object, method, fallback) {
  if (object && typeof object[method] === 'function') {
    try {
      const value = object[method]();
      return value === undefined || value === null ? fallback : value;
    } catch (error) {
      return fallback;
    }
  }
  return fallback;
}

/** Converts a Date, ISO string or {dateTime}/{date} object into a Date. */
function toDate(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return isNaN(value.getTime()) ? null : value;
  }
  if (typeof value === 'string' || typeof value === 'number') {
    const parsed = new Date(value);
    return isNaN(parsed.getTime()) ? null : parsed;
  }
  if (value.dateTime) return toDate(value.dateTime);
  if (value.date) return toDate(value.date + 'T00:00:00');
  return null;
}

/**
 * Splits a title of the form "{status} | {client_name}".
 *
 * Splits on the first pipe so a client name may itself contain one.
 * Spaces around the pipe are optional and trimmed either way.
 * Returns null when the title does not follow that structure.
 */
function splitTitle(title) {
  if (!title) return null;

  const text = String(title).trim();
  const separator = text.indexOf('|');

  if (separator <= 0) return null;

  const status = text.slice(0, separator).trim();
  const client = text.slice(separator + 1).trim();

  if (!status || !client) return null;

  return { status: status, client: client };
}

/**
 * Builds a stable field set from any supported event shape.
 *
 * @param {Object} raw    Event, wrapper or REST payload.
 * @param {Date}   [now]  Reference point for past/future decisions.
 */
function normalizeEvent(raw, now) {
  if (raw && raw.normalized) return raw;

  const reference = now || new Date();

  // Accept both a bare event and a { event: ... } wrapper.
  const source = raw && raw.event !== undefined && !raw.start ? raw.event : raw;

  if (!source) return null;

  const startValue = callIf(source, 'getStartTime', null) || source.start;
  const start = toDate(startValue);

  if (!start) return null;

  // All-day events have no clock time to show.
  const allDay =
    callIf(source, 'isAllDayEvent', null) === true ||
    (Boolean(startValue && startValue.date) && !startValue.dateTime);

  const title = callIf(source, 'getTitle', null) || source.summary || source.title || '(no title)';
  const parsed = splitTitle(title);

  return {
    normalized: true,
    start: start,
    allDay: allDay,
    title: title,
    // A training session is Realizado once it has started and Agendado until then,
    // whatever the title says before the pipe.
    status: isTraining(title) ? trainingStatus(start, reference) : parsed && parsed.status,
    client: parsed ? parsed.client : null,
    matchesPattern: Boolean(parsed),
  };
}

/** True when the title mentions a training session, whatever the case. */
function isTraining(title) {
  return title ? title.toLowerCase().indexOf(TRAINING_MARKER.toLowerCase()) !== -1 : false;
}

/** Past or future status for a training session, based on its start time. */
function trainingStatus(start, reference) {
  return start.getTime() <= reference.getTime() ? TRAINING_PAST_STATUS : TRAINING_FUTURE_STATUS;
}

/** The timezone used for clock times: the spreadsheet's own, unless overridden. */
function getTimeZone() {
  const override = PropertiesService.getScriptProperties().getProperty('TIMEZONE');
  return override || SpreadsheetApp.getActive().getSpreadsheetTimeZone();
}

/** Identity of an event row, used to carry manual entries across syncs. */
function manualKey(date, start, client, status) {
  return [date, start, client, status].join('|');
}

/** Renders a cell value as yyyy-MM-dd, whatever type it came back as. */
function dateKey(value, timeZone) {
  if (Object.prototype.toString.call(value) === '[object Date]') {
    return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
  }
  return value ? String(value).trim() : '';
}

/**
 * Reads the hand-filled columns from the current sheet, keyed by event.
 *
 * Every sync rewrites the sheet, so anything typed by hand would otherwise be
 * lost. Values are matched back to their event by date, start, client and
 * status: change any of those in the event title or time and the note is no
 * longer recognised.
 */
function readManualValues(sheet) {
  const values = {};
  if (!sheet) return values;

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return values;

  const firstManual = HEADERS.indexOf(MANUAL_COLUMNS[0]);
  // Raw values, not display text: a date cell must come back as a Date so it
  // is written back as a date rather than as a string.
  const rows = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

  rows.forEach(function (row) {
    const manual = row.slice(firstManual, firstManual + MANUAL_COLUMNS.length);
    if (manual.some(function (value) {
      return value !== '' && value !== null;
    })) {
      values[manualKey(dateKey(row[0], getTimeZone()), row[1], row[2], row[3])] = manual;
    }
  });

  return values;
}

/** Maps one calendar event to a spreadsheet row. */
function toRow(rawEvent) {
  const event = normalizeEvent(rawEvent);
  if (!event) return null;

  return [
    event.start,
    event.allDay ? '' : Utilities.formatDate(event.start, getTimeZone(), 'HH:mm'),
    event.client,
    event.status,
  ];
}

/* ------------------------------------------------------------------ sheet */

/** Returns the events sheet, creating and formatting it when needed. */
function prepareSheet() {
  const spreadsheet = SpreadsheetApp.getActive();
  let sheet = spreadsheet.getSheetByName(SETTINGS.SHEET_NAME);

  if (!sheet) {
    sheet = spreadsheet.insertSheet(SETTINGS.SHEET_NAME);
  }

  sheet.clear();

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  sheet
    .getRange(1, 1, 1, HEADERS.length)
    .setFontWeight('bold')
    .setBackground('#1a73e8')
    .setFontColor('#ffffff');
  sheet.setFrozenRows(1);

  sheet.setColumnWidth(1, 110);
  sheet.setColumnWidth(2, 80);
  sheet.setColumnWidth(3, 240);
  sheet.setColumnWidth(4, 130);
  sheet.setColumnWidth(5, 260);
  sheet.setColumnWidth(6, 260);
  sheet.setColumnWidth(7, 130);
  sheet.setColumnWidth(8, 150);

  const dueColumn = HEADERS.indexOf(DUE_DATE_COLUMN) + 1;
  if (dueColumn) {
    sheet.getRange(2, dueColumn, Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat('yyyy-mm-dd');
  }

  // Status keeps the dropdown as guidance: a title may legitimately carry any
  // status, and a strict rule would make setValues() throw and break the sync.
  applyDropdown(sheet, STATUS_COLUMN, STATUS_OPTIONS, true);
  applyDropdown(sheet, NEXT_STEP_COLUMN, NEXT_STEP_OPTIONS, false);

  sheet.getRange('A:A').setNumberFormat('yyyy-mm-dd');
  return sheet;
}

/** Applies a dropdown to a column. */
function applyDropdown(sheet, columnName, options, allowInvalid) {
  const column = HEADERS.indexOf(columnName) + 1;
  if (!column) return null;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(options, true)
    .setAllowInvalid(allowInvalid)
    .build();

  sheet.getRange(2, column, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
  return { column: column, rule: rule };
}

/** Writes all rows in a single call, clearing any previous content. */
function writeRows(sheet, rows) {
  const requiredRows = rows.length + 1; // header + data

  // The sheet may have been trimmed by an earlier sync, so grow it first.
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  sheet.getRange(2, 1, sheet.getMaxRows() - 1, HEADERS.length).clearContent();

  if (!rows.length) return;

  sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
  sheet.getRange(2, 1, rows.length, 1).setNumberFormat('yyyy-mm-dd');
}

/* ---------------------------------------------------------------- helpers */

function startOfDay(date) {
  const result = new Date(date);
  result.setHours(0, 0, 0, 0);
  return result;
}

function addDays(date, days) {
  return new Date(new Date(date).getTime() + days * 86400000);
}

function parseDate(value) {
  if (!value) return null;
  const parts = String(value).trim().split('-').map(Number);
  if (parts.length !== 3 || parts.some(isNaN)) return null;
  return new Date(parts[0], parts[1] - 1, parts[2]);
}