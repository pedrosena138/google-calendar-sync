# agenda-sync — Google Apps Script

Syncs your Google Calendar into a sheet. **No OAuth client or client secret needed** — Apps Script talks to Calendar and Sheets with built-in services, authorised by your Google account.

## Install

1. Open your Google Sheet.
2. **Extensions → Apps Script**.
3. Delete the placeholder `myFunction`, paste the contents of `src/Code.gs` in its place, and save.
4. Reload the spreadsheet — an **Events** sheet appears.

The first run asks for Calendar access; click *Review permissions* → your account → *Advanced* → *Go to (project name)* → **Allow**. Google shows this warning for any unverified script, including your own.

## Use

A **Calendar Sync** menu appears in the toolbar:

| Menu item | What it does |
| --- | --- |
| Sync now | Today → today + 14 days (configurable) |
| Sync next 30 days | Today → today + 30 days |
| Settings… | Change calendar ids and the number of days |

Settings are stored in **Script properties** (`CALENDAR_IDS`, `DAYS_AHEAD`). Defaults to `primary` — your own calendar. For others, copy the id from Google Calendar's *Settings → Integrate calendar*:

```
primary, work@company.com, en.GB#holiday@group.v.calendar.google.com
```

## Output

An `Events` sheet with eight columns: **Date, Start, Client, Status, Client Situation, Questions, Next Step, Next Step Due Date**.

- Rows are sorted by start time; the header row is frozen and coloured. No row is bolded or tinted for being today.
- `Date` is a real date value, so you can sort and filter it.
- `Start` is the start time as `HH:mm` text, in the spreadsheet's timezone. Set a script property named `TIMEZONE` (e.g. `Europe/Lisbon`) to override that. All-day events have a blank `Start`.
- The first four columns come from your calendar. The last four are yours to fill in.
- Each sync **replaces** the previous contents.

### Your columns

`Client Situation`, `Questions`, `Next Step` and `Next Step Due Date` are left empty for you, and **your entries are preserved across syncs**. Before rewriting the sheet the script reads those columns and matches each entry back to its event by date, start time, client and status, then writes it back on the new row.

`Next Step Due Date` is formatted as `yyyy-mm-dd`, so you can type a date and sort or filter by it. It is carried across syncs as a real date value, not as text.

The consequence of the matching: if you change an event's date, time or title, that row no longer matches and its notes are dropped. Re-syncing an unchanged event always keeps what you typed.

### Status and Next Step

Both columns show a dropdown when you click a cell:

| Column | Options |
| --- | --- |
| Status | Agendado, No-Show, Realizado, Reagendado |
| Next Step | WhatsApp, Ligação, E-mail |

`Status` is filled in from your event titles, and its dropdown is **advisory only** — a title carrying any other status word is still written rather than rejected, because a strict rule would make the whole sync fail. To change the choices, edit `STATUS_OPTIONS` near the top of `src/Code.gs`.

`Next Step` is yours to pick, and its dropdown is strict: anything other than the three options is rejected.

**Training sessions are decided by the title and the clock, not the dropdown.** If the event title contains the word *Treinamento* (any case, anywhere in the title), the status is worked out from the start time, overriding whatever appears before the pipe:

| Title | When | Status | Client |
| --- | --- | --- | --- |
| `Cancelada \| Acme Treinamento` | already started | Realizado | Acme Treinamento |
| `Agendada \| Treinamento Beta` | still to come | Agendado | Treinamento Beta |
| `No-Show \| Acme` | any | No-Show | Acme |

An event counts as *already happened* once its **start time** has passed, so a training running right now shows as Realizado rather than Agendado. Say the word if you would rather compare the end time.

To change the keyword or the statuses, edit `TRAINING_MARKER`, `TRAINING_PAST_STATUS` and `TRAINING_FUTURE_STATUS` near the top of `src/Code.gs`. A title of just `Treinamento`, with no `status | client` pattern, is still left out of the sheet because there is no client to go with it.

### Status and Client

`Status` and `Client` are parsed out of the event title, which must look like:

```
Confirmed | Acme Corp
Pending | Beta Ltd
```

That title becomes a `Client` of `Acme Corp` and a `Status` of `Confirmed`; the title itself is not written to the sheet. The split happens on the **first** pipe, so a client name may contain one (`Follow-up | Acme | EU office` gives client `Acme | EU office`). Spaces around the pipe are optional — `Confirmed|Acme` works too — and surrounding whitespace is trimmed.

**Events whose title does not follow that pattern are left out of the sheet entirely.** A title needs a pipe with text on both sides — `Team lunch`, `Confirmed |`, `| Acme` and the older `Confirmed: Acme Corp` format are all skipped. The confirmation toast reports how many events were skipped, so an empty sheet with a high skip count means the titles do not match the pattern rather than that the calendar is empty.

## Run it automatically

**Triggers** (clock icon) → Add Trigger:

| Field | Value |
| --- | --- |
| Function | `syncEvents` |
| Event source | Time-driven |
| Type | Day timer, e.g. every 6 hours |

Use `syncCustomRange` instead to sync a fixed historical window — set script properties `WINDOW_START` and `WINDOW_END` as `YYYY-MM-DD` first.

## Notes

- Scope is `calendar.readonly`: the script cannot create or change events.
- Recurring events are expanded into individual occurrences within the window.
- Multi-day events appear once, on their start date.
- Apps Script quotas apply (30s execution, 50k cells written per call). Very large windows may be slow — keep `DAYS_AHEAD` reasonable.

`src/appsscript.json` holds the scopes. If you installed by pasting instead, Apps Script manages scopes automatically and you can ignore that file.

## Development

```
src/Code.gs          the Apps Script (the product)
src/appsscript.json  manifest and OAuth scopes
test/code.test.js    tests, run with: pnpm test
```

The tests load `src/Code.gs` with stubbed Google services, so they run in plain Node with no dependencies.

## Troubleshooting

**`TypeError: event.getAttendees is not a function`** — fixed in the current version; the row path only reads the date, the all-day flag and the title. If you see it, you are running an older copy — re-paste `src/Code.gs`.