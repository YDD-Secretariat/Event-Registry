# Event Registry

One registration system for **every** programme — Rally, Retreat, Conference, anything.
Pick the programme on the first screen; the app (name, groups, labels, which pages are
available) configures itself from the **Programs** sheet in the database workbook.

```
Event-Registry/
├── index.html … upload.html     the web app (static files)
├── config.js                    shared client — only line to edit: APPS_SCRIPT_URL
├── apps-script/Code.gs          backend that reads/writes the workbook
└── database/Event_Registry_Database.xlsx   the database (one workbook, all programmes)
```

## Set up (about 10 minutes)

1. **Database.** Upload `database/Event_Registry_Database.xlsx` to Google Drive → open it →
   *File ▸ Save as Google Sheets*.
2. **Backend.** In that Sheet: *Extensions ▸ Apps Script* → replace the contents with
   `apps-script/Code.gs` → Save → choose `setup` in the function list → Run (approve the
   permissions; Drive access is for attendance-sheet photos).
3. **Deploy.** *Deploy ▸ New deployment ▸ Web app* → Execute as **Me**, access **Anyone** →
   copy the `/exec` URL.
4. **Connect.** Paste that URL into `APPS_SCRIPT_URL` at the top of `config.js`.
5. **Host.** Publish the folder the way you host the current apps (Vercel/Render static, or
   `npm install && npm start`). Open it, pick a programme, done.

Tip: give a team a direct link that pre-selects the programme:
`https://your-site/index.html?program=AVS-RETREAT-2026`

## Add a new programme (e.g. Corpers Conference)

In the **Programs** sheet, copy a row and edit it. No code, no redeploy.

| Column | What it does |
|---|---|
| `ProgramID` | Unique code, no spaces (e.g. `CORPS-CONF-2027`) |
| `Name`, `Tagline` | Shown on every page |
| `Status` | `Active` or `Closed` (closed = read-only, hidden from the picker) |
| `Groups`, `OtherGroup` | Comma-separated member groups; the catch-all group (e.g. `Others`) |
| `ModSchools` … `ModUploads` | `TRUE/FALSE` — which pages exist for this programme |
| `VisitorsInMembers` | `TRUE`: visitors counted inside Members under `OtherGroup` (Rally style). `FALSE`: separate total (AVS style) |
| `MemberLabel`, `WorkerLabel`, `SchoolLabel`, `StudentLabel`, `GroupLabel` | Rename the words (e.g. Workers → Delegates) |

Optional: paste that programme's roster into **WorkerRoster** (set `ProgramID` in column A)
so people can be found by name for individual check-in.

## What was unified

The Rally code was the template; AVS was brought level with it.

- **One codebase.** Same pages, same navigation, same dashboard for every programme.
- **Rally features now available to AVS:** individual check-in with group selection,
  bulk workers (switchable), the kitchen dashboard's complete view, CSV exports.
- **AVS behaviours kept as a switch:** visitors as a separate total
  (`VisitorsInMembers = FALSE`), grand total = students + members + workers + visitors.
- **Offline queue is programme-safe:** a write saved offline remembers its programme, so a
  late sync can never land in the wrong event. Retries are de-duplicated by the server.
- **Dashboard totals now come from the server** (single source of truth) instead of being
  re-added in the browser.

## Notes

- Your two existing deployments keep working untouched; this uses a **new** Apps Script URL.
- Migrated: the AVS data you uploaded (358 roster names, 11 check-ins, activity log), tagged
  `AVS-RETREAT-2026`. The October Rally workbook wasn't in the upload, so `OCT-RALLY-2026` is
  configured but empty — paste its old rows under the new `ProgramID` column if you want its history.
- Older AVS check-ins have free-text stations (e.g. "Anthony Group 2"); new ones store the chosen
  group, so the dashboard's by-group list will show both until the event is over.
- Don't sort or delete **WorkerRoster** rows after registration starts (check-ins point at the row's `RosterID`).
- The Next.js `webapp/` inside October-Rally is a separate track and is not part of this kit.
- Access control is unchanged from your current apps (anyone with the link). If a programme needs
  a PIN, that is a small addition — say so.
