# TruckTalk Connect — Google Sheets Add‑on

TruckTalk Connect is a Google Sheets Editor Add‑on that analyzes the active sheet tab used as a lightweight TMS and either returns a typed Loads JSON payload or reports actionable issues (with severities and suggestions). It includes a modern sidebar UI with chat‑like guidance, mapping hints, optional AI assistance via proxy, and non‑destructive auto‑fix suggestions.


## What It Does

- Analyze the active sheet (first ~200 data rows) and build a compact snapshot.
- Detect header→field mappings (with synonyms and exact matches).
- Validate rows for required fields, duplicates, and datetime quality.
- Produce `AnalysisResult` with issues and, if no errors, a `Load[]` JSON payload.
- Offer a sidebar experience: Analyze, view issues, copy JSON, preview push, re‑analyze.
- Optional AI assistance (via proxy) for mapping suggestions, normalization hints, and issue summary.
- Optional non‑destructive auto‑fix plan and apply: create missing columns, normalize datetimes.


## Data Contracts

- Load
  - `loadId`, `fromAddress`, `fromAppointmentDateTimeUTC`, `toAddress`, `toAppointmentDateTimeUTC`, `status`, `driverName`, `driverPhone?`, `unitNumber`, `broker`
- AnalysisResult
  - `ok: boolean`, `issues: {code,severity,message,rows?,column?,suggestion?}[]`, `loads?: Load[]`, `mapping: Record<header,field>`, `meta: { analyzedRows, analyzedAt }`

See codingtask.md for the full field list and examples.


## Project Structure

- `appsscript.json`: Manifest with Sheets add‑on config and OAuth scopes.
- `server.gs`: Apps Script backend (analysis, validation, optional AI proxy usage, auto‑fixes, rate‑limit).
- `ui.html`: Sidebar UI (Chat | Results), actions, mapping display, auto‑fix controls.
- `tests.gs`: Lightweight unit tests for pure utilities (run via `runUnitTests`).


## Scopes

These are declared in `appsscript.json` and are required for normal operation:

- `https://www.googleapis.com/auth/spreadsheets.currentonly`: Read the active sheet only.
- `https://www.googleapis.com/auth/script.container.ui`: Render sidebar UI.
- `https://www.googleapis.com/auth/script.external_request`: Allow external HTTP requests (needed for optional AI proxy and demo push).
- `https://www.googleapis.com/auth/userinfo.email`: Identify the user for soft rate‑limiting.


## How To Run

Option A — Apps Script Editor only (no CLI):
- Open a Google Sheet → Extensions → Apps Script.
- Create files `server.gs`, `ui.html`, `tests.gs` and copy contents from this repo.
- Replace the default `appsscript.json` with this repo’s manifest (Editor → Project Settings → Show manifest file).
- In the editor: Deploy → Test deployments → Select “Editor Add‑on” for Sheets and install.
- Back in the Sheet: Extensions → TruckTalk Connect → Open Sidebar → click “Analyze current tab”.

Option B — CLASP (recommended for iteration):
- Prereqs: `npm i -g @google/clasp`, signed in with a Google account.
- `clasp login`
- Create or link a script:
  - New: `clasp create --type sheets --title "TruckTalk Connect"`
  - Existing: `clasp clone <your-script-id>`
- Copy these four files into the project root: `server.gs`, `ui.html`, `tests.gs`, `appsscript.json`.
- `clasp push` then `clasp open`.
- Deploy as a test add‑on: Deploy → Test deployments → install → open in Sheets.


## Configuration (Script/User Properties)

Add these in Apps Script: Project Settings → Script properties.

- `PROXY_URL` (optional): URL of an OpenAI proxy endpoint if you want AI mapping/normalization/summary features. If not set, the add‑on still works without AI.
- `OPENAI_MODEL` (optional): Defaults to `gpt-5` in code; change if your proxy expects a specific model id.
- `ASSUME_TZ_AS_WARN` (optional): Set to `true` to downgrade the “assumed timezone” aggregation from error to warning for demos.

User‑scoped properties are used internally for normalization maps:
- `TT_STATUS_MAP`, `TT_BROKER_MAP` (stored via UI actions), and simple preferences `TT_PREF_*`.

If you don’t plan to use AI features, you can omit `PROXY_URL` entirely.


## Using The Add‑on

1) Open the sidebar
- Extensions → TruckTalk Connect → Open Sidebar.

2) Analyze current tab
- Click “Analyze current tab”. The add‑on inspects headers + up to 200 rows.

3) Review results
- If errors exist: Issues panel shows errors and warnings (e.g., MISSING_COLUMN, BAD_DATE_FORMAT, TIMEZONE_MISSING, DUPLICATE_ID). Suggestions explain how to fix.
- If no errors: A Loads JSON panel appears with Copy JSON and Preview Push, plus warnings (e.g., NON_ISO_OUTPUT, STATUS_VOCAB) if any.
- Mapping and meta are always shown for transparency.

4) Optional features
- Mapping suggestions: When ambiguous, AI can propose header→field mappings (never auto‑applied).
- Normalization hints: AI can show example conversions for problematic date strings (no destructive edits).
- Auto‑Fix Assistant: Suggest and optionally apply non‑destructive fixes (create missing required columns, normalize date/times into dedicated ISO fields). Timezone handling remains strict by default.
- Push to TruckTalk (Demo): Posts payload to a mock URL and simulates responses; for demo only.


## Header Mapping & Synonyms

The add‑on recognizes case‑insensitive synonyms. Examples:
- `loadId`: Load ID, Ref, VRID, Reference, Ref #
- `fromAddress`: From, PU, Pickup, Origin, Pickup Address
- `fromAppointmentDateTimeUTC`: PU Time, Pickup Appt, Pickup Date/Time
- `toAddress`: To, Drop, Delivery, Destination, Delivery Address
- `toAppointmentDateTimeUTC`: DEL Time, Delivery Appt, Delivery Date/Time
- `status`: Status, Load Status, Stage
- `driverName`: Driver, Driver Name
- `driverPhone`: Phone, Driver Phone, Contact
- `unitNumber`: Unit, Truck, Truck #, Tractor, Unit Number
- `broker`: Broker, Customer, Shipper

If multiple candidates match, the UI asks for confirmation.


## Validation Rules (high‑level)

- Required columns present (all but `driverPhone`).
- Empty required cells flagged per row.
- Duplicate `loadId` errors.
- Date/time must be parseable, include explicit timezone, and be normalized to ISO 8601 UTC.
- Status vocabulary surfaced for normalization (warning).
- Soft rate limit: 10 analyses/min/user.


## Auto‑Fixes (Optional)

Plan and apply non‑destructive fixes:
- Create missing required columns.
- Normalize `fromAppointmentDateTimeUTC` / `toAppointmentDateTimeUTC` into dedicated columns when safely derivable (from split date/time or explicit‑TZ values). Skips “Excel time‑only” placeholders.
- Optional offset control in UI for combining split date/time during fix.

All fixes require explicit user action from the sidebar.


## OpenAI Proxy (Optional)

Set `PROXY_URL` to any trusted endpoint that forwards your structured request to OpenAI and returns a structured response. Keep your OpenAI key server‑side. The add‑on never requires embedding API keys in code.


## Testing

- Sample spreadsheet: see codingtask.md for the shared example with happy/broken rows.
- Unit tests: In Apps Script editor, run `runUnitTests()` (from `tests.gs`) and check logs.
- Manual flow: Open sidebar → Analyze → Review issues → Fix in sheet → Re‑analyze → Copy JSON.


## Troubleshooting

- Add‑on not visible: Ensure you installed the test deployment and are opening in Google Sheets (not Docs/Slides).
- External requests blocked: You must include the `script.external_request` scope (already in manifest) and set `PROXY_URL`.
- AI features missing: Check `PROXY_URL` and your proxy health/logs. The add‑on works without AI.
- Timezone errors: The add‑on treats missing explicit TZ as an error by default. For demos, you can set `ASSUME_TZ_AS_WARN=true` (Script property) to downgrade to warning.
- Too many requests: Soft rate limit is 10 analyses per minute per user.
- Large sheets: The snapshot reads up to 200 data rows for performance; adjust in code if needed.


## Limitations

- No destructive edits by default; auto‑fixes require explicit confirmation.
- Only the active sheet tab is analyzed.
- Assumes one row = one load.
- AI never fabricates values; unknowns remain empty and flagged.


## License

MIT (see package.json). If redistributing, review scopes and security guidance for your environment.
