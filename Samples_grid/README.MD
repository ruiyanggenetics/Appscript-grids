# Samples_Grid (Human Samples)

Web-based freezer box grid UI for the **SAMPLES** table (human samples only), backed by Google Sheets + Apps Script.

This project renders a box grid (9x9 / 9x12 / 8x12), enforces an **allowlist**, limits access to **Human samples** boxes only, and supports placing/moving/clearing vials plus batch status updates (including “Used-up” removal logic).

---

## Data model assumptions

### BOXES sheet
Sheet name: `BOXES`

Required columns:
- `BoxID` — human-readable label (e.g. `MINUS80_RACK1_PLASMABOX1`)
- `BOX_UID` — key/UID used as the reference target
- `BoxType` — `9x9`, `9x12`, or `8x12`
- `BoxSampleType` — must equal **`Human samples`** for this grid to allow it

### SAMPLES sheet
Sheet name: `SAMPLES`

Assumptions:
- `SAMPLES[BoxID]` stores **BOX_UID** (a Ref to `BOXES[BOX_UID]`), **not** the human BoxID label.

Required columns:
- `VIAL_Key`
- `BoxID` (stores BOX_UID)
- `Position` (e.g. `A01`)
- `Position_Key`
- `Status`
- `StatusChangedAt` *(optional but supported; required in code as written)*
- `StatusChangedBy` *(optional but supported; required in code as written)*

### Allowlist sheet
Sheet name: `GRID_ALLOWED_USERS`

Required columns:
- `Email`
- `ACTIVE` (`TRUE/YES/1` are treated as true)

If `REQUIRE_ALLOWLIST_FOR_READ` is `true`, **all reads and writes** (including `doGet`) require allowlist authorization.

---

## Security & restrictions

### Allowlist enforcement
- **Reads + writes are blocked** unless the caller’s email is in `GRID_ALLOWED_USERS` and `ACTIVE` is true.
- Caller identity comes from `Session.getActiveUser().getEmail()`.

### Human sample boxes only
Only boxes where:
- `BOXES[BoxSampleType] == "Human samples"`
are listed/opened. Attempting to open other boxes throws an error.

---

## Features

### UI actions
✅ Place vial (select an unplaced vial → click an empty well)  
✅ Move vial (Move mode: click occupied source → click empty target)  
✅ Clear well (Clear mode: click occupied well to clear Position/Position_Key)  
✅ Select wells (click/drag occupied wells)  
✅ Batch status apply (Apply to Selected)  
✅ Compact (stable compaction moves vials to earliest wells)  
✅ Auto-refresh every 2 seconds when idle

### Status “Used-up” handling (important)
When `Status` becomes `Used-up` (from grid, AppSheet edit, or sheet edit via trigger):

- `Position` is set to: `removed`
- `Position_Key` is set to: `<VIAL_Key>_<OLD Position_Key>_removed`

If old `Position_Key` is blank, the script attempts to reconstruct it as:
- `<HumanBoxID>_<OLD Position>`

This makes the well disappear from the grid because occupancy only includes `A01`-style positions.

---

## URL parameters (Web App)

Supported params:

- `table=SAMPLES` *(optional; ignored but tolerated)*
- `boxUid=<BOX_UID>` *(preferred)*
- `boxId=<BoxID>` *(compat; mapped to BOX_UID via BOXES)*
- `pos=A01` *(optional highlight in UI)*
- `vialKey=<VIAL_Key>` *(optional; also accepts `itemKey`)*

Robust resolution logic:
1. Use `boxUid` if present
2. Else convert `boxId` → `boxUid`
3. If `boxUid` doesn’t match a known UID, treat it as a BoxID and try converting

---

## Sheets UI (bound script)

If this Apps Script is **bound** to the spreadsheet:
- A custom menu is created on open: **Freezer Grid**
- You can open the grid for the selected row in `SAMPLES`

Menu items:
- Open grid for selected sample (SAMPLES)
- Open grid (last box)
- Rebuild caches (optional)
- Install protections (optional)

---

## Triggers

### Installable onEdit trigger (required for AppSheet edits)
To make **Status → Used-up** logic work reliably when AppSheet edits the sheet:

Apps Script → **Triggers** → **Add Trigger**
- Function: `onEdit`
- Event source: From spreadsheet
- Event type: On edit

This trigger watches edits to the `Status` column and applies the removal mapping.

---

## Caching

- BOX maps cached for 60 seconds (`CACHE_TTL_SECONDS`)
- Per-box occupancy cached for 60 seconds (`OCC:<boxUid>`)

Use “Rebuild caches” in the menu if needed.

---

## Troubleshooting

### “Not authorized”
Add your email to `GRID_ALLOWED_USERS` and set `ACTIVE` to TRUE.

### “Restricted to Human samples boxes”
Check `BOXES[BoxSampleType]` is exactly `Human samples`.

### Box not opening from AppSheet link
Make sure your AppSheet action passes the **BOX_UID** to `boxUid`:

```appsheet
CONCATENATE(
  "YOUR_WEBAPP_EXEC_URL",
  "?boxUid=", ENCODEURL([BoxID]),
  "&pos=", ENCODEURL([Position]),
  "&vialKey=", ENCODEURL([VIAL_Key])
)
