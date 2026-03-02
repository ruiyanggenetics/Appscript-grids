# glycerol_grid

Web-based freezer box grid for **bacterial glycerol stock** inventory.

This project renders a 9x9 (or 8x12) box grid, lets allowed users **place**, **move**, and **delete** glycerol records, and auto-refreshes every 2 seconds.

---

## What it works with

### Google Sheet
Spreadsheet:
- `1soWIXL1usgLSI5YyHnw8gDvxf9i45PgAcbwTQFijc0I`

Required sheets + columns:

#### `BOXES`
Must include:
- `BoxID` (human-readable label, e.g. `MINUS80_RACK1_GLYCEROLBOX1`)
- `BOX_UID` (key)
- `BoxType` (e.g. `9x9`, `8x12`)
- `BoxSampleType` (must equal **`Bacterial glycerol stock`** for this grid)

#### `PLASMID_GLYCEROL`
Must include:
- `PG_UID` (unique key)
- `Plasmid Name` (display label)
- `BoxID` (stores the **BOX_UID**, not the human BoxID)
- `Position` (e.g. `A01`)
- `Position_Key` (e.g. `MINUS80_RACK1_GLYCEROLBOX1_A01`)

#### `GRID_ALLOWED_USERS`
Must include:
- `Email`
- `ACTIVE` (true/yes/1)

If `REQUIRE_ALLOWLIST_FOR_READ` is enabled, *all reads and writes* require the user to be in this allowlist.

---

## Security / Restrictions

### Allowlist
Only users listed in `GRID_ALLOWED_USERS` with `ACTIVE = true` are authorized.

### Box filter (critical)
This app will only open/list boxes where:
- `BOXES[BoxSampleType] == "Bacterial glycerol stock"`

Users cannot browse or open unrelated boxes (e.g., Human samples, Cell line).

---

## Features

### Supported actions
✅ Place (select an unplaced item, click an empty well)  
✅ Move (move mode: click source occupied well, then empty target)  
✅ Delete (deletes entire row in `PLASMID_GLYCEROL`)  
✅ Auto-refresh every 2 seconds

### Not included
- No status fields
- No “Used-up” logic
- No partial deletion (delete always removes the full row)

---

## Web App parameters

The deployed web app supports:

- `?boxUid=<BOX_UID>` (preferred)
- `?boxId=<human BoxID>` (compat; will be mapped to BOX_UID)
- `?pos=A01` (optional)

Example:
