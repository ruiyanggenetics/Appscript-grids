/*********************************
 * Code.gs — SAMPLES ONLY GRID (BOX_UID-based) + ALLOWLIST + HUMAN-BOX FILTER
 *
 * ✅ Data assumptions (your current model):
 * - BOXES has: BoxID (human), BOX_UID (key), BoxType, BoxSampleType
 * - SAMPLES[BoxID] stores BOX_UID (Ref to BOXES)
 *
 * ✅ Web app supports:
 *   ?table=SAMPLES (optional; ignored but tolerated)
 *   ?boxUid=<BOX_UID>   (preferred)
 *   ?boxId=<BoxID>      (compat)
 *   ?vialKey=<VIAL_Key> (optional; also accepts itemKey)
 *   ?pos=A01            (optional)
 *
 * ✅ Allowlist enforced for ALL reads/writes:
 * - Sheet: GRID_ALLOWED_USERS
 * - Columns: Email | User | ACTIVE
 *
 * ✅ Human-only boxes:
 * - Only BOXES[BoxSampleType] == "Human samples"
 *
 * ✅ Used-up rule (exact):
 * - Position = "removed"
 * - Position_Key = <VIAL_Key>_<OLD Position_Key>_removed
 *   If OLD Position_Key blank, try compute <HumanBoxID>_<OLD Position>
 *********************************/

/************ CONFIG ************/
const CONFIG = {
  SAMPLES_SHEET: 'SAMPLES',
  BOXES_SHEET: 'BOXES',
  ALLOWED_USERS_SHEET: 'GRID_ALLOWED_USERS',

  REQUIRE_ALLOWLIST_FOR_READ: true,
  HUMAN_BOX_SAMPLETYPE: 'Human samples',

  // SAMPLES columns
  COLS: {
    VIALKEY: 'VIAL_Key',
    BOXID: 'BoxID',               // stores BOX_UID
    POSITION: 'Position',
    POSITION_KEY: 'Position_Key',
    STATUS: 'Status',
    STATUS_CHANGED_AT: 'StatusChangedAt',
    STATUS_CHANGED_BY: 'StatusChangedBy'
  },

  // BOXES columns
  BOX_COLS: {
    BOXID: 'BoxID',               // human label
    BOXUID: 'BOX_UID',            // key
    BOXTYPE: 'BoxType',
    BOXSAMPLETYPE: 'BoxSampleType'
  },

  // BoxType specs
  BOXTYPE_SPECS: {
    '9x9':  { rows: 'ABCDEFGHI'.split(''), cols: 9  },
    '9x12': { rows: 'ABCDEFGH'.split(''), cols: 12 },
    '8x12': { rows: 'ABCDEFGH'.split(''), cols: 12 }
  },

  DIALOG_WIDTH: 1400,
  DIALOG_HEIGHT: 950,
  CACHE_TTL_SECONDS: 60
};

const USED_UP = 'Used-up';

/************ SPREADSHEET ACCESS (webapp-safe) ************/
const SPREADSHEET_ID = 'XXXXXXXXXXX';

function getSS_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function readSheet_(name) {
  const ss = getSS_();
  const sh = ss.getSheetByName(name);
  if (!sh) throw new Error(`Sheet not found: ${name}`);
  const values = sh.getDataRange().getValues();
  return { sh, header: values[0] || [], values };
}

/************ COMPAT WRAPPERS ************/
function openGridDialogForSelectedRow() { openGridForSelectedRow(); }
function openGridDialogForSelectedRowDetails() { openGridForSelectedRow(); }

/************ ALLOWLIST ************/
function getCallerEmail_() {
  return String(Session.getActiveUser().getEmail() || '').toLowerCase().trim();
}
function truthy_(v) {
  const s = String(v ?? '').toLowerCase().trim();
  return s === 'true' || s === 'yes' || s === 'y' || s === '1';
}
function isAllowedUser_() {
  const email = getCallerEmail_();
  if (!email) return false;

  const { values } = readSheet_(CONFIG.ALLOWED_USERS_SHEET);
  if (!values || values.length < 2) return false;

  const header = values[0].map(h => String(h || '').trim());
  const emailIdx = header.indexOf('Email');
  const activeIdx = header.indexOf('ACTIVE');
  if (emailIdx < 0 || activeIdx < 0) {
    throw new Error(`Missing required columns in ${CONFIG.ALLOWED_USERS_SHEET}: Email, ACTIVE`);
  }

  for (let r = 1; r < values.length; r++) {
    const rowEmail = String(values[r][emailIdx] || '').toLowerCase().trim();
    if (!rowEmail) continue;
    if (rowEmail === email) return truthy_(values[r][activeIdx]);
  }
  return false;
}
function assertAllowedUser_(opName) {
  if (!isAllowedUser_()) {
    const email = getCallerEmail_() || '(email unavailable)';
    throw new Error(`Not authorized for ${opName || 'operation'}: ${email}`);
  }
}
function assertAllowedUserForReadIfEnabled_() {
  if (!CONFIG.REQUIRE_ALLOWLIST_FOR_READ) return;
  assertAllowedUser_('read');
}

/************ MENU (Sheets UI) ************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Freezer Grid')
    .addItem('Open grid for selected sample (SAMPLES)', 'openGridForSelectedRow')
    .addItem('Open grid (last box)', 'openGridLastBox')
    .addSeparator()
    .addItem('Rebuild caches (optional)', 'clearGridCaches')
    .addSeparator()
    .addItem('Install protections (optional)', 'installProtections_NoDeleteSamples_')
    .addToUi();
}

/************ SHEETS UI ENTRY (SAMPLES ONLY) ************/
function openGridForSelectedRow() {
  assertAllowedUser_('open grid');

  const ctx = getContextFromActiveRow_(); // must be on SAMPLES sheet
  PropertiesService.getScriptProperties().setProperties({
    LAST_BOXUID: ctx.boxUid,
    LAST_VIALKEY: ctx.vialKey || '',
    LAST_POSITION: ctx.position || ''
  });
  showGridDialog_(ctx.boxUid, ctx.vialKey || '', ctx.position || '');
}

function openGridLastBox() {
  assertAllowedUser_('open grid');

  const sp = PropertiesService.getScriptProperties();
  const boxUid = sp.getProperty('LAST_BOXUID') || '';
  const vialKey = sp.getProperty('LAST_VIALKEY') || '';
  const position = sp.getProperty('LAST_POSITION') || '';
  if (!boxUid) {
    SpreadsheetApp.getUi().alert('No last box stored. Click a row in SAMPLES and open once.');
    return;
  }
  showGridDialog_(boxUid, vialKey, position);
}

function clearGridCaches() {
  CacheService.getScriptCache().removeAll(['BOX_MAP_UID', 'BOX_MAP_ID']);
  SpreadsheetApp.getActive().toast('Cache cleared.', 'Freezer Grid', 4);
}

/************ DIALOG (Sheets UI) ************/
function showGridDialog_(boxUid, selectedVialKey, selectedPos) {
  assertHumanSampleBoxUid_(boxUid);

  const boxInfo = getBoxInfoByUid_(boxUid); // {boxUid, boxId, boxType, boxSampleType}
  const spec = CONFIG.BOXTYPE_SPECS[boxInfo.boxType];
  if (!spec) throw new Error(`Unsupported BoxType: ${boxInfo.boxType}`);

  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.mode = 'grid';
  tpl.boxUid = boxUid;
  tpl.boxId = boxInfo.boxId;
  tpl.boxType = boxInfo.boxType;
  tpl.selectedVialKey = selectedVialKey || '';
  tpl.selectedPos = selectedPos || '';

  const html = tpl.evaluate()
    .setTitle('Box Grid')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setWidth(CONFIG.DIALOG_WIDTH)
    .setHeight(CONFIG.DIALOG_HEIGHT);

  SpreadsheetApp.getUi().showModalDialog(html, `Grid: ${boxInfo.boxId} (${boxInfo.boxType})`);
}

/************ WEB APP ENTRY ************/
function doGet(e) {
  assertAllowedUserForReadIfEnabled_();

  const rawBoxUid = String(e?.parameter?.boxUid || '').trim();
  const rawBoxId  = String(e?.parameter?.boxId  || '').trim();
  const vialKey   = String(e?.parameter?.vialKey || e?.parameter?.itemKey || '').trim();
  const selectedPos = String(e?.parameter?.pos || '').trim().toUpperCase();

  // Robust resolution:
  // 1) Prefer boxUid param
  // 2) Else use boxId param -> lookup uid
  // 3) Else if boxUid param was actually a human BoxID, try lookup as BoxID
  let boxUid = rawBoxUid;

  if (!boxUid && rawBoxId) {
    boxUid = lookupBoxUidByBoxId_(rawBoxId);
  }

  // If boxUid provided but not found as a UID, treat it as a BoxID and try converting.
  if (boxUid) {
    try {
      // will throw if not a known uid
      getBoxInfoByUid_(boxUid);
    } catch (err) {
      const maybeUid = lookupBoxUidByBoxId_(boxUid);
      if (maybeUid) boxUid = maybeUid;
    }
  }

  // Landing page
  if (!boxUid) {
    const tpl = HtmlService.createTemplateFromFile('Index');
    tpl.mode = 'browser';
    tpl.boxUid = '';
    tpl.boxId = '';
    tpl.boxType = '';
    tpl.selectedVialKey = vialKey || '';
    tpl.selectedPos = selectedPos || '';
    return tpl.evaluate()
      .setTitle('Box Browser')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // Enforce human boxes only
  assertHumanSampleBoxUid_(boxUid);

  const boxInfo = getBoxInfoByUid_(boxUid);
  const tpl = HtmlService.createTemplateFromFile('Index');
  tpl.mode = 'grid';
  tpl.boxUid = boxUid;
  tpl.boxId = boxInfo.boxId;
  tpl.boxType = boxInfo.boxType;
  tpl.selectedVialKey = vialKey || '';
  tpl.selectedPos = selectedPos || '';

  return tpl.evaluate()
    .setTitle(`Grid: ${boxInfo.boxId} (${boxInfo.boxType})`)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/************ API: BOX LIST (HUMAN ONLY) ************/
function apiListBoxes() {
  assertAllowedUserForReadIfEnabled_();

  const mapUid = buildBoxMapByUid_(); // uid -> info
  const out = [];
  for (const uid of Object.keys(mapUid)) {
    const b = mapUid[uid];
    if (String(b.boxSampleType || '').trim() !== CONFIG.HUMAN_BOX_SAMPLETYPE) continue;
    out.push({ boxUid: uid, boxId: b.boxId, boxType: b.boxType });
  }
  out.sort((a,b)=>String(a.boxId).localeCompare(String(b.boxId)));
  return out;
}

/************ API: GRID STATE ************/
function apiGetGridState(boxUid) {
  assertAllowedUserForReadIfEnabled_();
  assertHumanSampleBoxUid_(boxUid);

  const boxInfo = getBoxInfoByUid_(boxUid);
  const spec = CONFIG.BOXTYPE_SPECS[boxInfo.boxType];
  if (!spec) throw new Error(`Unsupported BoxType: ${boxInfo.boxType}`);

  const cacheKey = `OCC:${boxUid}`;
  const cache = CacheService.getScriptCache();
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const occ = getOccupancyForBoxUid_(boxUid);
  const labels = generateLabels_(spec.rows, spec.cols);

  const cells = labels.map(label => ({
    label,
    occupied: occ.has(label),
    vialKey: occ.has(label) ? occ.get(label).vialKey : '',
    status: occ.has(label) ? (occ.get(label).status || '') : ''
  }));

  const state = { boxUid, boxId: boxInfo.boxId, boxType: boxInfo.boxType, cells };
  cache.put(cacheKey, JSON.stringify(state), CONFIG.CACHE_TTL_SECONDS);
  return state;
}

function apiListUnplacedVials(boxUid) {
  assertAllowedUserForReadIfEnabled_();
  assertHumanSampleBoxUid_(boxUid);

  const { header, values } = readSheet_(CONFIG.SAMPLES_SHEET);
  const idx = colIndex_(header);

  const out = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (String(row[idx.BoxID] || '').trim() !== boxUid) continue;

    const vialKey = String(row[idx.Vial_Key] || '').trim();
    const pos = String(row[idx.Position] || '').trim();
    if (!vialKey) continue;
    if (pos) continue;
    out.push(vialKey);
  }
  out.sort();
  return out;
}

/************ API: PLACE / MOVE / CLEAR / STATUS / COMPACT ************/
function apiPlaceVial(boxUid, vialKey, targetPos) {
  assertAllowedUser_('place vial');
  assertHumanSampleBoxUid_(boxUid);
  if (!vialKey) throw new Error('Missing vialKey');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const boxInfo = getBoxInfoByUid_(boxUid);
    validatePos_(boxInfo.boxType, targetPos);
    const pos = normPos_(targetPos);

    const occ = getOccupancyForBoxUid_(boxUid);
    if (occ.has(pos)) throw new Error(`Target ${pos} occupied by ${occ.get(pos).vialKey}`);

    const rowIndex = findRowIndexByKey_(CONFIG.SAMPLES_SHEET, CONFIG.COLS.VIALKEY, vialKey);
    if (!rowIndex) throw new Error(`VIAL_Key not found: ${vialKey}`);

    const { sh, header } = readSheet_(CONFIG.SAMPLES_SHEET);
    const idx = colIndex_(header);

    const rowVals = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
    const rowBoxUid = String(rowVals[idx.BoxID] || '').trim();
    if (rowBoxUid !== boxUid) throw new Error(`Vial belongs to different BOX_UID=${rowBoxUid}`);

    const currentPos = String(rowVals[idx.Position] || '').trim();
    if (currentPos) throw new Error(`Vial already placed at ${currentPos}`);

    sh.getRange(rowIndex, idx.Position + 1).setValue(pos);
    sh.getRange(rowIndex, idx.Position_Key + 1).setValue(`${boxInfo.boxId}_${pos}`);

    invalidateBoxCache_(boxUid);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function apiMoveVial(boxUid, fromPos, toPos) {
  assertAllowedUser_('move vial');
  assertHumanSampleBoxUid_(boxUid);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const boxInfo = getBoxInfoByUid_(boxUid);
    validatePos_(boxInfo.boxType, fromPos);
    validatePos_(boxInfo.boxType, toPos);

    const fromP = normPos_(fromPos);
    const toP = normPos_(toPos);

    const occ = getOccupancyForBoxUid_(boxUid);
    if (!occ.has(fromP)) throw new Error(`Source ${fromP} empty`);
    if (occ.has(toP)) throw new Error(`Target ${toP} occupied`);

    const rowIndex = occ.get(fromP).rowIndex;

    const { sh, header } = readSheet_(CONFIG.SAMPLES_SHEET);
    const idx = colIndex_(header);

    sh.getRange(rowIndex, idx.Position + 1).setValue(toP);
    sh.getRange(rowIndex, idx.Position_Key + 1).setValue(`${boxInfo.boxId}_${toP}`);

    invalidateBoxCache_(boxUid);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function apiClearWell(boxUid, pos) {
  assertAllowedUser_('clear well');
  assertHumanSampleBoxUid_(boxUid);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const boxInfo = getBoxInfoByUid_(boxUid);
    validatePos_(boxInfo.boxType, pos);

    const p = normPos_(pos);
    const occ = getOccupancyForBoxUid_(boxUid);
    if (!occ.has(p)) return { ok: true, action: 'noop' };

    const rowIndex = occ.get(p).rowIndex;

    const { sh, header } = readSheet_(CONFIG.SAMPLES_SHEET);
    const idx = colIndex_(header);

    sh.getRange(rowIndex, idx.Position + 1).setValue('');
    sh.getRange(rowIndex, idx.Position_Key + 1).setValue('');

    invalidateBoxCache_(boxUid);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function apiBatchSetStatus(boxUid, positions, newStatus) {
  assertAllowedUser_('set status');
  assertHumanSampleBoxUid_(boxUid);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    if (!Array.isArray(positions) || positions.length === 0) return { ok: true, updated: 0 };

    const boxInfo = getBoxInfoByUid_(boxUid);
    const occ = getOccupancyForBoxUid_(boxUid);

    let updated = 0;
    for (const rawPos of positions) {
      validatePos_(boxInfo.boxType, rawPos);
      const p = normPos_(rawPos);
      if (!occ.has(p)) continue;
      setStatusOnRow_(occ.get(p).rowIndex, newStatus);
      updated++;
    }

    invalidateBoxCache_(boxUid);
    return { ok: true, updated };
  } finally {
    lock.releaseLock();
  }
}

function apiCompactBox(boxUid) {
  assertAllowedUser_('compact box');
  assertHumanSampleBoxUid_(boxUid);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const boxInfo = getBoxInfoByUid_(boxUid);
    const spec = CONFIG.BOXTYPE_SPECS[boxInfo.boxType];
    if (!spec) throw new Error(`Unsupported BoxType: ${boxInfo.boxType}`);

    const labels = generateLabels_(spec.rows, spec.cols);
    const occ = getOccupancyForBoxUid_(boxUid);

    const occupied = [];
    for (const lab of labels) if (occ.has(lab)) occupied.push({ from: lab, rowIndex: occ.get(lab).rowIndex });

    const moves = [];
    for (let i = 0; i < occupied.length; i++) {
      const to = labels[i];
      if (occupied[i].from !== to) moves.push({ rowIndex: occupied[i].rowIndex, to });
    }
    if (!moves.length) return { ok: true, moved: 0 };

    const { sh, header } = readSheet_(CONFIG.SAMPLES_SHEET);
    const idx = colIndex_(header);

    for (const m of moves) {
      sh.getRange(m.rowIndex, idx.Position + 1).setValue(m.to);
      sh.getRange(m.rowIndex, idx.Position_Key + 1).setValue(`${boxInfo.boxId}_${m.to}`);
    }

    invalidateBoxCache_(boxUid);
    return { ok: true, moved: moves.length };
  } finally {
    lock.releaseLock();
  }
}

/************ INSTALLABLE TRIGGER: onEdit (Status->Used-up) ************/
function onEdit(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== CONFIG.SAMPLES_SHEET) return;

    const row = e.range.getRow();
    if (row <= 1) return;

    const lastCol = sh.getLastColumn();
    const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
    const idx = colIndex_(header);

    if (e.range.getColumn() !== (idx.Status + 1)) return;

    const newStatus = String(e.value || '').trim();
    if (newStatus !== USED_UP) return;

    applyUsedUpRemovalOnRow_(sh, row, idx);

    const rowVals2 = sh.getRange(row, 1, 1, lastCol).getValues()[0];
    const boxUid = String(rowVals2[idx.BoxID] || '').trim();
    if (boxUid) invalidateBoxCache_(boxUid);
  } catch (err) {
    console.error(err);
  }
}

/************ STATUS HELPERS ************/
function setStatusOnRow_(rowIndex, newStatus) {
  const { sh, header } = readSheet_(CONFIG.SAMPLES_SHEET);
  const idx = colIndex_(header);

  const status = String(newStatus || '').trim();
  sh.getRange(rowIndex, idx.Status + 1).setValue(status);

  if (idx.StatusChangedAt >= 0) sh.getRange(rowIndex, idx.StatusChangedAt + 1).setValue(new Date());
  if (idx.StatusChangedBy >= 0) sh.getRange(rowIndex, idx.StatusChangedBy + 1).setValue(getCallerEmail_() || '');

  if (status === USED_UP) applyUsedUpRemovalOnRow_(sh, rowIndex, idx);
}

function applyUsedUpRemovalOnRow_(sh, rowIndex, idx) {
  const rowVals = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];

  const vialKey = String(rowVals[idx.Vial_Key] || '').trim();
  const boxUid = String(rowVals[idx.BoxID] || '').trim();

  const oldPos = String(rowVals[idx.Position] || '').trim();
  let oldPK = String(rowVals[idx.Position_Key] || '').trim();

  if (!oldPK && boxUid && oldPos && oldPos.length === 3) {
    const b = getBoxInfoByUid_(boxUid);
    oldPK = `${b.boxId}_${oldPos.toUpperCase()}`;
  }
  if (!oldPK) oldPK = 'UNKNOWN';

  const targetPK = `${vialKey}_${oldPK}_removed`;

  if (String(rowVals[idx.Position] || '').trim() === 'removed' &&
      String(rowVals[idx.Position_Key] || '').trim() === targetPK) {
    return;
  }

  sh.getRange(rowIndex, idx.Position + 1).setValue('removed');
  sh.getRange(rowIndex, idx.Position_Key + 1).setValue(targetPK);

  if (idx.StatusChangedAt >= 0) sh.getRange(rowIndex, idx.StatusChangedAt + 1).setValue(new Date());
  if (idx.StatusChangedBy >= 0) sh.getRange(rowIndex, idx.StatusChangedBy + 1).setValue(getCallerEmail_() || '');
}

/************ INTERNAL HELPERS ************/
function invalidateBoxCache_(boxUid) {
  CacheService.getScriptCache().remove(`OCC:${boxUid}`);
}

function colIndex_(header) {
  function must(colName) {
    const i = header.indexOf(colName);
    if (i < 0) throw new Error(`Missing required column: ${colName}`);
    return i;
  }
  const out = {
    Vial_Key: must(CONFIG.COLS.VIALKEY),
    BoxID: must(CONFIG.COLS.BOXID),
    Position: must(CONFIG.COLS.POSITION),
    Position_Key: must(CONFIG.COLS.POSITION_KEY),
    Status: header.indexOf(CONFIG.COLS.STATUS),
    StatusChangedAt: header.indexOf(CONFIG.COLS.STATUS_CHANGED_AT),
    StatusChangedBy: header.indexOf(CONFIG.COLS.STATUS_CHANGED_BY)
  };
  if (out.Status < 0) throw new Error(`Missing required column in SAMPLES: ${CONFIG.COLS.STATUS}`);
  return out;
}

/************ BOX MAPS ************/
function buildBoxMapByUid_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('BOX_MAP_UID');
  if (cached) return JSON.parse(cached);

  const { values } = readSheet_(CONFIG.BOXES_SHEET);
  if (!values || values.length < 2) return {};

  const header = values[0].map(h => String(h || '').trim());
  const boxIdIdx = header.indexOf(CONFIG.BOX_COLS.BOXID);
  const boxUidIdx = header.indexOf(CONFIG.BOX_COLS.BOXUID);
  const boxTypeIdx = header.indexOf(CONFIG.BOX_COLS.BOXTYPE);
  const boxSampleIdx = header.indexOf(CONFIG.BOX_COLS.BOXSAMPLETYPE);

  if (boxIdIdx < 0 || boxUidIdx < 0 || boxTypeIdx < 0 || boxSampleIdx < 0) {
    throw new Error('BOXES must have: BoxID, BOX_UID, BoxType, BoxSampleType');
  }

  const m = {};
  for (let r = 1; r < values.length; r++) {
    const boxId = String(values[r][boxIdIdx] || '').trim();
    const boxUid = String(values[r][boxUidIdx] || '').trim();
    const boxType = String(values[r][boxTypeIdx] || '').trim();
    const boxSampleType = String(values[r][boxSampleIdx] || '').trim();
    if (!boxUid || !boxType) continue;
    m[boxUid] = { boxId, boxType, boxSampleType };
  }

  cache.put('BOX_MAP_UID', JSON.stringify(m), CONFIG.CACHE_TTL_SECONDS);
  return m;
}

function buildBoxMapById_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('BOX_MAP_ID');
  if (cached) return JSON.parse(cached);

  const mapUid = buildBoxMapByUid_();
  const m = {};
  for (const uid of Object.keys(mapUid)) {
    const b = mapUid[uid];
    if (b.boxId) m[b.boxId] = uid;
  }
  cache.put('BOX_MAP_ID', JSON.stringify(m), CONFIG.CACHE_TTL_SECONDS);
  return m;
}

function lookupBoxUidByBoxId_(boxId) {
  const m = buildBoxMapById_();
  return m[String(boxId || '').trim()] || '';
}

function getBoxInfoByUid_(boxUid) {
  const m = buildBoxMapByUid_();
  const info = m[String(boxUid || '').trim()];
  if (!info) throw new Error(`BOX_UID not found in BOXES: ${boxUid}`);
  return { boxUid: String(boxUid || '').trim(), boxId: info.boxId, boxType: info.boxType, boxSampleType: info.boxSampleType };
}

function assertHumanSampleBoxUid_(boxUid) {
  const b = getBoxInfoByUid_(boxUid);
  const actual = String(b.boxSampleType || '').trim();
  if (actual !== CONFIG.HUMAN_BOX_SAMPLETYPE) {
    throw new Error(`Restricted to "${CONFIG.HUMAN_BOX_SAMPLETYPE}" boxes. BoxSampleType="${actual}".`);
  }
}

/************ OCCUPANCY ************/
function getOccupancyForBoxUid_(boxUid) {
  const ss = getSS_();
  const sh = ss.getSheetByName(CONFIG.SAMPLES_SHEET);
  if (!sh) throw new Error(`Sheet not found: ${CONFIG.SAMPLES_SHEET}`);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return new Map();

  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0];
  const idx = colIndex_(header);
  const values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const occ = new Map(); // pos -> {vialKey,rowIndex,status}
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    if (String(row[idx.BoxID] || '').trim() !== boxUid) continue;

    const vialKey = String(row[idx.Vial_Key] || '').trim();
    const pos = String(row[idx.Position] || '').trim().toUpperCase();
    if (!vialKey || !pos) continue;
    if (pos.length !== 3) continue; // ignores removed

    const status = String(row[idx.Status] || '').trim();
    if (!occ.has(pos)) occ.set(pos, { vialKey, rowIndex: i + 2, status });
  }
  return occ;
}

/************ FIND ROW BY KEY ************/
function findRowIndexByKey_(sheetName, keyColName, keyValue) {
  const ss = getSS_();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Sheet not found: ${sheetName}`);

  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow < 2) return 0;

  const header = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h || '').trim());
  const keyIdx = header.indexOf(keyColName);
  if (keyIdx < 0) throw new Error(`Missing column: ${keyColName}`);

  const keys = sh.getRange(2, keyIdx + 1, lastRow - 1, 1).getValues();
  for (let i = 0; i < keys.length; i++) {
    if (String(keys[i][0] || '').trim() === String(keyValue || '').trim()) return i + 2;
  }
  return 0;
}

/************ POSITION UTILS ************/
function normPos_(pos) { return String(pos || '').trim().toUpperCase(); }
function validatePos_(boxType, pos) {
  const spec = CONFIG.BOXTYPE_SPECS[String(boxType || '').trim()];
  if (!spec) throw new Error(`Unsupported BoxType: ${boxType}`);

  const p = normPos_(pos);
  if (p.length !== 3) throw new Error(`Invalid position format: ${p} (expected A01)`);

  const row = p.substring(0, 1);
  const col = p.substring(1, 3);

  if (spec.rows.indexOf(row) < 0) throw new Error(`Invalid row ${row} for box type ${boxType}`);

  const colNum = Number(col);
  if (!colNum || colNum < 1 || colNum > spec.cols) throw new Error(`Invalid column ${col} for box type ${boxType}`);
}
function generateLabels_(rows, colsCount) {
  const labels = [];
  for (const r of rows) for (let c = 1; c <= colsCount; c++) labels.push(`${r}${String(c).padStart(2, '0')}`);
  return labels;
}

/************ SHEETS UI CONTEXT ************/
function apiGetContextFromActiveRow() { return getContextFromActiveRow_(); }

function getContextFromActiveRow_() {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getActiveSheet();
  if (sh.getName() !== CONFIG.SAMPLES_SHEET) {
    throw new Error(`Go to sheet "${CONFIG.SAMPLES_SHEET}", select a data row, then try again.`);
  }
  const r = sh.getActiveRange();
  if (!r) throw new Error('No active cell selected.');
  const row = r.getRow();
  if (row <= 1) throw new Error('Select a data row (not header).');

  const header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const idx = colIndex_(header);

  const vals = sh.getRange(row, 1, 1, sh.getLastColumn()).getValues()[0];
  const vialKey = String(vals[idx.Vial_Key] || '').trim();
  const boxUid = String(vals[idx.BoxID] || '').trim();
  const position = String(vals[idx.Position] || '').trim().toUpperCase();

  if (!boxUid) throw new Error('Selected row has blank BoxID (BOX_UID ref).');
  assertHumanSampleBoxUid_(boxUid);

  return { vialKey: vialKey || '', boxUid, position: position || '' };
}

/************ OPTIONAL: DISCOURAGE DELETION IN SHEETS ************/
function installProtections_NoDeleteSamples_() {
  const ss = getSS_();
  const sh = ss.getSheetByName(CONFIG.SAMPLES_SHEET);
  if (!sh) throw new Error(`Sheet not found: ${CONFIG.SAMPLES_SHEET}`);

  const protections = sh.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  protections.forEach(p => { if (p.getDescription() === 'NO_DELETE_SAMPLES_SHEET') p.remove(); });

  const p = sh.protect();
  p.setDescription('NO_DELETE_SAMPLES_SHEET');
  p.setWarningOnly(false);

  const dataRange = sh.getDataRange();
  p.setUnprotectedRanges([dataRange]);

  SpreadsheetApp.getActive().toast(
    'Installed SAMPLES sheet protection (optional). For AppSheet: set “Adds and updates” to fully prevent deletes.',
    'Freezer Grid',
    6
  );
}
