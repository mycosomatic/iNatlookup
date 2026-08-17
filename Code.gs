/**
 * iNaturalist metadata lookup for Google Sheets.
 *
 * Watches recognized lookup columns ("id" plus selected observation fields)
 * and fills selected return columns with metadata from the iNaturalist API.
 * See README.md in this folder for setup and behavior notes.
 */

// === CONFIG ===
const MENU_NAME = 'iNaturalist';
const LOG_SHEET_NAME = '_iNatLog';
const REQUEST_DELAY_MS = 200;        // pause between consecutive API requests
const PROP_RETURN_FIELDS = 'RETURN_FIELDS_JSON';
const PROP_AUTO_LOOKUP = 'AUTO_LOOKUP_ENABLED';
const PROP_PENDING_JOBS = 'PENDING_LOOKUP_JOBS';
const TRIGGER_HANDLER = 'onInatEdit';
const PROCESSING_NOTE = 'Looking up…';
const API_BASE = 'https://api.inaturalist.org/v1';
const ID_BATCH_SIZE = 100;           // observations fetched per request when looking up by id
const OFV_SEARCH_PER_PAGE = 100;     // page size for observation-field searches
const OFV_SEARCH_MAX_PAGES = 3;      // pages scanned per observation-field lookup
const CACHE_TTL_SECONDS = 1800;      // successful lookups are reused for 30 minutes
const MAX_PENDING_JOBS = 50;         // queued edit jobs kept while another run holds the lock
const PROCESSING_NOTE_MAX_ROWS = 3;  // show the "Looking up…" note only for small jobs
const PROCESS_CHUNK_ROWS = 30;       // rows looked up + written per checkpoint
const PROCESS_TIME_BUDGET_MS = 270000; // stop ~4.5 min in (Apps Script kills runs at ~6 min)
const CONTINUATION_HANDLER = 'processPendingJobsTrigger';

// ====== Selectable Fields ======
const OBSERVATION_FIELDS = [
  'Voucher Number(s)',
  'FUNDIS Tag Number',
  'Accession Number',
  'GenBank Accession Number',
  'Sequencing Failed',
  'Repeat Sequencing Requested',
  'MycoMap BLAST Results',
  'Sequence Validator',
  'DNA Barcode ITS',
  'Associated Observation',
  'Sensu Stricto',
  'Genetic Sample ID'
];

const STANDARD_METADATA_FIELDS = [
  'id',
  'URL',
  'username',
  'Scientific Name',
  'place_guess',
  'observed_on',
  'description',
  'tag_list',
  'coordinates_obscured',
  'latitude',
  'longitude'
];

const DEFAULT_RETURN_FIELDS = ['URL', 'username', 'Voucher Number(s)', 'Scientific Name'];

// The script watches any column whose header matches one of these
// (matching ignores case, spacing, and punctuation).
const LOOKUP_OPTIONS = [
  'id',
  'Voucher Number(s)',
  'FUNDIS Tag Number',
  'Accession Number',
  'GenBank Accession Number'
];

const STANDARD_EXTRACTORS = {
  'id':          o => o?.id ?? '',
  'URL':         o => o?.id ? `https://www.inaturalist.org/observations/${o.id}` : '',
  'username':    o => o?.user?.login ?? '',
  'Scientific Name': o => o?.taxon?.name ?? '',
  'place_guess': o => o?.place_guess ?? '',
  'observed_on': o => o?.observed_on ?? '',
  'description': o => o?.description ?? '',
  'tag_list':    o => {
    if (!o?.tags) return '';
    if (Array.isArray(o.tags)) {
      if (typeof o.tags[0] === 'string') return o.tags.join(',');
      if (typeof o.tags[0] === 'object' && o.tags[0]?.name) return o.tags.map(t => t.name).join(',');
    }
    return '';
  },
  'coordinates_obscured': o => (o?.obscured === true ? 'Y' : (o?.obscured === false ? 'N' : '')),
  'latitude':  o => (o?.latitude != null ? o.latitude : (o?.geojson?.coordinates?.[1] ?? '')),
  'longitude': o => (o?.longitude != null ? o.longitude : (o?.geojson?.coordinates?.[0] ?? '')),
};

// === MENU ===
function onOpen() {
  // Contexts without a spreadsheet UI (editor Run button, mobile app, API
  // opens) can't build a menu — skip quietly instead of logging an error.
  let ui;
  try { ui = SpreadsheetApp.getUi(); } catch (_) { return; }
  // Note: only reads a document property here — calling ScriptApp from onOpen
  // can throw in limited-auth contexts and would kill the whole menu.
  const autoOn = isAutoLookupEnabled();
  ui
    .createMenu(MENU_NAME)
    .addItem('Process All Rows', 'processAllRows')
    .addItem('Process Selected Rows', 'processSelectedRows')
    .addItem('Start From Row…', 'startFromRowPrompt')
    .addSeparator()
    .addItem('Choose Return Fields (Sidebar)', 'openFieldSelectorSidebar')
    .addSeparator()
    .addItem(autoOn ? 'Auto-lookup: ON  (click to disable)'
                    : 'Auto-lookup: OFF (click to enable)', 'toggleAutoLookup')
    .addSeparator()
    .addItem('Open Log', 'openLogSheet')
    .addToUi();
}

// === AUTO-LOOKUP TRIGGER ===
function toggleAutoLookup() {
  isAutoLookupEnabled() ? disableAutoLookup() : enableAutoLookup();
  onOpen();
}

// The document property is the single source of truth for on/off. The trigger
// is just plumbing: onInatEdit checks the flag itself, so disabling from any
// account stops lookups even when the trigger belongs to a different user.
function isAutoLookupEnabled() {
  return PropertiesService.getDocumentProperties().getProperty(PROP_AUTO_LOOKUP) === '1';
}

function enableAutoLookup() {
  removeAutoLookupTriggers();
  ScriptApp.newTrigger(TRIGGER_HANDLER)
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
  PropertiesService.getDocumentProperties().setProperty(PROP_AUTO_LOOKUP, '1');
  uiToast('Auto-lookup ON. Edit any recognized lookup column to fetch automatically.');
}

function disableAutoLookup() {
  removeAutoLookupTriggers();
  PropertiesService.getDocumentProperties().setProperty(PROP_AUTO_LOOKUP, '0');
  uiToast('Auto-lookup OFF.');
}

function removeAutoLookupTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === TRIGGER_HANDLER) ScriptApp.deleteTrigger(t);
  });
}

function onInatEdit(e) {
  try {
    if (!e || !e.range) return;
    if (!isAutoLookupEnabled()) return;
    startRunClock();
    const sheet = e.range.getSheet();
    if (sheet.getName() === LOG_SHEET_NAME) return;

    const rowStart = Math.max(2, e.range.getRow());
    const rowEnd = e.range.getRow() + e.range.getNumRows() - 1;
    if (rowStart > rowEnd) return;

    const { lookupCols } = readConfig(sheet);
    if (!lookupCols.length) return;

    const colStart = e.range.getColumn();
    const colEnd = colStart + e.range.getNumColumns() - 1;
    const editedCols = lookupCols
      .filter(c => c.index + 1 >= colStart && c.index + 1 <= colEnd)
      .map(c => c.index);
    if (!editedCols.length) return;

    const job = { sheetName: sheet.getName(), rowStart, rowEnd, editedCols };
    const lock = LockService.getDocumentLock();
    if (!lock.tryLock(20000)) {
      // A previous run is still busy — queue this edit instead of dropping it.
      enqueueJob(job);
      return;
    }
    try {
      runJob(job);
      drainPendingJobs();
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    try { logSingleRow(0, 'Trigger', '', 'TriggerError', String(err && err.message ? err.message : err)); } catch (_) {}
  }
}

// === RUN CLOCK ===
// Apps Script hard-kills executions at ~6 minutes. Every entry point starts
// this clock; long jobs stop before the limit, checkpoint what they have, and
// queue the remainder for a continuation run.
let RUN_DEADLINE = 0;
function startRunClock() { RUN_DEADLINE = Date.now() + PROCESS_TIME_BUDGET_MS; }
function pastDeadline() { return RUN_DEADLINE > 0 && Date.now() > RUN_DEADLINE; }

/** Schedule a one-off continuation run (~1 min out) to drain the queue. */
function scheduleContinuation() {
  try {
    const exists = ScriptApp.getProjectTriggers()
      .some(t => t.getHandlerFunction() === CONTINUATION_HANDLER);
    if (!exists) {
      ScriptApp.newTrigger(CONTINUATION_HANDLER).timeBased().after(60 * 1000).create();
    }
  } catch (_) {}
}

function processPendingJobsTrigger() {
  startRunClock();
  try {
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === CONTINUATION_HANDLER) ScriptApp.deleteTrigger(t);
    });
  } catch (_) {}
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) { scheduleContinuation(); return; }
  try {
    drainPendingJobs();
  } finally {
    lock.releaseLock();
  }
}

// === PENDING-JOB QUEUE ===
// Edits that arrive while another run holds the document lock are queued in a
// document property and drained by whichever run finishes first.
function enqueueJob(job) {
  const qLock = LockService.getScriptLock();
  if (!qLock.tryLock(5000)) return;
  try {
    const props = PropertiesService.getDocumentProperties();
    let jobs = [];
    try { jobs = JSON.parse(props.getProperty(PROP_PENDING_JOBS) || '[]'); } catch (_) {}
    const key = JSON.stringify(job);
    if (jobs.length < MAX_PENDING_JOBS && !jobs.some(j => JSON.stringify(j) === key)) {
      jobs.push(job);
      props.setProperty(PROP_PENDING_JOBS, JSON.stringify(jobs));
    }
  } finally {
    qLock.releaseLock();
  }
}

function takePendingJobs() {
  const qLock = LockService.getScriptLock();
  if (!qLock.tryLock(5000)) return [];
  try {
    const props = PropertiesService.getDocumentProperties();
    let jobs = [];
    try { jobs = JSON.parse(props.getProperty(PROP_PENDING_JOBS) || '[]'); } catch (_) {}
    if (jobs.length) props.deleteProperty(PROP_PENDING_JOBS);
    return jobs;
  } finally {
    qLock.releaseLock();
  }
}

function drainPendingJobs() {
  for (let i = 0; i < 20; i++) {
    if (pastDeadline()) { scheduleContinuation(); return; }
    const jobs = takePendingJobs();
    if (!jobs.length) return;
    for (let j = 0; j < jobs.length; j++) {
      if (pastDeadline()) {
        jobs.slice(j).forEach(enqueueJob);
        scheduleContinuation();
        return;
      }
      runJob(jobs[j]);
    }
  }
}

function runJob(job) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(job.sheetName);
  if (!sheet) return;
  const rowEnd = Math.min(job.rowEnd, sheet.getLastRow());
  if (job.rowStart > rowEnd) return;
  processRows(sheet, job.rowStart, rowEnd, { editedCols: job.editedCols || null, quiet: true });
}

// === ENTRY POINTS (manual) ===
function withDocumentLock(fn) {
  startRunClock();
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) return uiToast('Another lookup is still running — try again shortly.');
  try {
    fn();
    drainPendingJobs();
  } finally {
    lock.releaseLock();
  }
}

function processAllRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return uiToast('No data rows.');
  withDocumentLock(() => processRows(sheet, 2, lastRow, {}));
}

function processSelectedRows() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const range = sheet.getActiveRange();
  if (!range) return uiAlert('Select one or more rows first.');
  const start = range.getRow();
  const end = start + range.getNumRows() - 1;
  const firstDataRow = Math.max(2, start);
  if (firstDataRow > end) return uiAlert('Selection does not include any data rows.');
  withDocumentLock(() => processRows(sheet, firstDataRow, end, {}));
}

function startFromRowPrompt() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Start From Row', 'Enter a 1-based row number (≥ 2):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;
  const row = parseInt(resp.getResponseText(), 10);
  if (!Number.isFinite(row) || row < 2) return uiAlert('Please enter an integer ≥ 2.');
  const sheet = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  if (row > lastRow) return uiAlert(`Sheet only has ${lastRow} row(s).`);
  withDocumentLock(() => processRows(sheet, row, lastRow, {}));
}

// === CORE ===
/**
 * Fill return fields for rows rowStart..rowEnd.
 *
 * opts.editedCols — set by the auto-lookup trigger: only these lookup columns
 *   are considered, and rows whose edited cell is empty are skipped. Manual
 *   runs (no editedCols) prefer a filled "id" column (exact, one request per
 *   batch), then the other lookup columns in header order.
 * opts.quiet — no alerts/toasts (trigger mode).
 */
function processRows(sheet, rowStart, rowEnd, opts) {
  opts = opts || {};
  const cfg = readConfig(sheet);
  const { lookupCols } = cfg;
  const selectedFields = cfg.selectedFields.slice();

  if (!lookupCols.length) {
    if (!opts.quiet) uiAlert('Could not find any lookup columns. Add a column named "id", "FUNDIS Tag Number", or "Voucher Number(s)".');
    return;
  }
  if (rowStart > rowEnd) return;

  // Force 'id' to be returned if using an alternative lookup column.
  const hasAlternativeLookup = lookupCols.some(c => c.name.toLowerCase() !== 'id');
  if (hasAlternativeLookup && !selectedFields.some(f => f.toLowerCase() === 'id')) {
    selectedFields.unshift('id');
  }

  const colMap = ensureColumns(sheet, cfg.headers, selectedFields);
  const targetCols = selectedFields
    .map(f => colMap[normalizeOFVName(f)])
    .filter(c => c !== undefined);
  if (!targetCols.length) return;

  const numRows = rowEnd - rowStart + 1;
  const lastCol = sheet.getLastColumn();
  const sheetData = sheet.getRange(rowStart, 1, numRows, lastCol).getValues();

  // Resolve each row's lookup source.
  const idCol = lookupCols.find(c => c.name.toLowerCase() === 'id') || null;
  const rows = [];
  let skipped = 0;
  for (let i = 0; i < numRows; i++) {
    const resolved = resolveLookup(sheetData[i], lookupCols, idCol, opts.editedCols || null);
    if (!resolved) { skipped++; continue; }
    rows.push(Object.assign({ r: rowStart + i, i: i }, resolved));
  }

  if (!rows.length) {
    if (!opts.quiet) uiToast('No lookup values found in those rows.');
    return;
  }

  // Visual feedback for small jobs only (per-row flushes made big runs crawl).
  if (rows.length <= PROCESSING_NOTE_MAX_ROWS) {
    rows.forEach(x => sheet.getRange(x.r, x.colIndex + 1).setNote(PROCESSING_NOTE));
    SpreadsheetApp.flush();
  }

  const cache = getLookupCache();
  const logs = [];
  const runResults = new Map(); // lookup key -> {obs, note?, error?}

  // Column-value snapshots for the selected fields only (never rewrite
  // unrelated columns that happen to sit between two target columns).
  const colArrays = new Map();
  for (const c of targetCols) {
    if (!colArrays.has(c)) colArrays.set(c, sheetData.map(row => [row[c]]));
  }

  let filled = 0, notFound = 0, errors = 0;
  const deferredRows = []; // rows the time budget cut off

  // Look up, fill, and WRITE a chunk of rows at a time, so a run that dies at
  // the Apps Script execution limit keeps everything already checkpointed and
  // the remainder continues in a later run instead of being lost.
  for (let g = 0; g < rows.length; g += PROCESS_CHUNK_ROWS) {
    const group = rows.slice(g, g + PROCESS_CHUNK_ROWS);
    if (pastDeadline()) {
      deferredRows.push(...group);
      continue;
    }

    // id lookups for this chunk — one batched request
    const idsNeeded = [...new Set(group
      .filter(x => x.name === 'id' && !x.invalid && !runResults.has(lookupKey(x)))
      .map(x => x.value))];
    if (idsNeeded.length) {
      try {
        const obsById = fetchObservationsByIds(idsNeeded, cache);
        obsById.forEach((obs, id) => runResults.set('id|' + id, { obs }));
      } catch (e) {
        const msg = String(e && e.message ? e.message : e);
        idsNeeded.forEach(id => runResults.set('id|' + id, { obs: null, error: msg }));
      }
    }

    // observation-field lookups for this chunk — one search per unique value
    for (const x of group) {
      if (x.name === 'id' || x.invalid) continue;
      const key = lookupKey(x);
      if (runResults.has(key)) continue;
      const cached = cacheGet(cache, key);
      if (cached) { runResults.set(key, { obs: cached }); continue; }
      if (pastDeadline()) break;
      try {
        const found = searchObservationsForOFV(x.name, x.value);
        if (found.obs) cachePut(cache, key, found.obs);
        runResults.set(key, found);
      } catch (e) {
        runResults.set(key, { obs: null, error: String(e && e.message ? e.message : e) });
      }
      Utilities.sleep(REQUEST_DELAY_MS);
    }

    // fill this chunk's rows
    let firstI = null, lastI = null;
    for (const x of group) {
      const cell = sheet.getRange(x.r, x.colIndex + 1);
      let obs = null, note = '', errMsg = '';

      if (x.invalid) {
        errMsg = 'Not a recognizable observation id or URL';
      } else {
        const res = runResults.get(lookupKey(x));
        if (!res) { deferredRows.push(x); continue; } // deadline hit before its lookup ran
        obs = res.obs || null;
        note = res.note || '';
        errMsg = res.error || '';
      }

      if (firstI === null) firstI = x.i;
      lastI = x.i;

      if (errMsg) {
        errors++;
        cell.setNote('Error: ' + errMsg);
        logs.push([new Date(), sheet.getName(), x.r, x.name, x.value, 'Error', errMsg]);
        continue;
      }
      if (!obs) {
        notFound++;
        cell.setNote(`Not found via ${x.name}`);
        logs.push([new Date(), sheet.getName(), x.r, x.name, x.value, 'Not found', 'No observation matched']);
        continue;
      }

      const values = buildRowValues(obs, selectedFields);
      for (const f of selectedFields) {
        const c = colMap[normalizeOFVName(f)];
        if (c !== undefined) colArrays.get(c)[x.i][0] = values[f] ?? '';
      }
      filled++;
      if (note) cell.setNote(note); else cell.clearNote();
      logs.push([new Date(), sheet.getName(), x.r, x.name, x.value, 'OK',
                 note || `Filled ${selectedFields.length} field(s)`]);
    }

    // checkpoint: write this chunk's row span, one column at a time
    if (firstI !== null) {
      const span = lastI - firstI + 1;
      for (const [c, arr] of colArrays) {
        sheet.getRange(rowStart + firstI, c + 1, span, 1)
             .setValues(arr.slice(firstI, lastI + 1));
      }
      SpreadsheetApp.flush();
    }
  }

  writeLogsBulk(logs);

  // Queue whatever the time budget cut off; a continuation run picks it up.
  if (deferredRows.length) {
    enqueueJob({
      sheetName: sheet.getName(),
      rowStart: Math.min.apply(null, deferredRows.map(x => x.r)),
      rowEnd: rowEnd,
      editedCols: opts.editedCols || null
    });
    scheduleContinuation();
  }

  if (!opts.quiet) {
    const done = rows.length - deferredRows.length;
    const parts = [`${filled} filled`];
    if (notFound) parts.push(`${notFound} not found`);
    if (errors) parts.push(`${errors} error(s)`);
    if (skipped) parts.push(`${skipped} skipped (no lookup value)`);
    let msg = `Processed ${done} row(s): ${parts.join(', ')}`;
    if (deferredRows.length) msg += `; ${deferredRows.length} more continue automatically in ~1 min`;
    uiToast(msg);
  }
}

/** Cache/dedup key for a resolved lookup. */
function lookupKey(x) {
  return x.name.toLowerCase() === 'id'
    ? 'id|' + x.value
    : normalizeOFVName(x.name) + '|' + normalizeValue(x.value);
}

/**
 * Pick which lookup column drives a row.
 * Trigger mode (editedCols set): only the edited lookup column(s) count, and
 * an emptied cell means "skip", never "fall back to some other column".
 * Manual mode: a filled "id" column wins, then header order.
 */
function resolveLookup(rowValues, lookupCols, idCol, editedCols) {
  const candidates = editedCols
    ? lookupCols.filter(c => editedCols.indexOf(c.index) !== -1)
    : ((idCol && String(rowValues[idCol.index]).trim()) ? [idCol] : lookupCols);
  for (const col of candidates) {
    const raw = String(rowValues[col.index]).trim();
    if (!raw) continue;
    if (col.name.toLowerCase() === 'id') {
      const id = extractObservationId(raw);
      if (!id) return { name: 'id', value: raw, colIndex: col.index, invalid: true };
      return { name: 'id', value: id, colIndex: col.index };
    }
    return { name: col.name, value: raw, colIndex: col.index };
  }
  return null;
}

/** Accepts a bare observation id or a pasted iNaturalist observation URL. */
function extractObservationId(raw) {
  const s = String(raw).trim();
  let m = s.match(/observations\/(\d+)/);
  if (m) return m[1];
  m = s.match(/^(\d+)$/);
  return m ? m[1] : '';
}

function readConfig(sheet) {
  const lastCol = sheet.getLastColumn();
  const headers = lastCol
    ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim())
    : [];

  // Match headers against LOOKUP_OPTIONS ignoring case/spacing/punctuation, so
  // "Voucher Numbers" or "voucher number(s)" still counts.
  const lookupCols = [];
  headers.forEach((h, idx) => {
    const key = normalizeOFVName(h);
    if (!key) return;
    const matched = LOOKUP_OPTIONS.find(opt => normalizeOFVName(opt) === key);
    if (matched) lookupCols.push({ index: idx, name: matched });
  });

  return { headers, lookupCols, selectedFields: getSelectedFields() };
}

// === LOOKUP ===
/** Fetch a JSON payload with retry/backoff on 429 and 5xx. 404 returns null. */
function fetchJson(url) {
  let delay = 1000;
  for (let attempt = 0; attempt < 4; attempt++) {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    const code = resp.getResponseCode();
    if (code === 200) return JSON.parse(resp.getContentText());
    if (code === 404) return null;
    if (code === 429 || code >= 500) {
      Utilities.sleep(delay);
      delay *= 2;
      continue;
    }
    throw new Error(`HTTP ${code} from iNaturalist`);
  }
  throw new Error('iNaturalist API unavailable (rate-limited or down) — try again in a minute');
}

/** Batched id lookup: Map of id -> slim observation (or null if not found). */
function fetchObservationsByIds(ids, cache) {
  const out = new Map();
  const toFetch = [];
  for (const id of ids) {
    const hit = cacheGet(cache, 'id|' + id);
    if (hit) out.set(id, hit);
    else toFetch.push(id);
  }
  for (let i = 0; i < toFetch.length; i += ID_BATCH_SIZE) {
    if (i > 0) Utilities.sleep(REQUEST_DELAY_MS);
    const chunk = toFetch.slice(i, i + ID_BATCH_SIZE);
    const json = fetchJson(`${API_BASE}/observations?id=${chunk.join(',')}&per_page=${ID_BATCH_SIZE}`);
    const found = new Map();
    (json?.results || []).forEach(o => found.set(String(o.id), slimObservation(o)));
    for (const id of chunk) {
      const obs = found.get(String(id)) || null;
      out.set(id, obs);
      if (obs) cachePut(cache, 'id|' + id, obs);
    }
  }
  return out;
}

/**
 * Find an observation by observation-field value. The API filters
 * `field:Name=value` server-side, so the first page normally settles it;
 * exact-match verification (normalized) still runs client-side, and a few more
 * pages are scanned only if the filter returned inexact results.
 */
function searchObservationsForOFV(ofvName, ofvValue) {
  const nameKey = normalizeOFVName(ofvName);
  const valKey = normalizeValue(ofvValue);
  const fieldParam = 'field:' + encodeURIComponent(ofvName) + '=' + encodeURIComponent(ofvValue);
  const matches = [];
  for (let page = 1; page <= OFV_SEARCH_MAX_PAGES; page++) {
    if (page > 1) Utilities.sleep(REQUEST_DELAY_MS);
    const url = `${API_BASE}/observations?per_page=${OFV_SEARCH_PER_PAGE}&page=${page}` +
                `&order=desc&order_by=created_at&${fieldParam}`;
    const json = fetchJson(url);
    const results = json?.results || [];
    for (const o of results) {
      const ofvs = o?.ofvs || o?.observation_field_values || [];
      const exact = ofvs.some(it => {
        const n = normalizeOFVName(it?.name ?? it?.observation_field?.name ?? '');
        const v = normalizeValue(String(it?.value ?? ''));
        return n === nameKey && v === valKey;
      });
      if (exact) matches.push(o);
    }
    if (matches.length || results.length < OFV_SEARCH_PER_PAGE) break;
  }
  if (!matches.length) return { obs: null, note: '' };
  return {
    obs: slimObservation(matches[0]),
    note: matches.length > 1 ? `${matches.length} observations share this value; used the most recent` : ''
  };
}

/** Keep only the fields the extractors read — small enough to cache. */
function slimObservation(o) {
  if (!o) return null;
  let lat = o?.latitude ?? null;
  let lng = o?.longitude ?? null;
  if (lat == null && Array.isArray(o?.geojson?.coordinates)) {
    lng = o.geojson.coordinates[0];
    lat = o.geojson.coordinates[1];
  }
  if (lat == null && typeof o?.location === 'string' && o.location.indexOf(',') > 0) {
    const parts = o.location.split(',');
    lat = parts[0];
    lng = parts[1];
  }
  const ofvs = (o.ofvs || o.observation_field_values || []).map(it => ({
    name: String(it?.name ?? it?.observation_field?.name ?? ''),
    value: it?.value ?? ''
  }));
  return {
    id: o.id,
    user: { login: o?.user?.login ?? '' },
    taxon: { name: o?.taxon?.name ?? '' },
    place_guess: o?.place_guess ?? '',
    observed_on: o?.observed_on ?? '',
    description: o?.description ?? '',
    tags: Array.isArray(o?.tags)
      ? o.tags.map(t => (t && typeof t === 'object' ? String(t.name ?? '') : String(t)))
      : [],
    obscured: o?.obscured,
    latitude: lat,
    longitude: lng,
    ofvs: ofvs
  };
}

// === OUTPUT BUILDING ===
function buildRowValues(obs, selectedFields) {
  const values = {};
  const stdSet = new Set(STANDARD_METADATA_FIELDS.map(s => s.toLowerCase()));
  const ofvMap = buildOfvMapNormalized(obs);

  for (const f of selectedFields) {
    if (stdSet.has(f.toLowerCase())) {
      const extractor = STANDARD_EXTRACTORS[f];
      values[f] = extractor ? String(safe(() => extractor(obs))).trim() : '';
    } else {
      const key = normalizeOFVName(f);
      values[f] = ofvMap.get(key) ?? '';
    }
  }
  return values;
}

function buildOfvMapNormalized(obs) {
  const m = new Map();
  const ofvs = obs?.ofvs || obs?.observation_field_values || [];
  for (const it of ofvs) {
    const nameRaw = String(it?.name ?? it?.observation_field?.name ?? '').trim();
    if (!nameRaw) continue;
    const key = normalizeOFVName(nameRaw);
    const val = String(it?.value ?? '').trim();
    if (!m.has(key)) m.set(key, val);
  }
  return m;
}

function normalizeOFVName(s) { return (s || '').toLowerCase().replace(/[^a-z0-9]/g, ''); }
function normalizeValue(s)   { return (s || '').toLowerCase().trim(); }

// === CACHE ===
// Successful lookups are cached per document for CACHE_TTL_SECONDS, so a
// re-edit or a re-run does not repeat the API call. Failures are never cached.
function getLookupCache() {
  try { return CacheService.getDocumentCache(); } catch (_) { return null; }
}

function cacheGet(cache, key) {
  if (!cache) return null;
  try {
    const raw = cache.get(cacheKeyFor(key));
    return raw ? JSON.parse(raw) : null;
  } catch (_) { return null; }
}

function cachePut(cache, key, obj) {
  if (!cache) return;
  try {
    const raw = JSON.stringify(obj);
    if (raw.length < 90000) cache.put(cacheKeyFor(key), raw, CACHE_TTL_SECONDS);
  } catch (_) {}
}

function cacheKeyFor(key) {
  const full = 'inat|' + key;
  if (full.length <= 240) return full;
  return 'inat|' + Utilities.base64Encode(
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, key));
}

// === COLUMNS / WRITES ===
function ensureColumns(sheet, headers, selectedFields) {
  // Map normalized header -> 0-based column index; append any missing fields.
  const map = {};
  headers.forEach((h, i) => {
    const k = normalizeOFVName(h);
    if (k && !(k in map)) map[k] = i;
  });

  const newHeaders = [];
  for (const f of selectedFields) {
    const k = normalizeOFVName(f);
    if (!(k in map)) {
      map[k] = headers.length + newHeaders.length;
      newHeaders.push(f);
    }
  }

  if (newHeaders.length) {
    const needCols = headers.length + newHeaders.length - sheet.getMaxColumns();
    if (needCols > 0) sheet.insertColumnsAfter(sheet.getMaxColumns(), needCols);
    sheet.getRange(1, headers.length + 1, 1, newHeaders.length).setValues([newHeaders]);
  }
  return map;
}

// === FIELD SELECTION SIDEBAR ===
function openFieldSelectorSidebar() {
  const html = HtmlService.createTemplateFromFile('FieldSelector');
  html.data = {
    observationFields: OBSERVATION_FIELDS,
    metadataFields: STANDARD_METADATA_FIELDS,
    preselected: getSelectedFields(),
    defaults: DEFAULT_RETURN_FIELDS
  };
  SpreadsheetApp.getUi()
    .showSidebar(html.evaluate().setTitle('Select Return Fields').setWidth(360));
}

function saveSelectedFields(selected) {
  if (!Array.isArray(selected)) selected = [];
  setSelectedFields(selected);
  return { ok: true, count: selected.length };
}

// === SELECTED FIELDS STORAGE ===
function getSelectedFields() {
  const raw = PropertiesService.getDocumentProperties().getProperty(PROP_RETURN_FIELDS);
  if (raw) {
    try {
      const arr = JSON.parse(raw);
      if (Array.isArray(arr) && arr.length) return arr;
    } catch (_) {}
  }
  return DEFAULT_RETURN_FIELDS.slice();
}

function setSelectedFields(fieldsArray) {
  PropertiesService.getDocumentProperties()
    .setProperty(PROP_RETURN_FIELDS, JSON.stringify(fieldsArray || []));
}

// === LOG ===
function getOrCreateLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName(LOG_SHEET_NAME);
  if (!log) {
    log = ss.insertSheet(LOG_SHEET_NAME);
    log.appendRow(['Timestamp', 'Sheet', 'Row #', 'Lookup Field', 'Lookup Value', 'Status', 'Message']);
  }
  return log;
}

function writeLogsBulk(logsArray) {
  if (!logsArray || logsArray.length === 0) return;
  const logSheet = getOrCreateLogSheet();
  const startRow = logSheet.getLastRow() + 1;
  logSheet.getRange(startRow, 1, logsArray.length, logsArray[0].length).setValues(logsArray);
}

function logSingleRow(rowNumber, lookupField, lookupValue, status, message) {
  writeLogsBulk([[new Date(), SpreadsheetApp.getActiveSheet().getName(), rowNumber, lookupField, lookupValue, status, message]]);
}

function openLogSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = getOrCreateLogSheet();
  ss.setActiveSheet(log);
}

// === UTILS ===
function uiAlert(msg) { SpreadsheetApp.getUi().alert(msg); }
function uiToast(msg) { SpreadsheetApp.getActive().toast(msg, MENU_NAME, 5); }
function safe(fn) { try { return fn(); } catch (e) { return ''; } }
