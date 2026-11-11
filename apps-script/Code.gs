/**
 * @OnlyCurrentDoc
 * (Limits the permission to this one spreadsheet, so Google asks for a narrow access instead of
 * all spreadsheets. Needs this script to be opened from the sheet: Extensions > Apps Script.)
 *
 * PG Creative Requirement Dashboard
 * Simple form that writes ONLY into the blue columns of the
 * "Creative Tracker/Requirement - 2026" tab. All other columns
 * (formulas, strategist columns, etc.) are never touched.
 */

const SHEET_ID = '1x5E1LJ263AY1tAXOxDiwG566WMgCgBNYTi1axqQAdxM';
const TAB_NAME = 'Creative Tracker/Requirement - 2026';

// Blue columns (1 = A). Header keyword is only used for a safety check.
const COL = {
  month: 5,       // E  Month
  date: 6,        // F  Date
  format: 8,      // H  Format
  link: 9,        // I  Product/Collection Link
  concept: 10,    // J  Concept
  reference: 11,  // K  Reference (Mandatory)
  qty: 12,        // L  Quantity
  sizes: 15,      // O  Sizes
  status: 16,     // P  Status
  comments: 18,   // R  Additional Comments
  variations: 19  // S  Varations
};
const BLUE_COLS = Object.values(COL);
const TESTNO_COL = 4; // D, read only (shown in tracker list)
const TOTAL_COLS = 22; // A to V

const HEADER_CHECK = {
  5: 'month', 6: 'date', 8: 'format', 9: 'product', 10: 'concept',
  11: 'reference', 12: 'quantity', 15: 'size', 16: 'status',
  18: 'additional', 19: 'varation'
};

const DEFAULTS = {
  format: ['Static', 'Video', 'UGC', 'GIF', 'Carousel'],
  concept: ['Product Led', 'Testimonial', 'UGC', 'Offer', 'Problem Solution', 'Comparison'],
  status: ['Waiting for Assets', 'Waiting For Approval', 'Awaiting Script Approval', 'Approved',
           'In Progress', 'Needs Edit', 'Size Changes', 'variations added', 'Live'],
  variations: ['No Variation Needed', 'Variation Needed'],
  sizes: ['1:1', '4:5', '9:16', '16:9', '1.91:1']
};
const DEFAULT_STATUS = 'In Progress';

// Opening the web app link directly shows the form (needs the Index HTML file in this project).
// If that file is not added, the link just answers that the API is running, which is all the
// GitHub Pages version of the form needs.
function doGet() {
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('PG Creative Requests')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return ContentService.createTextOutput('PG Creative API is running.');
  }
}

// The GitHub Pages form talks to the sheet through here. Only these four functions can be called.
const API_FNS = {
  getConfig: function () { return getConfig(); },
  submitRequest: function (d) { return submitRequest(d); },
  getTracker: function () { return getTracker(); },
  updateStatus: function (row, status) { return updateStatus(row, status); }
};

function doPost(e) {
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    const fn = API_FNS[req.fn];
    if (!fn) throw new Error('Unknown action');
    out = { ok: true, data: fn.apply(null, req.args || []) };
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ---------- helpers ---------- */

function getSheet_() {
  // Bound to the sheet (opened from Extensions > Apps Script): use it directly, narrow permission.
  // Fallback to the ID keeps it working if the script is ever used as a standalone project.
  let ss = null;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) {}
  if (!ss || ss.getId() !== SHEET_ID) ss = SpreadsheetApp.openById(SHEET_ID);
  // Match the tab name ignoring extra spaces and capital letters (the real tab name has a trailing space).
  const norm = function (x) { return String(x).replace(/\s+/g, ' ').trim().toLowerCase(); };
  const want = norm(TAB_NAME);
  const sheets = ss.getSheets();
  let sh = null;
  for (let i = 0; i < sheets.length; i++) { if (norm(sheets[i].getName()) === want) { sh = sheets[i]; break; } }
  if (!sh) throw new Error('Tab not found: ' + TAB_NAME + '. Tabs in this file: ' + sheets.map(function (t) { return '"' + t.getName() + '"'; }).join(', '));
  return sh;
}

function headerRow_(sh) {
  const n = Math.min(10, sh.getMaxRows());
  const v = sh.getRange(1, TESTNO_COL, n, 1).getDisplayValues();
  for (let i = 0; i < v.length; i++) {
    if (/test\s*no/i.test(v[i][0])) return i + 1;
  }
  return 1;
}

function lastBlueRow_(sh, hr) {
  const max = sh.getMaxRows();
  if (max <= hr) return hr;
  const vals = sh.getRange(hr + 1, 8, max - hr, 4).getValues(); // H to K
  for (let i = vals.length - 1; i >= 0; i--) {
    if (vals[i].some(function (v) { return String(v).trim() !== ''; })) return hr + 1 + i;
  }
  return hr;
}

function readList_(sh, col, rows) {
  const max = sh.getMaxRows();
  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];
    if (r < 1 || r > max) continue;
    const rule = sh.getRange(r, col).getDataValidation();
    if (!rule) continue;
    const t = rule.getCriteriaType();
    const v = rule.getCriteriaValues();
    if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_LIST) {
      return v[0].map(String);
    }
    if (t === SpreadsheetApp.DataValidationCriteria.VALUE_IN_RANGE) {
      const vals = v[0].getValues().reduce(function (a, b) { return a.concat(b); }, [])
        .map(String).filter(function (s) { return s.trim() !== ''; });
      return Array.from(new Set(vals));
    }
  }
  return null;
}

function optionsFor_(sh, key, hr, last) {
  const col = COL[key];
  const rows = [last, last - 1, last - 2, hr + 1, hr + 2];
  const fromRule = readList_(sh, col, rows);
  if (fromRule && fromRule.length) return fromRule;

  const set = new Set(DEFAULTS[key] || []);
  if (last > hr) {
    sh.getRange(hr + 1, col, last - hr, 1).getDisplayValues().forEach(function (r) {
      const s = String(r[0]).trim();
      if (s) set.add(s);
    });
  }
  return Array.from(set);
}

/* ---------- called from the page ---------- */

function getConfig() {
  const sh = getSheet_();
  const hr = headerRow_(sh);
  const last = lastBlueRow_(sh, hr);

  const warnings = [];
  const header = sh.getRange(hr, 1, 1, TOTAL_COLS).getDisplayValues()[0];
  Object.keys(HEADER_CHECK).forEach(function (c) {
    const h = String(header[Number(c) - 1] || '').toLowerCase();
    if (h.indexOf(HEADER_CHECK[c]) === -1) {
      warnings.push('Column ' + colLetter_(Number(c)) + ' header is "' + header[Number(c) - 1] +
        '", expected something like "' + HEADER_CHECK[c] + '". Columns may have moved.');
    }
  });

  const options = {
    format: optionsFor_(sh, 'format', hr, last),
    concept: optionsFor_(sh, 'concept', hr, last),
    status: optionsFor_(sh, 'status', hr, last),
    variations: optionsFor_(sh, 'variations', hr, last),
    sizes: optionsFor_(sh, 'sizes', hr, last)
  };
  const defStatus = options.status.indexOf(DEFAULT_STATUS) > -1 ? DEFAULT_STATUS : '';
  return {
    options: options,
    defaultStatus: defStatus,
    warnings: warnings,
    sheetUrl: 'https://docs.google.com/spreadsheets/d/' + SHEET_ID + '/edit#gid=' + sh.getSheetId()
  };
}

function submitRequest(d) {
  const required = ['date', 'format', 'link', 'concept', 'reference', 'qty', 'status'];
  for (let i = 0; i < required.length; i++) {
    if (!d || String(d[required[i]] || '').trim() === '') {
      return { ok: false, error: 'Please fill: ' + required[i] };
    }
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  let row = 0;
  try {
    const sh = getSheet_();
    const hr = headerRow_(sh);
    const last = lastBlueRow_(sh, hr);
    row = last + 1;

    if (row > sh.getMaxRows()) sh.insertRowsAfter(sh.getMaxRows(), 1);
    prepareRow_(sh, row, hr, last);

    const put = function (key, val) { sh.getRange(row, COL[key]).setValue(val); };
    put('month', d.month);
    put('date', d.date);
    put('format', d.format);
    put('link', String(d.link).trim());
    put('concept', d.concept);
    put('reference', String(d.reference).trim());
    put('qty', isNaN(Number(d.qty)) ? d.qty : Number(d.qty));
    if (d.sizes) put('sizes', d.sizes);
    put('status', d.status);
    if (d.comments) put('comments', d.comments);
    if (d.variations) put('variations', d.variations);

    SpreadsheetApp.flush();
    return { ok: true, row: row, testNo: sh.getRange(row, TESTNO_COL).getDisplayValue() };
  } catch (err) {
    // roll back anything partially written in blue cells
    if (row > 0) {
      try {
        const sh2 = getSheet_();
        BLUE_COLS.forEach(function (c) { sh2.getRange(row, c).clearContent(); });
      } catch (e) {}
    }
    return { ok: false, error: String(err && err.message ? err.message : err) };
  } finally {
    lock.releaseLock();
  }
}

// Copies look, dropdowns and formulas of non-blue columns from the row above,
// only where the new row is still empty.
function prepareRow_(sh, row, hr, last) {
  if (last <= hr) return;
  const src = sh.getRange(last, 1, 1, TOTAL_COLS);
  const dst = sh.getRange(row, 1, 1, TOTAL_COLS);
  src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  src.copyTo(dst, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);

  const f = src.getFormulasR1C1()[0];
  const cur = dst.getFormulas()[0];
  const curVals = dst.getValues()[0];
  for (let c = 0; c < TOTAL_COLS; c++) {
    const colNo = c + 1;
    if (BLUE_COLS.indexOf(colNo) > -1) continue;
    if (f[c] && !cur[c] && String(curVals[c]).trim() === '') {
      sh.getRange(row, colNo).setFormulaR1C1(f[c]);
    }
  }
}

function getTracker() {
  const sh = getSheet_();
  const hr = headerRow_(sh);
  const last = lastBlueRow_(sh, hr);
  const out = { rows: [], counts: {}, total: 0 };
  if (last <= hr) return out;

  const vals = sh.getRange(hr + 1, 1, last - hr, TOTAL_COLS).getDisplayValues();
  for (let i = vals.length - 1; i >= 0; i--) {
    const v = vals[i];
    const hasData = [v[7], v[8], v[9], v[10]].some(function (x) { return String(x).trim() !== ''; });
    if (!hasData) continue;
    out.total++;
    const st = String(v[COL.status - 1]).trim() || '(blank)';
    out.counts[st] = (out.counts[st] || 0) + 1;
    if (out.rows.length < 60) {
      out.rows.push({
        row: hr + 1 + i,
        testNo: v[TESTNO_COL - 1],
        month: v[COL.month - 1],
        date: v[COL.date - 1],
        format: v[COL.format - 1],
        link: v[COL.link - 1],
        concept: v[COL.concept - 1],
        qty: v[COL.qty - 1],
        status: v[COL.status - 1],
        variations: v[COL.variations - 1],
        comments: v[COL.comments - 1]
      });
    }
  }
  return out;
}

function updateStatus(row, status) {
  try {
    const sh = getSheet_();
    const hr = headerRow_(sh);
    if (row <= hr || row > sh.getMaxRows()) return { ok: false, error: 'Invalid row' };
    sh.getRange(row, COL.status).setValue(status);
    SpreadsheetApp.flush();
    return { ok: true };
  } catch (err) {
    return { ok: false, error: String(err && err.message ? err.message : err) };
  }
}

function colLetter_(n) {
  let s = '';
  while (n > 0) {
    const m = (n - 1) % 26;
    s = String.fromCharCode(65 + m) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
