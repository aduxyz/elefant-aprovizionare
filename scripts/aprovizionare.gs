// ============================================================
const VERSION = 'v2.45'; // 2026-04-09 14:30:00
const RECENT_TITLES = 30; // days — titles with SalesLY=0 created within this window are included
const MOQ = 2;            // Minimum Order Quantity — applied when Cantitate > 0
//
// Re-run behaviour: if sheet exists → clearContent (keeps formatting),
//   then rewrites ALL content (headers + data + formulas).
//   Formatting is applied only on first creation.
//
// reissues: "Activ=DA" = the edition with the most recent PublishDate
//           in the group (unique criterion; all others = NU)
//
// procurement = 45 cols written by script, sorted:
//   1. Supplier desc by Total Necesar (highest-need supplier first)
//   2. Within supplier: Necesar desc, then PublishDate desc
//   Spreadsheet errors (#DIV/0! etc.) replaced with ''
//
// MAIN source: EBS D.O.I. (BRUT) — 33-col ERP export.
//   Script ONLY reads MAIN; never writes to it.
//   Cols calculated by script → aprovizionare: J=Necesar, K=Cantitate,
//   L=DOS, S=VZ medie/zi. Inactiv removed (not in EBS BRUT export).
//
// USAGE:
//   Menu "📦 Aprovizionare" → "🔄 Pregătire aprovizionare"
// ============================================================

// ── Avg daily sales formula (column S in aprovizionare) ──────
// Replaces MAIN::S (=O/30) with a weighted blend: 60% current week + 40% prev month.
// MAX(0,...) guards against negative prev-month values on edge-case data.
const AVG_DAILY_SALES_FORMULA = '=N2/7*0.6+MAX(0,(P2-O2)/30)*0.4';

// ── Quantity formula (column K in aprovizionare) ─────────────
// v4_adaptive formula (2026-03-29): auto-correcting cap + S1W spike detection + hard cap.
// 1. autocap = MIN(SalesLM, MAX(prev_month, L2M*0.55))
// 2. S1W correction (SalesLM>=10 and S1W*4 < SalesLM):
//    ratio = S1W*4/SalesLM; correction = ratio^0.7 * 0.8 + 0.2
//    Effect: spike over (S1W≈0) → forecast × 0.2 | spike current → no change
// 3. Hard cap: MIN(result, MAX(5, S1W*4)) — never order more than ~4 weeks of recent sales
// To update: edit the string below, then re-run "Pregătire aprovizionare".
// Row references use row 2 — adjustFormulaRow_() shifts them for each data row.
// Legend: N2=SalesL1W, O2=SalesLM, P2=SalesL2M, I2=OnlineStock
const QUANTITY_FORMULA = `=LET(q,MIN(MAX(0,ROUND(MIN(O2,MAX(P2-O2,P2*0.55))*IF(AND(O2>=10,N2*4<O2),POWER(N2*4/O2,0.7)*0.8+0.2,1)*1.15-I2,0)),MAX(5,N2*4)),IF(q>0,MAX(${MOQ},q),0))`;

// Public entry points for library use (underscore functions are private in libraries)
function exportOrdersToNewSpreadsheet() { exportOrdersToNewSpreadsheet_(); }
function exportAllOrders()              { exportAllOrders_(); }

// ── Sheet names ─────────────────────────────────────────────
const MAIN_SHEET_NAME          = 'MAIN';
const APROVIZIONARE_SHEET_NAME = 'aprovizionare';
const REEDITARI_SHEET_NAME     = 'reeditări';
const DASHBOARD_SHEET_NAME     = 'dashboard furnizori';
const COMENZI_SHEET_NAME       = 'comenzi';
const LISTA_COMENZI_SHEET_NAME = 'listă comenzi';

// ── MAIN column indices (0-based) — EBS D.O.I. (BRUT) format, 33 cols ──
// Source: EBS ERP export. Script ONLY reads MAIN; never writes to it.
// Necesar, Cantitate, DOS, VZ medie/zi are calculated by script → aprovizionare.
const C = {
  ElefantSKU:           0,  // A
  EAN:                  1,  // B
  ArticleCode:          2,  // C
  Title:                3,  // D
  Author:               4,  // E
  Supplier:             5,  // F
  RRP:                  6,  // G
  Discount:             7,  // H
  OnlineStock:          8,  // I
  AvailRec:             9,  // J  (EBS: Disponibil RECEPTII)
  SalesL1W:            10,  // K
  SalesLM:             11,  // L
  SalesL2M:            12,  // M
  SalesLS:             13,  // N
  SalesLY:             14,  // O  ← filter: > 0 only
  PublishDate:         15,  // P  ← used in reissues (active edition)
  Category:            16,  // Q
  Subcategory:         17,  // R
  Publisher:           18,  // S
  SupplierCode:        19,  // T
  TotalReserved:       20,  // U
  AvailMMAuchan:       21,  // V
  Collection:          22,  // W
  Format:              23,  // X
  AgeGroup:            24,  // Y
  SubscriptionSkuCode: 25,  // Z
  AllPurchases:        26,  // AA
  SalesAll:            27,  // AB
  OutOfPrint:          28,  // AC  ← flag Zombie
  Unavailable:         29,  // AD  ← flag Zombie
  AvailCustf:          30,  // AE
  AvailStpr:           31,  // AF
  PurchasePrice:       32   // AG  ← flag Bargain
};
const NUM_MAIN_COLS = 33;

// ── Calculated column formulas written to aprovizionare ──────────────
const REQUIRED_FORMULA = '=O2-I2';              // J: Necesar = SalesLM − OnlineStock
const DOS_FORMULA      = '=IFERROR((K2+I2)/S2,0)';  // L: DOS = (Cantitate + OnlineStock) / VZ medie/zi

// ── Aprovizionare output layout ──────────────────────────────────────
// MAIN's 33 cols are expanded to 37 derived cols by inserting 4 calculated cols:
//   MAIN[0..8]  → aprov[0..8]   (A–I: ElefantSKU..OnlineStock)
//   aprov[9]    = Necesar        (J, REQUIRED_FORMULA)
//   aprov[10]   = Cantitate      (K, QUANTITY_FORMULA)
//   aprov[11]   = DOS            (L, DOS_FORMULA)
//   MAIN[9]     → aprov[12]     (M = Disp. Rec.)
//   MAIN[10..14]→ aprov[13..17] (N–R = SalesL1W..SalesLY)
//   aprov[18]   = VZ medie/zi   (S, AVG_DAILY_SALES_FORMULA)
//   MAIN[15..32]→ aprov[19..36] (T–AK)
//   EXTRA[0..7] → aprov[37..44] (AL–AS)
const NUM_APROV_DERIVED = 37;  // derived cols before extras (total = 37 + 8 = 45)
const APROV_CANTITATE   = 10;  // K — Cantitate position in aprovizionare (0-based)
const APROV_PPC         = 36;  // AK — PurchasePrice position in aprovizionare (0-based)

// Headers for the aprovizionare sheet (37 derived cols, in output order)
const APROV_HEADERS = [
  // A–I: from MAIN (0–8)
  'ElefantSKU','EAN','Cod Articol','Articol','Autor','Furnizor',
  'RRP','Reducere','Stoc Online',
  // J–L: calculated by script
  'Necesar','Cantitate','DOS',
  // M: from MAIN (9)
  'Disp. Rec.',
  // N–R: from MAIN (10–14)
  'SalesL1W','SalesLM','SalesL2M','SalesLS','SalesLY',
  // S: calculated by script
  'VZ medie/zi',
  // T–AK: from MAIN (15–32)
  'Data Creare','Categorie','Subcategorie','Producator',
  'Cod Furnizor','Total Rez.','Disp. MM/Auchan','Colectie','Format',
  'Varsta','Cod SKU Abon','Achiz. All','Sales All','Epuizat',
  'Indisponibil','Disp. Custf.','Disp. Stpr.','Pret Achiz. Frz.'
];

// ── Extra columns written by script (AM–AT) ─────────────────
const EXTRA_HEADERS = [
  'REEDITARE',   // AL
  'Nr_Editii',   // AM
  'Alte_EAN',    // AN
  'Chilipir',    // AO
  'Spike_vz',    // AP
  'Suprastoc',   // AQ
  'Ruptură',     // AR
  'Zombie',      // AS
  'Noutăți'      // AT
];
const NUM_EXTRA = EXTRA_HEADERS.length; // 9

// Extra column indices (0-based, relative to extra block)
const EX = {
  REISSUE:       0,  // AL
  EDITION_COUNT: 1,  // AM
  OTHER_EANS:    2,  // AN
  BARGAIN:       3,  // AO
  SALES_SPIKE:   4,  // AP
  OVERSTOCK:     5,  // AQ
  STOCKOUT:      6,  // AR
  ZOMBIE:        7,  // AS
  NOUTATI:       8   // AT
};

// Columns shown in the orders sheet
const ORDER_HEADERS = [
  'ElefantSKU', 'EAN', 'Cod articol', 'Articol', 'Autor', 'Furnizor',
  'RRP', 'Reducere', 'Cantitate'
];

// ============================================================
// TIMER UTILITY
// ============================================================
const LOG_SHEET_NAME   = '_log';
const LOG_MAX_ROWS     = 500; // trim oldest rows when exceeded

function makeTimer() {
  const t0    = Date.now();
  const runTs = new Date();
  let tLast   = t0;
  const laps  = [];

  return {
    lap(label) {
      const now   = Date.now();
      const delta = now - tLast;
      const total = now - t0;
      tLast = now;
      Logger.log('[+' + (delta / 1000).toFixed(2) + 's | ' + (total / 1000).toFixed(2) + 's] ' + label);
      laps.push([new Date(now), label, delta, total]);
    },

    flush(ss) {
      // Write all laps to _log sheet (newest run first, trim oldest from bottom)
      if (!laps.length) return;
      let logSheet = ss.getSheetByName(LOG_SHEET_NAME);
      if (!logSheet) {
        logSheet = ss.insertSheet(LOG_SHEET_NAME);
        logSheet.getRange(1, 1, 1, 4).setValues([['Timestamp', 'Operație', 'Delta (ms)', 'Total (ms)']]);
        logSheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#4472C4').setFontColor('white');
        logSheet.setFrozenRows(1);
      }

      // Total is known at flush time — embed it in the separator label
      const totalSecs = Math.round(laps[laps.length - 1][3] / 1000);
      const separator = [[runTs, '── run start ' + VERSION + ' (total: ' + totalSecs + ' s) ──', '', '']];
      const rows = separator.concat(laps);

      // Insert new run at top (after header), timestamps ascending within run
      logSheet.insertRowsBefore(2, rows.length);
      logSheet.getRange(2, 1, rows.length, 4).setValues(rows);

      // Trim oldest rows from the bottom
      const dataRows = logSheet.getLastRow() - 1;
      if (dataRows > LOG_MAX_ROWS) {
        logSheet.deleteRows(LOG_MAX_ROWS + 2, dataRows - LOG_MAX_ROWS);
      }
    }
  };
}

// Convert 1-based column index to letter(s): 1→"A", 26→"Z", 27→"AA", etc.
// Parses a date value that may be a JS Date object OR a "dd.mm.yyyy" string.
// new Date("26.10.2020") returns NaN in JS (month 26 invalid), so we handle it explicitly.
function parseDate_(d) {
  if (!d) return 0;
  if (d instanceof Date) return isNaN(d.getTime()) ? 0 : d.getTime();
  const m = String(d).match(/^(\d{1,2})\.(\d{1,2})\.(\d{4})$/);
  if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1])).getTime();
  const t = new Date(d).getTime();
  return isNaN(t) ? 0 : t;
}

function colToLetter_(col) {
  let s = '';
  while (col > 0) {
    col--;
    s = String.fromCharCode(65 + col % 26) + s;
    col = Math.floor(col / 26);
  }
  return s;
}

// Copy a formula from MAIN row 2 to a target row in aprovizionare.
// Replaces every cell-reference row number (e.g. K2 → K7) in the formula.
// Also wraps any division operation in IFERROR(..., 0) so #DIV/0! → 0.
// Example: "=(K2+I2)/S2" copied to row 7 → "=IFERROR((K7+I7)/S7,0)"
function adjustFormulaRow_(formula, toRow) {
  // 1. Adjust row numbers: match column letter(s) + optional $ + "2" not followed by digit
  let adjusted = formula.replace(
    /(\$?[A-Z]{1,3}\$?)2(?!\d)/g,
    '$1' + toRow
  );
  // 2. Wrap in IFERROR if the formula contains a division and isn't already wrapped
  if (adjusted.indexOf('/') !== -1 && adjusted.indexOf('IFERROR') === -1) {
    // Strip leading "=" before wrapping
    adjusted = '=IFERROR(' + adjusted.slice(1) + ',0)';
  }
  return adjusted;
}

// Replace spreadsheet error values (#DIV/0!, #REF!, etc.) with empty string
function cleanVal_(v) {
  if (v instanceof Error) return '';
  if (typeof v === 'string' && /^#(DIV\/0|REF|NAME|VALUE|N\/A|NULL|NUM)!$/.test(v)) return '';
  return v;
}

// Return existing sheet or create a new one; signals whether it's new.
// Uses a manual loop instead of getSheetByName() to avoid Unicode/diacritics
// mismatches (e.g. 'ă' precomposed vs combining form) that can cause
// getSheetByName() to return null even when the sheet exists.
function getOrCreateSheet_(ss, name) {
  const sheets = ss.getSheets();
  for (const s of sheets) {
    if (s.getName() === name) return { sheet: s, isNew: false };
  }
  return { sheet: ss.insertSheet(name), isNew: true };
}

// Clear all cell content from row 1 downward, preserving formatting
function clearAllContent_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow > 0 && lastCol > 0) {
    sheet.getRange(1, 1, lastRow, lastCol).clearContent();
  }
}

// Evaluate QUANTITY_FORMULA logic in JS so filteredRows can be sorted by Quantity
// before writing to aprovizionare. Must stay in sync with QUANTITY_FORMULA constant.
// Optional `sums` = {sumL1W, sumLM, sumL2M} — used for reissue groups (combined editions).
function evalQuantity_(r, ex, sums) {
  if (ex[EX.ZOMBIE]) return 0;
  const N  = sums ? sums.sumL1W : (Number(r[C.SalesL1W])    || 0);  // SalesL1W
  const O  = sums ? sums.sumLM  : (Number(r[C.SalesLM])     || 0);  // SalesLM
  const P  = sums ? sums.sumL2M : (Number(r[C.SalesL2M])    || 0);  // SalesL2M
  const I  = Number(r[C.OnlineStock]) || 0;  // OnlineStock

  // 1. Auto-correcting cap: MIN(SalesLM, MAX(prev_month, L2M*0.55))
  const prevMonth = Math.max(0, P - O);
  const autocap   = Math.min(O, Math.max(prevMonth, P * 0.55));

  // 2. S1W correction: smooth spike detection
  let correction = 1;
  if (O >= 10 && N * 4 < O) {
    const ratio = N * 4 / O;
    correction = Math.pow(ratio, 0.7) * 0.8 + 0.2;
  }

  const forecast = Math.max(0, Math.round(autocap * correction * 1.15 - I));

  // 3. Hard cap: never order more than ~4 weeks of recent weekly sales
  const hardCap = Math.max(5, N * 4);
  const qty = Math.min(forecast, hardCap);
  return qty > 0 ? Math.max(MOQ, qty) : 0;
}

// ============================================================
// CUSTOM MENU
// ============================================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📦 Aprovizionare ' + VERSION)
    .addItem('🔄 Pregătire aprovizionare', 'generateSheets')
    .addSeparator()
    .addItem('📤 Exportă comandă furnizor selectat', 'exportOrdersToNewSpreadsheet_')
    .addItem('📦 Exportă comenzi pt furnizorii bifați', 'exportAllOrders_')
    .addToUi();
}

// ============================================================
// MAIN ENTRY POINT
// ============================================================
function generateSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const T = makeTimer();

  // ── 0/6  Remove any active filters from MAIN and aprovizionare ──────
  [MAIN_SHEET_NAME, APROVIZIONARE_SHEET_NAME].forEach(name => {
    const s = getSheetByName_(ss, name);
    if (s) { const f = s.getFilter(); if (f) f.remove(); }
  });
  T.lap('removed filters');

  // ── 1/6  Read MAIN ────────────────────────────────────────
  ss.toast('1/6 — Citesc MAIN...', '📦 Aprovizionare', -1);
  const mainSheet = getSheetByName_(ss, MAIN_SHEET_NAME);
  if (!mainSheet) throw new Error('Sheet-ul MAIN nu există!');

  const allData = mainSheet.getDataRange().getValues();
  const allRows = allData.slice(1); // skip header row
  T.lap('read MAIN: ' + allRows.length + ' rows');

  // ── 2/6  Filter + compute ─────────────────────────────────
  ss.toast('2/6 — Filtrez & calculez...', '📦 Aprovizionare', -1);

  // Build reissue groups first — needed for deduplication during filtering.
  const reissueMap = buildReissueMap_(allRows);
  T.lap('reissueMap: ' + Object.keys(reissueMap).length + ' groups');

  // Build summed sales per reissue group — needed for the filter below.
  const reissueSumMap = buildReissueSumMap_(allRows, reissueMap);

  // Filter rows into aprovizionare. Three conditions — in order:
  //  A. Always exclude: Epuizat=1 OR Indisponibil=1 (can't be ordered regardless of stock).
  //  B. Normal pass:   effectiveSalesLY > 0 AND Necesar (SalesLM − Stock) >= 0.
  //     For reissue groups: use GROUP summed sales so a new edition with SalesLY=0
  //     still passes when the group has historical demand.
  //  C. Recent titles: SalesLY=0 AND Data Creare within RECENT_TITLES days
  //     (new titles have no sales history yet — include them unconditionally).
  const recentCutoff = Date.now() - RECENT_TITLES * 24 * 60 * 60 * 1000;

  function isExcluded_(r) {
    const o = String(r[C.OutOfPrint]).trim().toLowerCase();
    const u = String(r[C.Unavailable]).trim().toLowerCase();
    return o === '1' || o === 'da' || o === 'true' ||
           u === '1' || u === 'da' || u === 'true';
  }

  const allFilteredPairs = allRows
    .map((r, origIdx) => ({ r, origIdx }))
    .filter(({ r }) => {
      if (isExcluded_(r)) return false;                                           // A
      const key  = String(r[C.Title]).trim() + '|' + String(r[C.Author]).trim() + '|' + String(r[C.Supplier]).trim();
      const sums = reissueSumMap[key] || null;
      const effectiveLY = sums ? sums.sumLY : (Number(r[C.SalesLY])  || 0);
      const effectiveLM = sums ? sums.sumLM : (Number(r[C.SalesLM])  || 0);
      if (effectiveLY > 0 && (effectiveLM - (Number(r[C.OnlineStock]) || 0)) >= 0) return true;  // B
      const created = parseDate_(r[C.PublishDate]);
      return effectiveLY === 0 && created > 0 && created >= recentCutoff;        // C
    });
  // Deduplicate reissue groups: keep only the most recently published edition per group
  // (but only if that edition passed the filter above).
  const filteredPairs   = deduplicateReissues_(allFilteredPairs, reissueMap, allRows);
  const filteredRows    = filteredPairs.map(p => p.r);
  const filteredOrigIdx = filteredPairs.map(p => p.origIdx);
  T.lap('filter SalesLY>0 & Necesar>=0 & dedup reissues: ' + filteredRows.length +
        ' rows (of ' + allRows.length + ', before dedup: ' + allFilteredPairs.length + ')');

  // Compute 8 extra columns for filtered rows
  const extraData = computeExtraData_(filteredRows, reissueMap);
  T.lap('extraData: ' + filteredRows.length + ' × ' + NUM_EXTRA);

  // For each filtered row: get the group sums (null for non-reissue articles)
  const salesSumsForFiltered = filteredRows.map(r => {
    const key = String(r[C.Title]).trim() + '|' + String(r[C.Author]).trim() + '|' + String(r[C.Supplier]).trim();
    return reissueSumMap[key] || null;
  });

  // ── Sort filteredRows + extraData + origIdx together ───────
  // Primary:   supplier desc by total Required (Necesar = SalesLM − OnlineStock)
  // Secondary: Required (Necesar) desc
  // Tertiary:  PublishDate desc (newest first)
  // Note: Required and Quantity use summed SalesLM for reissue articles.
  const quantityCache = filteredRows.map((r, i) => evalQuantity_(r, extraData[i], salesSumsForFiltered[i]));
  const requiredCache = filteredRows.map((r, i) => {
    const sums = salesSumsForFiltered[i];
    const slm = sums ? sums.sumLM : (Number(r[C.SalesLM]) || 0);
    return Math.max(0, slm - (Number(r[C.OnlineStock]) || 0));
  });

  const supplierRequiredTotal = {};
  for (let i = 0; i < filteredRows.length; i++) {
    const supplier = String(filteredRows[i][C.Supplier]).trim();
    supplierRequiredTotal[supplier] =
      (supplierRequiredTotal[supplier] || 0) + requiredCache[i];
  }

  const sortedPairs = filteredRows
    .map((r, i) => ({ r, ex: extraData[i], origIdx: filteredOrigIdx[i], qty: quantityCache[i], req: requiredCache[i], sums: salesSumsForFiltered[i] }))
    .sort((a, b) => {
      const totalA = supplierRequiredTotal[String(a.r[C.Supplier]).trim()] || 0;
      const totalB = supplierRequiredTotal[String(b.r[C.Supplier]).trim()] || 0;
      if (totalB !== totalA) return totalB - totalA;                          // desc by supplier total Required (Necesar)
      if (b.req !== a.req)   return b.req - a.req;                            // desc by row Required (Necesar)
      return parseDate_(b.r[C.PublishDate]) - parseDate_(a.r[C.PublishDate]);    // desc by PublishDate
    });
  const sortedRows    = sortedPairs.map(p => p.r);
  const sortedExtra   = sortedPairs.map(p => p.ex);
  const sortedSums    = sortedPairs.map(p => p.sums);
  const sortedOrigIdx = sortedPairs.map(p => p.origIdx);
  const sortedQty     = sortedPairs.map(p => p.qty);
  T.lap('sort: ' + sortedRows.length + ' rows');

  // ── 3/6  Procurement sheet ────────────────────────────────
  ss.toast('3/6 — Generez aprovizionare...', '📦 Aprovizionare', -1);
  generateProcurement_(ss, sortedRows, sortedExtra, sortedSums, sortedOrigIdx, sortedQty, T);

  // ── 4/6  Reissues sheet ───────────────────────────────────
  ss.toast('4/6 — Generez reeditări...', '📦 Aprovizionare', -1);
  generateReissues_(ss, allRows, reissueMap, T);

  // ── 5/6  Supplier dashboard sheet ─────────────────────────
  ss.toast('5/6 — Generez dashboard furnizori...', '📦 Aprovizionare', -1);
  generateDashboard_(ss, filteredRows, extraData, T);

  // ── 6/6  Orders sheet + order list ───────────────────────
  ss.toast('6/6 — Generez comenzi...', '📦 Aprovizionare', -1);
  generateOrders_(ss, T);
  generateOrderList_(ss, T);

  // Enforce tab order: aprovizionare, comenzi, listă comenzi, reeditări, dashboard
  reorderSheets_(ss, T);

  T.lap('TOTAL DONE');
  T.flush(ss);
  ss.toast('✅ Gata! Verifică sheet-urile.', '📦 Aprovizionare', 10);
}

// Helper: find sheet by name using manual loop (avoids diacritics/Unicode issues)
function getSheetByName_(ss, name) {
  for (const s of ss.getSheets()) {
    if (s.getName() === name) return s;
  }
  return null;
}

// Move generated sheets to the desired tab order right after MAIN
function reorderSheets_(ss, T) {
  const sheetOrder = [
    APROVIZIONARE_SHEET_NAME,
    COMENZI_SHEET_NAME,
    LISTA_COMENZI_SHEET_NAME,
    REEDITARI_SHEET_NAME,
    DASHBOARD_SHEET_NAME
  ];
  const mainSheet = getSheetByName_(ss, MAIN_SHEET_NAME);
  const mainIdx   = mainSheet ? mainSheet.getIndex() : 1; // 1-based
  sheetOrder.forEach((name, i) => {
    const s = getSheetByName_(ss, name);
    if (s) {
      ss.setActiveSheet(s);
      ss.moveActiveSheet(mainIdx + 1 + i);
    }
  });
  T.lap('reordered tabs');
}

// ============================================================
// SHEET: listă comenzi (order log)
// Persistent log of exported order spreadsheets.
// Created once with headers; NEVER cleared on re-run — the list accumulates.
// New entries are always inserted at row 2 (most recent first).
// Columns: A=Data, B=Comandă (HYPERLINK), C=URL (plain, hidden — lookup key)
// ============================================================
function generateOrderList_(ss, T) {
  const { sheet, isNew } = getOrCreateSheet_(ss, LISTA_COMENZI_SHEET_NAME);
  if (!isNew) {
    T.lap('order list sheet already exists — preserved');
    return; // do not touch existing entries
  }

  const headers = ['Data', 'Comandă', 'URL'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#4472C4').setFontColor('white');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 140); // Date column
  sheet.setColumnWidth(2, 400); // Link column
  sheet.hideColumns(3);         // URL column hidden — used by script only
  T.lap('created order list sheet');
}

// ============================================================
// REISSUE MAP
// Groups ALL MAIN rows by Title|Author|Supplier.
// Only groups with >= 2 distinct EANs are kept.
// ============================================================
function buildReissueMap_(allRows) {
  const groupMap = {};
  for (let i = 0; i < allRows.length; i++) {
    const r = allRows[i];
    const key = String(r[C.Title]).trim()    + '|' +
                String(r[C.Author]).trim()   + '|' +
                String(r[C.Supplier]).trim();
    if (!groupMap[key]) groupMap[key] = { eans: [], rowIndices: [] };
    groupMap[key].eans.push(String(r[C.EAN]));
    groupMap[key].rowIndices.push(i);
  }
  const reissueGroups = {};
  for (const key in groupMap) {
    const uniqueEans = [...new Set(groupMap[key].eans)];
    if (uniqueEans.length >= 2) {
      reissueGroups[key] = {
        eans: uniqueEans,
        count: uniqueEans.length,
        rowIndices: groupMap[key].rowIndices
      };
    }
  }
  return reissueGroups;
}

// ============================================================
// REISSUE DEDUPLICATION
// From a list of already-filtered {r, origIdx} pairs, keeps only ONE row per
// reissue group — but ONLY if that row is also the globally most recent edition
// of the group (across ALL MAIN rows, not just the filtered ones).
// If the most recent edition didn't pass the filter, NO edition from the group
// is shown (older editions can't be ordered anyway).
// Non-reissue rows are always kept unchanged.
// ============================================================
function deduplicateReissues_(filteredPairs, reissueMap, allRows) {
  // Pass 1: find the globally most recent origIdx per reissue group (from allRows)
  const groupActiveOrigIdx = {};  // groupKey → origIdx of the most recent edition in MAIN
  for (let i = 0; i < allRows.length; i++) {
    const r = allRows[i];
    const key = String(r[C.Title]).trim() + '|' + String(r[C.Author]).trim() + '|' + String(r[C.Supplier]).trim();
    if (!reissueMap[key]) continue;
    const ts = parseDate_(r[C.PublishDate]);
    if (!groupActiveOrigIdx[key] || ts > groupActiveOrigIdx[key].ts) {
      groupActiveOrigIdx[key] = { origIdx: i, ts };
    }
  }
  // Pass 2: keep non-reissue rows + only the globally active edition (if it passed the filter)
  return filteredPairs.filter(({ r, origIdx }) => {
    const key = String(r[C.Title]).trim() + '|' + String(r[C.Author]).trim() + '|' + String(r[C.Supplier]).trim();
    if (!reissueMap[key]) return true;
    return origIdx === groupActiveOrigIdx[key].origIdx;
  });
}

// ============================================================
// REISSUE SALES SUMS
// For each reissue group, sums the sales metrics across ALL editions (allRows).
// Returns groupKey → {sumL1W, sumLM, sumL2M, sumLS, sumLY}
// ============================================================
function buildReissueSumMap_(allRows, reissueMap) {
  const sumMap = {};
  for (const key of Object.keys(reissueMap)) {
    const grp = reissueMap[key];
    let sumL1W = 0, sumLM = 0, sumL2M = 0, sumLS = 0, sumLY = 0;
    for (const ri of grp.rowIndices) {
      const r = allRows[ri];
      sumL1W += Number(r[C.SalesL1W]) || 0;
      sumLM  += Number(r[C.SalesLM])  || 0;
      sumL2M += Number(r[C.SalesL2M]) || 0;
      sumLS  += Number(r[C.SalesLS])  || 0;
      sumLY  += Number(r[C.SalesLY])  || 0;
    }
    sumMap[key] = { sumL1W, sumLM, sumL2M, sumLS, sumLY };
  }
  return sumMap;
}

// ============================================================
// COMPUTE EXTRA DATA — 8 flag/reissue columns per filtered row
// ============================================================
function computeExtraData_(filteredRows, reissueMap) {
  const recentCutoff = Date.now() - RECENT_TITLES * 24 * 60 * 60 * 1000;
  const extraRows = [];
  for (let i = 0; i < filteredRows.length; i++) {
    const r      = filteredRows[i];
    const extraRow = new Array(NUM_EXTRA).fill('');

    const groupKey = String(r[C.Title]).trim()    + '|' +
                     String(r[C.Author]).trim()   + '|' +
                     String(r[C.Supplier]).trim();
    const grp = reissueMap[groupKey];

    // Reissue columns
    if (grp) {
      extraRow[EX.REISSUE]       = 1;
      extraRow[EX.EDITION_COUNT] = grp.count;
      const myEan = String(r[C.EAN]);
      extraRow[EX.OTHER_EANS]    = grp.eans.filter(e => e !== myEan).join(', ');
    } else {
      extraRow[EX.REISSUE]       = 0;
      extraRow[EX.EDITION_COUNT] = 1;
      extraRow[EX.OTHER_EANS]    = '';
    }

    // Flag columns
    const rrp           = Number(r[C.RRP])           || 0;
    const purchasePrice = Number(r[C.PurchasePrice])  || 0;
    const salesL1W      = Number(r[C.SalesL1W])       || 0;
    const salesLM       = Number(r[C.SalesLM])        || 0;
    const salesL2M      = Number(r[C.SalesL2M])       || 0;
    const stock         = Number(r[C.OnlineStock])     || 0;
    // Mirror AVG_DAILY_SALES_FORMULA (not in MAIN EBS BRUT — calculated inline)
    const avgSales      = salesL1W / 7 * 0.6 + Math.max(0, (salesL2M - salesLM) / 30) * 0.4;
    const outOfPrint    = String(r[C.OutOfPrint]).trim().toLowerCase();
    const unavailable   = String(r[C.Unavailable]).trim().toLowerCase();

    // Bargain: purchase price < 20% of RRP
    extraRow[EX.BARGAIN]     = (rrp > 0 && purchasePrice > 0 && purchasePrice / rrp < 0.2) ? 1 : 0;
    // Sales spike: last-week sales > 50% of last-month sales
    extraRow[EX.SALES_SPIKE] = (salesLM > 0 && salesL1W > salesLM * 0.5) ? 1 : 0;
    // Overstock: stock > 100 and avg daily sales <= 2
    extraRow[EX.OVERSTOCK]   = (stock > 100 && avgSales <= 2) ? 1 : 0;
    // Stockout: no stock but selling well (avg > 5/day)
    extraRow[EX.STOCKOUT]    = (stock === 0 && avgSales > 5)  ? 1 : 0;

    // Zombie: out-of-print or unavailable, yet still has stock
    const isOutOfPrint  = (outOfPrint  === '1' || outOfPrint  === 'da' || outOfPrint  === 'true');
    const isUnavailable = (unavailable === '1' || unavailable === 'da' || unavailable === 'true');
    extraRow[EX.ZOMBIE]  = ((isOutOfPrint || isUnavailable) && stock > 0) ? 1 : 0;

    // Noutăți: SalesLY=0 and created within RECENT_TITLES days
    const salesLY  = Number(r[C.SalesLY]) || 0;
    const created  = parseDate_(r[C.PublishDate]);
    extraRow[EX.NOUTATI] = (salesLY === 0 && created > 0 && created >= recentCutoff) ? 1 : 0;

    extraRows.push(extraRow);
  }
  return extraRows;
}

// ============================================================
// SHEET: aprovizionare (procurement)
// 46 columns written by script in sorted order; errors cleaned.
// Layout: 37 derived cols (33 MAIN + 4 calculated: J=Necesar, K=Cantitate,
//          L=DOS, S=VZ medie/zi) followed by 9 extra flag cols (AL–AT).
// Calculated cols are highlighted with yellow background.
// All 46 cols are built in JS from sortedRows and written via one setValues call;
// then J, K, L, S are overwritten with live formula strings via setFormulas.
// ============================================================
function generateProcurement_(ss, sortedRows, sortedExtra, sortedSums, sortedOrigIdx, sortedQty, T) {
  const { sheet, isNew } = getOrCreateSheet_(ss, APROVIZIONARE_SHEET_NAME);
  T.lap(isNew ? 'created procurement sheet' : 'reusing procurement sheet');

  if (!isNew) {
    clearAllContent_(sheet);
    T.lap('clearAllContent procurement');
  }

  const allHeaders = [...APROV_HEADERS, ...EXTRA_HEADERS];
  sheet.getRange(1, 1, 1, allHeaders.length).setValues([allHeaders]);

  if (isNew) {
    sheet.getRange(1, 1, 1, allHeaders.length)
      .setFontWeight('bold').setBackground('#4472C4').setFontColor('white');
    sheet.setFrozenRows(1);
    T.lap('header formatting procurement');
  }

  if (sortedRows.length === 0) return;
  const n = sortedRows.length;

  // ── Step 1: Build complete output array in JS and write in one shot ────────
  // All 46 columns are assembled from sortedRows[i] (MAIN data) + sortedSums[i]
  // (reissue group sales sums) + sortedExtra[i] (flag columns).
  // J/K/L/S placeholders are written here so the sheet isn't empty while
  // setFormulas runs next; they are immediately overwritten by Step 2.
  const outputData = [];
  for (let i = 0; i < n; i++) {
    const r    = sortedRows[i];
    const sums = sortedSums[i];
    const ex   = sortedExtra[i];
    const N  = sums ? sums.sumL1W : (Number(r[C.SalesL1W]) || 0);
    const O  = sums ? sums.sumLM  : (Number(r[C.SalesLM])  || 0);
    const P  = sums ? sums.sumL2M : (Number(r[C.SalesL2M]) || 0);
    const Qs = sums ? sums.sumLS  : (Number(r[C.SalesLS])  || 0);
    const R  = sums ? sums.sumLY  : (Number(r[C.SalesLY])  || 0);
    const I  = Number(r[C.OnlineStock]) || 0;
    const qty  = sortedQty[i];
    const vzzi = N / 7 * 0.6 + Math.max(0, (P - O) / 30) * 0.4;
    const dos  = vzzi > 0 ? (qty + I) / vzzi : 0;
    outputData.push([
      r[0],r[1],r[2],r[3],r[4],r[5],r[6],r[7],r[8],  // A-I: MAIN[0..8]
      O - I, qty, dos,                                 // J,K,L (placeholder values)
      r[9],                                            // M: MAIN[9] = Disp. Rec.
      N, O, P, Qs, R,                                  // N-R: sales sums
      vzzi,                                            // S (placeholder value)
      r[15],r[16],r[17],r[18],r[19],r[20],r[21],r[22],// T-AA: MAIN[15..22]
      r[23],r[24],r[25],r[26],r[27],r[28],r[29],r[30],r[31],r[32], // AB-AK: MAIN[23..32]
      ...ex                                            // AL-AT: extra flags
    ]);
  }
  sheet.getRange(2, 1, n, allHeaders.length).setValues(outputData);
  T.lap('setValues aprovizionare (' + (n * allHeaders.length) + ' cells)');

  // ── Step 2: Overwrite J, K, L, S with live formula strings ─────────────────
  // setValues wrote placeholder numbers; replace with real formulas so the user
  // can inspect them and they auto-recalculate if MAIN is edited.
  {
    const reqF = [], qtyF = [], dosF = [], vzF = [];
    for (let i = 0; i < n; i++) {
      const row = i + 2;
      reqF.push([adjustFormulaRow_(REQUIRED_FORMULA,          row)]);
      qtyF.push([adjustFormulaRow_(QUANTITY_FORMULA,          row)]);
      dosF.push([adjustFormulaRow_(DOS_FORMULA,               row)]);
      vzF.push( [adjustFormulaRow_(AVG_DAILY_SALES_FORMULA,   row)]);
    }
    sheet.getRange(2, 10, n, 1).setFormulas(reqF);  // J: Necesar
    sheet.getRange(2, 11, n, 1).setFormulas(qtyF);  // K: Cantitate
    sheet.getRange(2, 12, n, 1).setFormulas(dosF);  // L: DOS
    sheet.getRange(2, 19, n, 1).setFormulas(vzF);   // S: VZ medie/zi
    T.lap('setFormulas J,K,L,S (' + n + ' rows)');
  }

  // ── Step 3: Formatting ───────────────────────────────────────
  const lastDataRow = n + 1;
  sheet.getRangeList([
    `I2:I${lastDataRow}`, `J2:J${lastDataRow}`, `K2:K${lastDataRow}`,
    `L2:L${lastDataRow}`, `N2:N${lastDataRow}`, `O2:O${lastDataRow}`,
    `P2:P${lastDataRow}`, `Q2:Q${lastDataRow}`, `R2:R${lastDataRow}`
  ]).setNumberFormat('0');
  sheet.getRange(2, 20, n, 1).setNumberFormat('dd.mm.yyyy');
  sheet.getRangeList([`J2:J${lastDataRow}`, `L2:L${lastDataRow}`, `S2:S${lastDataRow}`])
    .setBackground('#FFFDE7');
  sheet.getRange(2, 11, n, 1).setBackground('#d9ead3').setFontWeight('bold');
  T.lap('styling procurement');

  // ── Step 4: CF (every run) ───────────────────────────────────
  const cfFullRow   = sheet.getRange(2, 1, n, allHeaders.length);
  const noutatiCol  = colToLetter_(NUM_APROV_DERIVED + 1 + EX.NOUTATI);
  const bargainCol  = colToLetter_(NUM_APROV_DERIVED + 1 + EX.BARGAIN);
  const stockoutCol = colToLetter_(NUM_APROV_DERIVED + 1 + EX.STOCKOUT);
  const zombieCol   = colToLetter_(NUM_APROV_DERIVED + 1 + EX.ZOMBIE);
  sheet.setConditionalFormatRules([
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + noutatiCol + '2=1')
      .setBackground('#CDED74').setRanges([cfFullRow]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$AL2=1')
      .setBackground('#FFF2CC').setRanges([cfFullRow]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + bargainCol + '2=1')
      .setBackground('#C6EFCE')
      .setRanges([sheet.getRange(2, NUM_APROV_DERIVED + 1 + EX.BARGAIN, n, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + stockoutCol + '2=1')
      .setBackground('#FFC7CE')
      .setRanges([sheet.getRange(2, NUM_APROV_DERIVED + 1 + EX.STOCKOUT, n, 1)]).build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenFormulaSatisfied('=$' + zombieCol + '2=1')
      .setBackground('#FFE0B2')
      .setRanges([sheet.getRange(2, NUM_APROV_DERIVED + 1 + EX.ZOMBIE, n, 1)]).build()
  ]);
  T.lap('CF procurement');
}

// ============================================================
// SHEET: reeditări (reissues)
// Groups sorted desc by most recent PublishDate in the group.
// Within each group: most recent edition first (Activ=DA), older ones below (Activ=NU).
// ============================================================
function generateReissues_(ss, allRows, reissueMap, T) {
  const { sheet, isNew } = getOrCreateSheet_(ss, REEDITARI_SHEET_NAME);
  T.lap(isNew ? 'created reissues sheet' : 'reusing reissues sheet');

  if (!isNew) {
    clearAllContent_(sheet);
    T.lap('clearAllContent reissues');
  }

  const headers = [
    'Grup', 'EAN', 'Articol', 'Autor', 'Furnizor',
    'Data Publicare', 'Stoc', 'SalesL1W', 'SalesLM', 'SalesL2M', 'SalesLS', 'SalesLY', 'VZ/zi',
    'Epuizat', 'Indisponibil', 'Nr. Ediții', 'Activ'
  ];

  // Always write header row
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  if (isNew) {
    sheet.getRange(1, 1, 1, headers.length)
      .setFontWeight('bold').setBackground('#4472C4').setFontColor('white');
    sheet.setFrozenRows(1);
    T.lap('formatting reissues');
  }

  // Sort groups desc by max PublishDate in the group
  // (the group whose most recent edition was published last comes first)
  const groupMaxDate = {};
  for (const key of Object.keys(reissueMap)) {
    const grp = reissueMap[key];
    let maxTimestamp = 0;
    for (const ri of grp.rowIndices) {
      const t = parseDate_(allRows[ri][C.PublishDate]);
      if (t > maxTimestamp) maxTimestamp = t;
    }
    groupMaxDate[key] = maxTimestamp;
  }
  const sortedGroupKeys = Object.keys(reissueMap)
    .sort((a, b) => groupMaxDate[b] - groupMaxDate[a]);

  const outputRows = [];
  let groupIdx = 0;

  for (const key of sortedGroupKeys) {
    groupIdx++;
    const grp = reissueMap[key];

    // Sort editions within group desc by PublishDate → index 0 = active edition
    const groupRows = grp.rowIndices
      .map(ri => allRows[ri])
      .sort((a, b) => {
        const da = parseDate_(a[C.PublishDate]);
        const db = parseDate_(b[C.PublishDate]);
        return db - da;
      });

    for (let j = 0; j < groupRows.length; j++) {
      const r = groupRows[j];
      outputRows.push([
        groupIdx,
        r[C.EAN],
        r[C.Title],
        r[C.Author],
        r[C.Supplier],
        r[C.PublishDate],
        r[C.OnlineStock],
        r[C.SalesL1W],
        r[C.SalesLM],
        r[C.SalesL2M],
        r[C.SalesLS],
        r[C.SalesLY],
        // Mirror AVG_DAILY_SALES_FORMULA (not in MAIN EBS BRUT)
        Number(r[C.SalesL1W]) / 7 * 0.6 + Math.max(0, (Number(r[C.SalesL2M]) - Number(r[C.SalesLM])) / 30) * 0.4,
        r[C.OutOfPrint],
        r[C.Unavailable],
        grp.count,
        j === 0 ? 'DA' : 'NU'   // DA = most recent PublishDate in group
      ]);
    }
  }

  if (outputRows.length > 0) {
    sheet.getRange(2, 1, outputRows.length, headers.length).setValues(outputRows);
    T.lap('setValues reissues: ' + outputRows.length + ' rows');

    // Format publish date column (col 6 = F)
    sheet.getRange(2, 6, outputRows.length, 1).setNumberFormat('dd.mm.yyyy');

    // Integer format on numeric sales/stock columns:
    // G=7(Stoc), H=8(SalesL1W), I=9(SalesLM), J=10(SalesL2M), K=11(SalesLS), L=12(SalesLY)
    [7, 8, 9, 10, 11, 12].forEach(col1 =>
      sheet.getRange(2, col1, outputRows.length, 1).setNumberFormat('0')
    );

    // CF: active (DA) → light grey; inactive (NU) → muted grey text
    // Col Q (17) = Activ (updated: was N/14 before SalesL1W/L2M/LS were added)
    const range = sheet.getRange(2, 1, outputRows.length, headers.length);
    sheet.setConditionalFormatRules([
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$Q2="DA"')
        .setBackground('#d9d9d9').setFontColor('#000000').setRanges([range]).build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$Q2="NU"')
        .setBackground(null).setFontColor('#b7b7b7').setRanges([range]).build()
    ]);
    T.lap('formatting + CF reissues');
  }
}

// ============================================================
// SHEET: dashboard furnizori (supplier dashboard)
// Aggregated KPIs per supplier, sorted desc by Total Required
// ============================================================
function generateDashboard_(ss, filteredRows, extraData, T) {
  const { sheet, isNew } = getOrCreateSheet_(ss, DASHBOARD_SHEET_NAME);
  T.lap(isNew ? 'created dashboard sheet' : 'reusing dashboard sheet');

  if (!isNew) {
    clearAllContent_(sheet);
    T.lap('clearAllContent dashboard');
  }

  const headers = [
    'Activ', 'Total', 'Furnizor', 'Nr. SKU', 'Stoc Total',
    'SalesLM Total', 'SalesLY Total', 'Nr. Reeditări',
    'Nr. Chilipir', 'Nr. Rupturi', 'Nr. Zombie', 'Nr. Suprastoc'
  ];

  // Always write header row — full row A:Z styled uniformly
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange('1:1')
    .setFontWeight('bold').setBackground('#4472C4').setFontColor('white');
  // Center header A (Activ)
  sheet.getRange('A1').setHorizontalAlignment('center');

  if (isNew) {
    sheet.setFrozenRows(1);
    T.lap('formatting dashboard');
  }

  // Aggregate per supplier
  const supplierMap = {};
  for (let i = 0; i < filteredRows.length; i++) {
    const r        = filteredRows[i];
    const extraRow = extraData[i];
    const supplier = String(r[C.Supplier]).trim();
    if (!supplierMap[supplier]) {
      supplierMap[supplier] = {
        sku: 0, stock: 0, salesLM: 0, salesLY: 0, required: 0,
        reissues: 0, bargains: 0, stockouts: 0, zombies: 0, overstocks: 0
      };
    }
    const m = supplierMap[supplier];
    m.sku++;
    m.stock      += Number(r[C.OnlineStock]) || 0;
    m.salesLM    += Number(r[C.SalesLM])     || 0;
    m.salesLY    += Number(r[C.SalesLY])     || 0;
    // Necesar = SalesLM − OnlineStock (calculated, not in MAIN EBS BRUT)
    m.required   += Math.max(0, (Number(r[C.SalesLM]) || 0) - (Number(r[C.OnlineStock]) || 0));
    m.reissues   += extraRow[EX.REISSUE]       ? 1 : 0;
    m.bargains   += extraRow[EX.BARGAIN]       ? 1 : 0;
    m.stockouts  += extraRow[EX.STOCKOUT]      ? 1 : 0;
    m.zombies    += extraRow[EX.ZOMBIE]        ? 1 : 0;
    m.overstocks += extraRow[EX.OVERSTOCK]     ? 1 : 0;
  }

  // Build rows sorted desc by Total Required (col index 1 after prepending Activ)
  const dashboardRows = Object.keys(supplierMap)
    .map(supplier => {
      const m = supplierMap[supplier];
      return [false, m.required, supplier, m.sku, m.stock, m.salesLM, m.salesLY,
              m.reissues, m.bargains, m.stockouts, m.zombies, m.overstocks];
    })
    .sort((a, b) => b[1] - a[1]);

  if (dashboardRows.length > 0) {
    sheet.getRange(2, 1, dashboardRows.length, headers.length).setValues(dashboardRows);
    T.lap('setValues dashboard: ' + dashboardRows.length + ' suppliers');

    // Checkbox data validation on col A (Activ)
    const checkboxRule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
    sheet.getRange(2, 1, dashboardRows.length, 1).setDataValidation(checkboxRule)
      .setHorizontalAlignment('center');

    // CF: text color based on Activ checkbox — applied every run
    const cfRange = sheet.getRange(2, 1, dashboardRows.length, headers.length);
    const cfRules = [
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$A2=TRUE')
        .setFontColor('#000000')
        .setRanges([cfRange])
        .build(),
      SpreadsheetApp.newConditionalFormatRule()
        .whenFormulaSatisfied('=$A2=FALSE')
        .setFontColor('#999999')
        .setRanges([cfRange])
        .build()
    ];
    sheet.setConditionalFormatRules(cfRules);
    T.lap('checkbox + CF dashboard');

    // Auto-resize all columns to fit content
    sheet.autoResizeColumns(1, headers.length);
    T.lap('autoResizeColumns dashboard');
  }
}

// ============================================================
// SHEET: comenzi (orders)
// Row 1: C=label "Furnizor:", D=supplier dropdown
// Row 3: column headers | Row 4+: order data
// ============================================================
function generateOrders_(ss, T) {
  const { sheet, isNew } = getOrCreateSheet_(ss, COMENZI_SHEET_NAME);
  T.lap(isNew ? 'created orders sheet' : 'reusing orders sheet');

  if (!isNew) {
    clearAllContent_(sheet);
    T.lap('clearAllContent orders');
  }

  // Row 1: controls — clear old dropdown validation before setting placeholder
  sheet.getRange('D1').clearDataValidations();
  sheet.getRange('C1').setValue('Furnizor:');
  sheet.getRange('D1').setValue('— alege furnizor —');

  // Row 3: column headers
  sheet.getRange(3, 1, 1, ORDER_HEADERS.length).setValues([ORDER_HEADERS]);

  if (isNew) {
    sheet.getRange('C1').setFontWeight('bold');
    sheet.getRange(3, 1, 1, ORDER_HEADERS.length)
      .setFontWeight('bold').setBackground('#4472C4').setFontColor('white');
    sheet.setFrozenRows(3);
    T.lap('formatting orders');
  }

  // Build dropdown: sorted alphabetically, with total quantity in parentheses.
  const procSheet = getSheetByName_(ss, APROVIZIONARE_SHEET_NAME);
  if (procSheet) {
    const lastRow = procSheet.getLastRow();
    if (lastRow > 1) {
      const procData = procSheet.getRange(2, 1, lastRow - 1, NUM_APROV_DERIVED).getValues();
      const supplierTotals = {};
      for (const r of procData) {
        const sup = String(r[C.Supplier]).trim();
        if (sup) supplierTotals[sup] = (supplierTotals[sup] || 0) + (Number(r[APROV_CANTITATE]) || 0);
      }
      const dropdownValues = Object.keys(supplierTotals)
        .sort()  // alphabetical
        .map(s => s + ' (' + Math.round(supplierTotals[s]) + ')');
      if (dropdownValues.length > 0) {
        sheet.getRange('D1').setDataValidation(
          SpreadsheetApp.newDataValidation()
            .requireValueInList(dropdownValues, true)
            .setAllowInvalid(false)
            .build()
        );
        T.lap('supplier dropdown: ' + dropdownValues.length + ' values (alphabetical)');
      }
    }
  }
}

// ============================================================
// onEdit trigger — populate orders when supplier dropdown (D1) changes
// ============================================================
function onEdit(e) {
  if (!e) return;
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== COMENZI_SHEET_NAME) return;
  if (e.range.getRow() !== 1 || e.range.getColumn() !== 4) return;

  const rawValue = e.value;

  // Show loading indicator in D2 first, then clear old rows
  const statusCell = sheet.getRange('D2');
  statusCell.setValue('⏳ Generare comandă...').setBackground('#FFF9C4');

  const lastRow = sheet.getLastRow();
  if (lastRow >= 4) {
    sheet.getRange(4, 1, lastRow - 3, ORDER_HEADERS.length).clearContent();
  }
  SpreadsheetApp.flush();

  if (!rawValue || rawValue === '— alege furnizor —') {
    statusCell.clearContent().setBackground(null);
    return;
  }

  populateOrdersFromProcurement_(e.source, sheet, extractSupplierName_(rawValue));

  // Clear loading indicator
  statusCell.clearContent().setBackground(null);
  SpreadsheetApp.flush();
}

// Strip the "(N)" total suffix from a dropdown value to get the clean supplier name
// e.g. "GRUP EDITORIAL LITERA SRL (10234)" → "GRUP EDITORIAL LITERA SRL"
function extractSupplierName_(dropdownValue) {
  return String(dropdownValue).replace(/\s*\(\d+\)\s*$/, '').trim();
}

// ============================================================
// Populate orders sheet from procurement data
// Shows only rows where Quantity > 0 (= Cant. pre-filled with Quantity)
// ============================================================
function populateOrdersFromProcurement_(ss, ordersSheet, supplier) {
  const T = makeTimer();

  const procSheet = getSheetByName_(ss, APROVIZIONARE_SHEET_NAME);
  if (!procSheet) return;

  const lastRow = procSheet.getLastRow();
  if (lastRow < 2) return;

  const procData = procSheet.getRange(2, 1, lastRow - 1, NUM_APROV_DERIVED + NUM_EXTRA).getValues();
  T.lap('read procurement: ' + procData.length + ' rows');

  // supplier arg is already a clean name (no "(N)" suffix)
  // Cols from aprovizionare (0-based): Supplier=5, Cantitate=APROV_CANTITATE(10)
  const supplierRows = procData
    .filter(r =>
      String(r[C.Supplier]).trim() === supplier &&
      Number(r[APROV_CANTITATE]) > 0
    )
    .sort((a, b) => (Number(b[APROV_CANTITATE]) || 0) - (Number(a[APROV_CANTITATE]) || 0));
  T.lap('filtered "' + supplier + '" Cantitate>0: ' + supplierRows.length + ' rows');

  // Clear previous order data (row 4+), keep header and dropdown rows
  const oldLastRow = ordersSheet.getLastRow();
  if (oldLastRow >= 4) {
    ordersSheet.getRange(4, 1, oldLastRow - 3, ORDER_HEADERS.length).clearContent();
  }

  if (supplierRows.length === 0) return;

  // Map to order columns: ElefantSKU, EAN, Cod articol, Articol, Autor, Furnizor, RRP, Reducere, Cantitate
  // Cols from aprovizionare (0-based): ElefantSKU=0, EAN=1, CodArticol=2, Title=3,
  //   Author=4, Supplier=5, RRP=6, Discount=7, Cantitate=APROV_CANTITATE(10)
  const orderRows = supplierRows.map(r => [
    r[0],                                    // ElefantSKU
    r[C.EAN],                                // EAN
    r[2],                                    // Cod articol
    r[C.Title],                              // Articol
    r[C.Author],                             // Autor
    r[C.Supplier],                           // Furnizor
    r[C.RRP],                                // RRP
    r[C.Discount],                           // Reducere
    Number(r[APROV_CANTITATE]) || 0          // Cantitate
  ]);

  ordersSheet.getRange(4, 1, orderRows.length, ORDER_HEADERS.length).setValues(orderRows);
  T.lap('setValues orders: ' + orderRows.length + ' rows');

  // Highlight "Cantitate" column (col 9) as editable; format as integer
  ordersSheet.getRange(4, 9, orderRows.length, 1)
    .setBackground('#E2EFDA').setFontWeight('bold').setNumberFormat('0');
  T.lap('highlight Cantitate column');
}

// ============================================================
// EXPORT (single supplier) — triggered from menu or called directly.
// Reads data from the comenzi sheet (rows 4+, as populated by dropdown).
// ============================================================
function exportOrdersToNewSpreadsheet_() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = getSheetByName_(ss, COMENZI_SHEET_NAME);
  if (!ordersSheet) {
    SpreadsheetApp.getUi().alert('Sheet-ul comenzi nu există. Rulează mai întâi scriptul de generare.');
    return;
  }

  const rawValue = String(ordersSheet.getRange('D1').getValue());
  if (!rawValue || rawValue === '— alege furnizor —') {
    SpreadsheetApp.getUi().alert('Selectează mai întâi un furnizor din dropdown-ul D1.');
    return;
  }

  const supplierName = extractSupplierName_(rawValue);
  const lastRow      = ordersSheet.getLastRow();
  if (lastRow < 4) {
    SpreadsheetApp.getUi().alert('Nu există date de exportat pentru furnizorul selectat.');
    return;
  }

  const headerValues = ordersSheet.getRange(3, 1, 1, ORDER_HEADERS.length).getValues();
  const dataValues   = ordersSheet
    .getRange(4, 1, lastRow - 3, ORDER_HEADERS.length)
    .getValues()
    .filter(r => r.some(v => v !== ''));

  if (dataValues.length === 0) {
    SpreadsheetApp.getUi().alert('Nu există date de exportat pentru furnizorul selectat.');
    return;
  }

  doExportSupplier_(ss, supplierName, headerValues, dataValues);

  const listaSheet = getSheetByName_(ss, LISTA_COMENZI_SHEET_NAME);
  if (listaSheet) {
    ss.setActiveSheet(listaSheet);
    listaSheet.getRange(2, 2).activate();
  }
}

// ============================================================
// EXPORT (all suppliers) — reads aprovizionare directly; does NOT
// depend on comenzi sheet state or dropdown selection.
// Skips suppliers with no rows in aprovizionare (user may have
// deleted rows after last generateSheets run — no error).
// ============================================================
function exportAllOrders_() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const dashSheet    = getSheetByName_(ss, DASHBOARD_SHEET_NAME);
  const procSheet    = getSheetByName_(ss, APROVIZIONARE_SHEET_NAME);

  if (!procSheet) {
    SpreadsheetApp.getUi().alert('Sheet-ul aprovizionare nu există. Rulează mai întâi scriptul de generare.');
    return;
  }

  // Collect checked suppliers from dashboard col A (Activ) + col C (Furnizor)
  let checkedSuppliers = null;
  if (dashSheet) {
    const dashLastRow = dashSheet.getLastRow();
    if (dashLastRow >= 2) {
      const dashData = dashSheet.getRange(2, 1, dashLastRow - 1, 3).getValues();
      // col A (idx 0) = Activ checkbox, col C (idx 2) = Furnizor
      checkedSuppliers = new Set(
        dashData
          .filter(row => row[0] === true)
          .map(row => String(row[2]).trim())
          .filter(s => s.length > 0)
      );
    }
  }

  if (checkedSuppliers !== null && checkedSuppliers.size === 0) {
    SpreadsheetApp.getUi().alert('Niciun furnizor bifat în dashboard. Bifează cel puțin un furnizor din coloana Activ.');
    return;
  }

  const lastRow = procSheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('Nu există date în aprovizionare.');
    return;
  }

  const procData = procSheet.getRange(2, 1, lastRow - 1, NUM_APROV_DERIVED + NUM_EXTRA).getValues();

  // Group rows by supplier (only Cantitate > 0, only checked suppliers)
  const supplierRowsMap = {};
  for (const r of procData) {
    const sup = String(r[C.Supplier]).trim();
    if (!sup || !(Number(r[APROV_CANTITATE]) > 0)) continue;
    if (checkedSuppliers !== null && !checkedSuppliers.has(sup)) continue;
    if (!supplierRowsMap[sup]) supplierRowsMap[sup] = [];
    supplierRowsMap[sup].push(r);
  }

  const suppliers = Object.keys(supplierRowsMap).sort();
  if (suppliers.length === 0) {
    SpreadsheetApp.getUi().alert('Nu există comenzi de exportat (Cantitate = 0 pentru furnizorii bifați).');
    return;
  }

  const headerValues = [ORDER_HEADERS];

  for (let i = 0; i < suppliers.length; i++) {
    const supplier = suppliers[i];
    ss.toast((i + 1) + ' / ' + suppliers.length + ' — ' + supplier,
             '📦 Export comenzi furnizori bifați', -1);

    const orderRows = supplierRowsMap[supplier]
      .sort((a, b) => (Number(b[APROV_CANTITATE]) || 0) - (Number(a[APROV_CANTITATE]) || 0))
      .map(r => [
        r[0],                                // ElefantSKU
        r[C.EAN],                            // EAN
        r[2],                                // Cod articol
        r[C.Title],                          // Articol
        r[C.Author],                         // Autor
        r[C.Supplier],                       // Furnizor
        r[C.RRP],                            // RRP
        r[C.Discount],                       // Reducere
        Number(r[APROV_CANTITATE]) || 0      // Cantitate
      ]);

    doExportSupplier_(ss, supplier, headerValues, orderRows);
  }

  ss.toast('\u2705 Exportate ' + suppliers.length + ' comenzi.', '📦 Aprovizionare', 10);

  const listaSheet = getSheetByName_(ss, LISTA_COMENZI_SHEET_NAME);
  if (listaSheet) {
    ss.setActiveSheet(listaSheet);
    listaSheet.getRange(2, 2).activate();
  }
}

// ============================================================
// EXPORT CORE — creates/appends a Drive spreadsheet for one supplier
// and updates "listă comenzi". Used by both single and bulk export.
//
// Logic:
//   1. Search Drive for "{supplier} - {date}" (not in Trash).
//      If multiple found, pick the one created most recently.
//   2. If found → insertSheet("Comanda HH:MM"); else create new file.
//   3. Write headers + data, freeze row 1, auto-resize columns.
//   4. setSharing: anyone in the domain with link can edit.
//   5. Update "listă comenzi" (upsert by base spreadsheet URL in col C).
// ============================================================
function doExportSupplier_(ss, supplierName, headerValues, dataValues) {
  const now      = new Date();
  const dateStr  = now.getFullYear()                            + '-' +
                   String(now.getMonth() + 1).padStart(2, '0') + '-' +
                   String(now.getDate()).padStart(2, '0');
  const timeStr  = String(now.getHours()).padStart(2, '0')   + ':' +
                   String(now.getMinutes()).padStart(2, '0');
  const filename = supplierName + ' - ' + dateStr;
  const tabName  = 'Comanda ' + timeStr;

  // ── Find or create Drive spreadsheet ─────────────────────────────────
  let targetSS   = null;
  let latestTime = 0;
  const driveFiles = DriveApp.getFilesByName(filename);
  while (driveFiles.hasNext()) {
    const f = driveFiles.next();
    if (f.isTrashed()) continue;
    const ct = f.getDateCreated().getTime();
    if (ct > latestTime) {
      latestTime = ct;
      try { targetSS = SpreadsheetApp.openById(f.getId()); } catch (_) { targetSS = null; }
    }
  }

  let newSheet;
  if (targetSS) {
    newSheet = targetSS.insertSheet(tabName, 0);
  } else {
    targetSS = SpreadsheetApp.create(filename);
    newSheet  = targetSS.getActiveSheet();
    newSheet.setName(tabName);
  }

  // ── Write content ─────────────────────────────────────────────────────
  newSheet.getRange(1, 1, 1, ORDER_HEADERS.length)
    .setValues(headerValues)
    .setFontWeight('bold').setBackground('#4472C4').setFontColor('white');
  newSheet.getRange(2, 1, dataValues.length, ORDER_HEADERS.length).setValues(dataValues);
  newSheet.setFrozenRows(1);
  newSheet.autoResizeColumns(1, ORDER_HEADERS.length);

  // setSharing may be blocked by Workspace domain policy — don't let it abort the export.
  try {
    DriveApp.getFileById(targetSS.getId())
      .setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
  } catch (e) {
    ss.toast('⚠️ Sharing nesetat pt "' + supplierName + '": ' + e.message, 'Aprovizionare', 8);
  }

  const baseUrl  = 'https://docs.google.com/spreadsheets/d/' + targetSS.getId() + '/edit';
  const sheetUrl = baseUrl + '#gid=' + newSheet.getSheetId();

  // ── Update "listă comenzi" (upsert by baseUrl) ────────────────────────
  const listaSheet = getSheetByName_(ss, LISTA_COMENZI_SHEET_NAME);
  if (listaSheet) {
    const linkLabel = filename + ' / ' + tabName;
    let existingRow = -1;
    if (listaSheet.getLastRow() > 1) {
      const listaData = listaSheet.getRange(2, 3, listaSheet.getLastRow() - 1, 1).getValues();
      for (let i = 0; i < listaData.length; i++) {
        if (String(listaData[i][0]).trim() === baseUrl) { existingRow = i + 2; break; }
      }
    }
    if (existingRow > 0) listaSheet.deleteRow(existingRow);
    listaSheet.insertRowBefore(2);
    listaSheet.getRange(2, 1).setValue(dateStr);
    listaSheet.getRange(2, 2).setFormula('=HYPERLINK("' + sheetUrl + '","' + linkLabel + '")');
    listaSheet.getRange(2, 3).setValue(baseUrl);
  }
}
// v2.45 (2026-04-09 14:30:00)
