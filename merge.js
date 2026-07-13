// ═══════════════════════════════════════════════════════════════════════════════
// CONFIG
// ═══════════════════════════════════════════════════════════════════════════════

const INPUT_SHEET  = "Input";
const OUTPUT_SHEET = "Output";
const MODULE_NAME  = "POS";
const OPTION_TYPE  = "Reports";
const TITLE        = "POS Reports";

const FC = 3, LC = 9, HEADER_ROW = 4, BODY_START = 5;

const HEADERS       = ["Sl. No", "Module", "Option Type", "Option Name", "Parameter Type", "Parameter Value", "Field Heading of Report"];
const INPUT_HEADERS = ["SL.NO", "Report Name", "Parameters"];
const COL_WIDTHS    = [50, 100, 120, 220, 130, 150, 450];

const DARK_BLUE  = "#123687";
const LIGHT_BLUE = "#99BFF2";
const GREY       = "#CCCCCC";
const WHITE      = "#FFFFFF";
const BS         = SpreadsheetApp.BorderStyle;

// ═══════════════════════════════════════════════════════════════════════════════
// MENU
// ═══════════════════════════════════════════════════════════════════════════════

function onOpen() {
  SpreadsheetApp.getUi().createMenu("Scripts").addItem("Regenerate Output", "regenerate").addToUi();
  ensureSheets();
}

// ═══════════════════════════════════════════════════════════════════════════════
// HELPERS
// ═══════════════════════════════════════════════════════════════════════════════

const ss  = () => SpreadsheetApp.getActive();
const key = (r, t, v) => `${r}||${t}||${v}`;

function ensureSheets() {
  const s = ss();
  if (!s.getSheetByName(OUTPUT_SHEET)) s.insertSheet(OUTPUT_SHEET);

  let inSheet = s.getSheetByName(INPUT_SHEET);
  if (inSheet) return;

  inSheet = s.insertSheet(INPUT_SHEET);
  inSheet.getRange(1, 1, 1, 3).setValues([INPUT_HEADERS])
    .setBackground(DARK_BLUE).setFontColor(WHITE).setFontWeight("bold").setHorizontalAlignment("center");
  [60, 250, 600].forEach((w, i) => inSheet.setColumnWidth(i + 1, w));
  inSheet.setFrozenRows(1);
}

/**
 * Parse "Type:Val1, Val2|Other:A, B" into
 * [{ type, values: [...] }, ...]
 *
 * Values are always split on commas so:
 *   Report type:B2B, B2C  →  type "Report type", values ["B2B","B2C"]
 */
function parseParameters(str) {
  if (!String(str || "").trim()) return [{ type: "", values: [""] }];

  return String(str).split("|").map(chunk => {
    chunk = chunk.trim();
    if (!chunk) return { type: "", values: [""] };

    const idx = chunk.indexOf(":");
    if (idx === -1) return { type: "", values: [chunk] };

    const type = chunk.slice(0, idx).trim();
    const raw  = chunk.slice(idx + 1);

    // Split on comma; trim; drop empties
    const values = raw.split(",")
      .map(v => v.trim())
      .filter(Boolean);

    return { type, values: values.length ? values : [""] };
  });
}

function readInput() {
  const sh = ss().getSheetByName(INPUT_SHEET);
  const last = sh.getLastRow();
  if (last < 2) return [];

  const reports = [];
  // getRange(row, col, numRows, numCols) — numRows = last-1 reads rows 2..last
  sh.getRange(2, 1, last - 1, 3).getValues().forEach(r => {
    const name = String(r[1] || "").trim();
    if (!name) return;
    reports.push({
      sl_no: String(r[0] || reports.length + 1).trim(),
      name,
      parameter_types: parseParameters(r[2])
    });
  });
  return reports;
}

function readExistingDescriptions() {
  const sh = ss().getSheetByName(OUTPUT_SHEET);
  const last = sh.getLastRow();
  if (last < BODY_START) return {};

  const vals = sh.getRange(BODY_START, FC, last - BODY_START + 1, LC - FC + 1).getValues();
  const map = {};
  let curReport = "", curType = "";

  vals.forEach(r => {
    if (r[3]) curReport = String(r[3]).trim();
    if (r[4] !== "") curType = String(r[4]).trim();
    const val  = String(r[5] || "").trim();
    const desc = String(r[6] || "").trim();
    if (curReport && desc) map[key(curReport, curType, val)] = desc;
  });
  return map;
}

// ═══════════════════════════════════════════════════════════════════════════════
// MAIN
// ═══════════════════════════════════════════════════════════════════════════════

function regenerate() {
  ensureSheets();

  const oldDescs = readExistingDescriptions();
  const reports  = readInput();

  if (!reports.length) return ss().toast("Input sheet is empty.");

  const sh = ss().getSheetByName(OUTPUT_SHEET);
  sh.getRange(1, 1, sh.getMaxRows(), sh.getMaxColumns()).breakApart();
  sh.clear();

  // ── Build rows + track merges + section bounds ───────────────────────────
  const rows = [], merges = [], sectionStarts = [], sectionEnds = [];
  let r = BODY_START;

  reports.forEach(rep => {
    const repStart = r;
    const total = rep.parameter_types.reduce((s, pt) => s + pt.values.length, 0);

    // Merge Sl.No / Module / Option Type / Option Name across all rows of this report
    if (total > 1) {
      for (let c = FC; c <= FC + 3; c++) merges.push([r, c, total, 1]);
    }

    let firstRow = true;
    rep.parameter_types.forEach(pt => {
      // Merge Parameter Type across its values (e.g. B2B + B2C)
      if (pt.values.length > 1) {
        merges.push([r, FC + 4, pt.values.length, 1]);
      }

      pt.values.forEach((val, i) => {
        const left = firstRow
          ? [rep.sl_no, MODULE_NAME, OPTION_TYPE, rep.name]
          : ["", "", "", ""];
        rows.push([
          ...left,
          i === 0 ? pt.type : "",
          val,
          oldDescs[key(rep.name, pt.type, val)] || ""
        ]);
        firstRow = false;
        r++;
      });
    });

    sectionStarts.push(repStart);
    sectionEnds.push(r - 1);
  });

  const lastRow = r - 1;
  const nCols = LC - FC + 1;

  // ── Write everything ─────────────────────────────────────────────────────
  sh.getRange(1, FC).setValue(TITLE).setFontWeight("bold").setFontSize(16).setFontColor(DARK_BLUE);
  sh.getRange(HEADER_ROW, FC, 1, nCols).setValues([HEADERS]);
  sh.getRange(BODY_START, FC, rows.length, nCols).setValues(rows);
  merges.forEach(m => sh.getRange(...m).merge());

  // ── Header formatting ────────────────────────────────────────────────────
  sh.getRange(HEADER_ROW, FC, 1, nCols)
    .setBackground(DARK_BLUE).setFontColor(WHITE).setFontWeight("bold").setFontSize(10)
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // ── Body formatting ──────────────────────────────────────────────────────
  sh.getRange(BODY_START, FC, rows.length, nCols)
    .setFontSize(9).setHorizontalAlignment("center").setVerticalAlignment("middle").setWrap(true);
  sh.getRange(BODY_START, LC, rows.length, 1).setHorizontalAlignment("left");

  // ── Column widths ────────────────────────────────────────────────────────
  COL_WIDTHS.forEach((w, i) => sh.setColumnWidth(FC + i, w));

  // ── Borders ──────────────────────────────────────────────────────────────
  const tableHeight = lastRow - HEADER_ROW + 1;
  const table = sh.getRange(HEADER_ROW, FC, tableHeight, nCols);

  // 1. Base grey grid (outer + all inner verticals + horizontals)
  table.setBorder(true, true, true, true, true, true, GREY, BS.SOLID);

  // 2. Light-blue vertical separators between columns
  for (let c = FC; c < LC; c++) {
    sh.getRange(HEADER_ROW, c, tableHeight, 1)
      .setBorder(null, null, null, true, null, null, LIGHT_BLUE, BS.SOLID_MEDIUM);
  }

  // 3. Light-blue line under header
  sh.getRange(HEADER_ROW, FC, 1, nCols)
    .setBorder(null, null, true, null, null, null, LIGHT_BLUE, BS.SOLID_MEDIUM);

  // 4. Clear internal horizontal lines inside multi-row merges
  //    (Sl.No→Option Name, and Parameter Type) so grey lines don't cut through merges.
  //    Parameter Value / Field Heading keep their row separators.
  merges.forEach(([row, col, numRows]) => {
    if (numRows > 1) {
      sh.getRange(row, col, numRows, 1)
        .setBorder(null, null, null, null, null, false);
    }
  });

  // 5. Light-blue separator under each report block (full width).
  //    CRITICAL: use the FULL block range, not just the last row.
  //    Merged cells (left 4 cols AND Parameter Type) only get a reliable bottom
  //    border when bottom is applied on the whole merge/block range.
  //    Example: Report type:B2B, B2C  → Parameter Type is merged 2 rows deep.
  sectionEnds.forEach((endRow, i) => {
    if (endRow >= lastRow) return;
    const startRow = sectionStarts[i];
    const height   = endRow - startRow + 1;
    sh.getRange(startRow, FC, height, nCols)
      .setBorder(null, null, true, null, null, null, LIGHT_BLUE, BS.SOLID_MEDIUM);
  });

  // 6. Thick dark-blue outer border
  table.setBorder(true, true, true, true, null, null, DARK_BLUE, BS.SOLID_THICK);

  // 7. Thick dark-blue header bottom + header verticals
  sh.getRange(HEADER_ROW, FC, 1, nCols)
    .setBorder(null, null, true, null, true, false, DARK_BLUE, BS.SOLID_THICK);

  // ── Finalize ─────────────────────────────────────────────────────────────
  sh.setHiddenGridlines(true);
  sh.setFrozenRows(HEADER_ROW);

  ss().toast(`✅ Output regenerated with ${reports.length} reports.`);
}
