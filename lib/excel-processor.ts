import ExcelJS from "exceljs";

export interface WorkerEntry {
  address: string;
  hours: number;
  materials: number;
  gas: number;
  ticket: number;
}

export interface WorkerBlock {
  name: string;
  rate: number;
  entries: WorkerEntry[];
}

export interface CostSummary {
  labor: number;
  materials: number;
  gas: number;
  other: number;
  total: number;
}

export interface WorkerTotal extends CostSummary {
  name: string;
  rate: number;
}

export interface SiteTotal extends CostSummary {
  address: string;
}

export interface ProcessResult {
  workers: WorkerBlock[];
  matchedSheets: string[];
  unmatchedAddresses: string[];
  rowsAdded: number;
  warnings: string[];
  workerTotals: WorkerTotal[];
  siteTotals: SiteTotal[];
}

export const MAX_FILE_SIZE = 50 * 1024 * 1024; // 50 MB

// --- Column index constants ---

/** Input file column indices (Sheet1 of input cost sheet) */
const IN_COL = {
  NAME: 2,      // B — Worker name
  RATE: 3,      // C — $/hr rate
  ADDRESS: 4,   // D — Address code
  HOURS: 5,     // E — Hours
  MATERIALS: 7, // G — Material cost
  GAS: 8,       // H — Gas cost
  TICKET: 9,    // I — Ticket cost
} as const;

/** Output file column indices (date-range sheets with project blocks) */
const OUT_COL = {
  LABEL: 1,     // A — 单号 / payment labels / 总开销
  AMOUNT: 2,    // B — Payment amounts / summary formulas
  PROJECT: 3,   // C — Project name
  DATE: 4,      // D — 日期
  WORKER: 5,    // E — 工人
  HOURS: 6,     // F — 工时
  RATE: 7,      // G — Hourly rate
  COST: 8,      // H — F*G formula (开销)
  OTHER: 9,     // I — 其他开销
  MATERIAL: 11, // K — 材料开销
  NOTES: 12,    // L — Notes
  MAT_DATE: 13, // M — Material date label
} as const;

const CENTURY_PREFIX = 2000;

/** Convert a 1-based column index to its Excel letter (1→A, 8→H, 11→K, etc.) */
function colLetter(index: number): string {
  let result = "";
  let n = index;
  while (n > 0) {
    n--;
    result = String.fromCharCode(65 + (n % 26)) + result;
    n = Math.floor(n / 26);
  }
  return result;
}

// Precomputed column letters for formula strings
const COL = {
  HOURS: colLetter(OUT_COL.HOURS),       // F
  RATE: colLetter(OUT_COL.RATE),          // G
  COST: colLetter(OUT_COL.COST),          // H
  OTHER: colLetter(OUT_COL.OTHER),        // I
  MATERIAL: colLetter(OUT_COL.MATERIAL),  // K
} as const;

// --- Types for date-range sheet format ---

interface ProjectBlock {
  dataStartRow: number; // first data row (header + 1)
  totalExpenseRow: number; // 总开销 row
  totalPriceRow: number; // 总价格 row
  profitRow: number; // 利润率 row
  projectNames: string[]; // non-empty col C values from data rows
  colAValues: string[]; // non-empty col A values from data rows (for fallback matching)
}

interface DateRange {
  month: number;
  year: number;
}

// --- Cell helpers ---

function getCellNumber(cell: ExcelJS.Cell): number {
  const val = cell.value;
  if (val == null) return 0;
  if (typeof val === "number") return val;
  if (typeof val === "object" && "result" in val) {
    const result = (val as ExcelJS.CellFormulaValue).result;
    if (typeof result === "number") return result;
    if (typeof result === "string") {
      const n = parseFloat(result);
      return isNaN(n) ? 0 : n;
    }
  }
  if (typeof val === "string") {
    const n = parseFloat(val);
    return isNaN(n) ? 0 : n;
  }
  return 0;
}

function getCellString(cell: ExcelJS.Cell): string {
  const val = cell.value;
  if (val == null) return "";
  if (typeof val === "string") return val.trim();
  if (typeof val === "number") return String(val);
  if (typeof val === "object" && "result" in val) {
    return String((val as ExcelJS.CellFormulaValue).result ?? "");
  }
  return String(val);
}

async function loadWorkbook(buffer: ArrayBuffer): Promise<ExcelJS.Workbook> {
  const wb = new ExcelJS.Workbook();
  // ExcelJS types don't align with ArrayBuffer in TS 6 — cast required
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  await (wb.xlsx as any).load(buffer);
  return wb;
}

// --- Input parsing ---

export function parseInputFile(worksheet: ExcelJS.Worksheet): { workers: WorkerBlock[]; warnings: string[] } {
  const workers: WorkerBlock[] = [];
  const warnings: string[] = [];
  let currentWorker: WorkerBlock | null = null;

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return; // skip header

    const bVal = getCellString(row.getCell(IN_COL.NAME));
    const cVal = getCellNumber(row.getCell(IN_COL.RATE));
    const dVal = getCellString(row.getCell(IN_COL.ADDRESS));
    const eVal = getCellNumber(row.getCell(IN_COL.HOURS));
    const gVal = getCellNumber(row.getCell(IN_COL.MATERIALS));
    const hVal = getCellNumber(row.getCell(IN_COL.GAS));
    const iVal = getCellNumber(row.getCell(IN_COL.TICKET));
    const hasData = eVal > 0 || gVal > 0 || hVal > 0 || iVal > 0;

    // Check if this is a "Total" row
    if (bVal === "Total") {
      currentWorker = null;
      return;
    }

    // Check if this is a new worker row (has name in B and rate in C)
    if (bVal && cVal > 0) {
      currentWorker = { name: bVal, rate: cVal, entries: [] };
      workers.push(currentWorker);

      // This row might also have address data
      if (dVal) {
        currentWorker.entries.push({
          address: dVal,
          hours: eVal,
          materials: gVal,
          gas: hVal,
          ticket: iVal,
        });
      } else if (hasData) {
        warnings.push(`第${rowNumber}行: ${bVal} 有数据但缺少工地地址(D列)`);
      }
      return;
    }

    // Row has data but no address — flag as warning
    if (!dVal && currentWorker) {
      if (hasData) {
        warnings.push(`第${rowNumber}行: ${currentWorker.name} 有数据但缺少工地地址(D列)`);
      }
      return;
    }

    // Row has address but no worker context
    if (dVal && !currentWorker) {
      if (hasData) {
        warnings.push(`第${rowNumber}行: 工地${dVal}有数据但缺少工人信息(B/C列)`);
      }
      return;
    }

    // This is a continuation row with address data
    if (dVal && currentWorker) {
      currentWorker.entries.push({
        address: dVal,
        hours: eVal,
        materials: gVal,
        gas: hVal,
        ticket: iVal,
      });
    }
  });

  return { workers, warnings };
}

// --- Date-range sheet helpers ---

// Used only in findProjectBlocks to skip payment rows when collecting project names from col C.
// NOT used for row availability — payment rows are available for worker data.
const PAYMENT_LABELS = new Set(["第一笔款", "第二笔款", "第三笔款"]);

const MONTH_ABBREVS: Record<string, number> = {
  jan: 1, feb: 2, mar: 3, apr: 4, may: 5, jun: 6,
  jul: 7, aug: 8, sep: 9, oct: 10, nov: 11, dec: 12,
};

function parseDateRange(name: string): DateRange | null {
  const match = name.match(/^([A-Za-z]{3})\s+(\d{2})$/);
  if (!match) return null;
  const month = MONTH_ABBREVS[match[1].toLowerCase()];
  if (!month) return null;
  return {
    month,
    year: CENTURY_PREFIX + parseInt(match[2], 10),
  };
}

function getDateRangeSheets(
  workbook: ExcelJS.Workbook
): { sheet: ExcelJS.Worksheet; range: DateRange }[] {
  const result: { sheet: ExcelJS.Worksheet; range: DateRange }[] = [];
  for (const ws of workbook.worksheets) {
    const range = parseDateRange(ws.name);
    if (range) {
      result.push({ sheet: ws, range });
    }
  }
  // Sort most recent first (year desc, then month desc)
  result.sort((a, b) => {
    if (a.range.year !== b.range.year) return b.range.year - a.range.year;
    return b.range.month - a.range.month;
  });
  return result;
}

function findProjectBlocks(sheet: ExcelJS.Worksheet): ProjectBlock[] {
  const blocks: ProjectBlock[] = [];
  const headerRows: number[] = [];

  // Find all header rows (col A = "单号")
  sheet.eachRow((row, rowNum) => {
    if (getCellString(row.getCell(OUT_COL.LABEL)) === "单号") {
      headerRows.push(rowNum);
    }
  });

  for (let i = 0; i < headerRows.length; i++) {
    const headerRow = headerRows[i];
    const dataStartRow = headerRow + 1;

    // Find 总开销/总价格/利润率 below this header, bounded by next header or sheet end
    let totalExpenseRow = -1;
    let totalPriceRow = -1;
    let profitRow = -1;

    const searchEnd =
      i + 1 < headerRows.length ? headerRows[i + 1] - 1 : sheet.rowCount;
    for (let r = dataStartRow; r <= searchEnd; r++) {
      const aVal = getCellString(sheet.getRow(r).getCell(OUT_COL.LABEL));
      if (aVal === "总开销" && totalExpenseRow === -1) totalExpenseRow = r;
      if (aVal === "总价格" && totalPriceRow === -1) totalPriceRow = r;
      if (aVal === "利润率" && profitRow === -1) profitRow = r;
      if (totalExpenseRow > 0 && totalPriceRow > 0 && profitRow > 0) break;
    }

    if (totalExpenseRow === -1) continue; // skip malformed blocks

    // Collect project names from col C and col A values from data rows
    const projectNames: string[] = [];
    const colAValues: string[] = [];
    const seenC = new Set<string>();
    const seenA = new Set<string>();

    for (let r = dataStartRow; r < totalExpenseRow; r++) {
      const row = sheet.getRow(r);
      const aVal = getCellString(row.getCell(OUT_COL.LABEL));

      // Skip payment rows when collecting project identifiers
      if (PAYMENT_LABELS.has(aVal)) continue;

      const cVal = getCellString(row.getCell(OUT_COL.PROJECT));
      if (cVal && !seenC.has(cVal)) {
        seenC.add(cVal);
        projectNames.push(cVal);
      }
      if (aVal && !seenA.has(aVal)) {
        seenA.add(aVal);
        colAValues.push(aVal);
      }
    }

    blocks.push({
      dataStartRow,
      totalExpenseRow,
      totalPriceRow,
      profitRow,
      projectNames,
      colAValues,
    });
  }

  return blocks;
}

function escapeRegExp(str: string): string {
  return str.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function matchesAddressCode(text: string, addressCode: string): boolean {
  // Word-boundary match: address code must appear as a standalone token
  // e.g. "7217" matches "7217 Bridlewood" but not "17217 Elm"
  const escaped = escapeRegExp(addressCode);
  const regex = new RegExp(`(?:^|\\b)${escaped}(?:\\b|$)`);
  return regex.test(text);
}

function findMatchingBlock(
  blocks: ProjectBlock[],
  addressCode: string
): ProjectBlock | null {
  // Primary: check col C for word-boundary match
  for (const block of blocks) {
    for (const name of block.projectNames) {
      if (matchesAddressCode(name, addressCode)) {
        return block;
      }
    }
  }

  // Fallback: check col A for word-boundary match
  for (const block of blocks) {
    for (const aVal of block.colAValues) {
      if (matchesAddressCode(aVal, addressCode)) {
        return block;
      }
    }
  }

  return null;
}

// --- Write logic ---

function shiftFormulaReferences(formula: string, fromRow: number): string {
  return formula.replace(/(\$?)([A-Z]+)(\$?)(\d+)/g, (match, colAnchor, col, rowAnchor, rowStr) => {
    // Don't shift absolute row references (e.g. B$1, $B$1)
    if (rowAnchor === "$") return match;
    const row = parseInt(rowStr, 10);
    if (row >= fromRow) {
      return `${colAnchor}${col}${row + 1}`;
    }
    return match;
  });
}

/**
 * Convert all shared formulas in column H to explicit =Fn*Gn formulas.
 * Must run BEFORE any writes/insertions so that spliceRows can never
 * orphan shared-formula clones (which causes "Shared Formula master
 * must exist above and or left of clone" errors at writeBuffer time).
 */
function breakSharedFormulasInColH(sheet: ExcelJS.Worksheet): void {
  const rowCount = sheet.rowCount;
  for (let r = 1; r <= rowCount; r++) {
    const cell = sheet.getRow(r).getCell(OUT_COL.COST);
    const val = cell.value;
    if (!val || typeof val !== "object") continue;

    if ("sharedFormula" in val) {
      // Clone → explicit formula
      cell.value = { formula: `${COL.HOURS}${r}*${COL.RATE}${r}` } as ExcelJS.CellFormulaValue;
    } else if ("formula" in val && "shareType" in val) {
      // Master of shared group → plain formula (strip ref/shareType)
      const formula = (val as ExcelJS.CellFormulaValue).formula;
      cell.value = { formula } as ExcelJS.CellFormulaValue;
    }
  }
}

/**
 * Fix all formula references on the sheet after a row insertion at spliceAt.
 * ExcelJS's spliceRows shifts rows but does NOT update formula references.
 *
 * Column H shared formulas are already broken by breakSharedFormulasInColH,
 * so we only need to shift explicit formula references here.
 */
function fixFormulasAfterSplice(sheet: ExcelJS.Worksheet, spliceAt: number): void {
  sheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === spliceAt) return; // skip the newly inserted row

    row.eachCell({ includeEmpty: false }, (cell) => {
      const val = cell.value;
      if (!val || typeof val !== "object") return;

      if ("sharedFormula" in val) {
        // Shouldn't happen for col H (already broken), but handle other columns
        if (cell.col === COL.COST) {
          cell.value = {
            formula: `${COL.HOURS}${rowNumber}*${COL.RATE}${rowNumber}`,
          } as ExcelJS.CellFormulaValue;
        }
      } else if ("formula" in val) {
        const formula = (val as ExcelJS.CellFormulaValue).formula;
        const updated = shiftFormulaReferences(formula, spliceAt);
        if (updated !== formula) {
          cell.value = { formula: updated } as ExcelJS.CellFormulaValue;
        }
      }
    });
  });
}

/** A row is available for worker data if D (date) and E (worker name) are both empty.
 *  Rows with existing dates are project info rows — not available for new workers.
 *  Material data (K/L/M) is independent and does not affect worker placement. */
function isRowAvailableForWorker(sheet: ExcelJS.Worksheet, r: number): boolean {
  const row = sheet.getRow(r);
  const dVal = getCellString(row.getCell(OUT_COL.DATE));
  const eVal = getCellString(row.getCell(OUT_COL.WORKER));
  return !dVal && !eVal;
}

/**
 * Insert a new row into a block at spliceAt, fix all formulas on the sheet,
 * shift subsequent blocks, and update summary formulas.
 */
function spliceBlockRow(
  sheet: ExcelJS.Worksheet,
  block: ProjectBlock,
  allBlocks: ProjectBlock[],
  spliceAt: number
): void {
  sheet.spliceRows(spliceAt, 0, []);
  fixFormulasAfterSplice(sheet, spliceAt);

  // Summary rows shifted down by 1
  block.totalExpenseRow++;
  block.totalPriceRow++;
  block.profitRow++;

  // Shift all subsequent blocks on this sheet
  for (const other of allBlocks) {
    if (other === block) continue;
    if (other.dataStartRow >= spliceAt) other.dataStartRow++;
    if (other.totalExpenseRow >= spliceAt) other.totalExpenseRow++;
    if (other.totalPriceRow >= spliceAt) other.totalPriceRow++;
    if (other.profitRow >= spliceAt) other.profitRow++;
  }

  // Write explicit =F*G on newly inserted row
  sheet.getRow(spliceAt).getCell(OUT_COL.COST).value = {
    formula: `${COL.HOURS}${spliceAt}*${COL.RATE}${spliceAt}`,
  } as ExcelJS.CellFormulaValue;

  // Rewrite 总开销 SUM formulas to include the new row
  const dataEnd = block.totalExpenseRow - 1;
  const summaryRow = sheet.getRow(block.totalExpenseRow);
  summaryRow.getCell(OUT_COL.COST).value = {
    formula: `SUM(${COL.COST}${block.dataStartRow}:${COL.COST}${dataEnd})`,
  } as ExcelJS.CellFormulaValue;
  summaryRow.getCell(OUT_COL.OTHER).value = {
    formula: `SUM(${COL.OTHER}${block.dataStartRow}:${COL.OTHER}${dataEnd})`,
  } as ExcelJS.CellFormulaValue;
  summaryRow.getCell(OUT_COL.MATERIAL).value = {
    formula: `SUM(${COL.MATERIAL}${block.dataStartRow}:${COL.MATERIAL}${dataEnd})`,
  } as ExcelJS.CellFormulaValue;
  summaryRow.getCell(OUT_COL.AMOUNT).value = {
    formula: `SUM(${COL.COST}${block.totalExpenseRow}:${COL.MATERIAL}${block.totalExpenseRow})`,
  } as ExcelJS.CellFormulaValue;
}

/**
 * Find an empty row in the block (including payment and buffer rows),
 * or insert one before 总开销 as a last resort.
 * Returns the row number where data should be written.
 */
function findOrCreateInsertRow(
  sheet: ExcelJS.Worksheet,
  block: ProjectBlock,
  allBlocks: ProjectBlock[]
): number {
  // Search entire block for rows where D and E are empty
  for (let r = block.dataStartRow; r < block.totalExpenseRow; r++) {
    if (isRowAvailableForWorker(sheet, r)) {
      return r;
    }
  }

  // Truly no empty rows — insert before 总开销
  const insertAt = block.totalExpenseRow;
  spliceBlockRow(sheet, block, allBlocks, insertAt);
  return insertAt;
}

function isCellEmpty(cell: ExcelJS.Cell): boolean {
  const val = cell.value;
  return val == null || val === "" || val === 0;
}

/** Write material K/L/M values to a specific row. */
function setMaterialCells(
  row: ExcelJS.Row,
  materials: number,
  workerName: string,
  dateLabel: string
): void {
  row.getCell(OUT_COL.MATERIAL).value = materials;
  row.getCell(OUT_COL.NOTES).value = `${workerName}材料`;
  row.getCell(OUT_COL.MAT_DATE).value = dateLabel;
}

/** Write material cost to the first empty K cell within the block, inserting a row if needed. */
function writeMaterialToBlock(
  sheet: ExcelJS.Worksheet,
  block: ProjectBlock,
  allBlocks: ProjectBlock[],
  workerName: string,
  materials: number,
  dateLabel: string
): void {
  for (let r = block.dataStartRow; r < block.totalExpenseRow; r++) {
    if (isCellEmpty(sheet.getRow(r).getCell(OUT_COL.MATERIAL))) {
      setMaterialCells(sheet.getRow(r), materials, workerName, dateLabel);
      return;
    }
  }
  // No empty K cell — insert a new row before 总开销
  const insertAt = block.totalExpenseRow;
  spliceBlockRow(sheet, block, allBlocks, insertAt);
  setMaterialCells(sheet.getRow(insertAt), materials, workerName, dateLabel);
}

function writeEntryToBlock(
  sheet: ExcelJS.Worksheet,
  block: ProjectBlock,
  allBlocks: ProjectBlock[],
  worker: WorkerBlock,
  entry: WorkerEntry,
  dateLabel: string
): void {
  const insertRow = findOrCreateInsertRow(sheet, block, allBlocks);

  // Write the data row
  const row = sheet.getRow(insertRow);
  row.getCell(OUT_COL.DATE).value = dateLabel;
  row.getCell(OUT_COL.WORKER).value = worker.name;

  if (entry.hours > 0) {
    row.getCell(OUT_COL.HOURS).value = entry.hours;
    row.getCell(OUT_COL.RATE).value = worker.rate;
    // Only write explicit formula if the cell doesn't already have one
    const hVal = row.getCell(OUT_COL.COST).value;
    const hasFormula =
      hVal && typeof hVal === "object" && ("formula" in hVal || "sharedFormula" in hVal);
    if (!hasFormula) {
      row.getCell(OUT_COL.COST).value = {
        formula: `${COL.HOURS}${insertRow}*${COL.RATE}${insertRow}`,
      } as ExcelJS.CellFormulaValue;
    }
  }

  if (entry.gas + entry.ticket > 0) {
    row.getCell(OUT_COL.OTHER).value = entry.gas + entry.ticket;
  }

  if (entry.materials > 0) {
    // Search from top of block to pack materials contiguously in K column
    writeMaterialToBlock(sheet, block, allBlocks, worker.name, entry.materials, dateLabel);
  }

  row.commit();
}

// --- Cost summary computation ---

function computeCostSummaries(
  workers: WorkerBlock[],
  unmatchedAddresses: Set<string>
): { workerTotals: WorkerTotal[]; siteTotals: SiteTotal[] } {
  // Worker totals only include matched entries; omit workers with no matched data
  const workerTotals: WorkerTotal[] = workers.map((w) => {
    let labor = 0;
    let materials = 0;
    let gas = 0;
    let other = 0;
    for (const e of w.entries) {
      if (e.hours === 0 && e.materials === 0 && e.gas === 0 && e.ticket === 0) continue;
      if (unmatchedAddresses.has(e.address)) continue;
      labor += e.hours * w.rate;
      materials += e.materials;
      gas += e.gas;
      other += e.ticket;
    }
    return { name: w.name, rate: w.rate, labor, materials, gas, other, total: labor + materials + gas + other };
  }).filter((wt) => wt.total > 0);

  const siteMap = new Map<string, { labor: number; materials: number; gas: number; other: number }>();
  for (const w of workers) {
    for (const e of w.entries) {
      if (e.hours === 0 && e.materials === 0 && e.gas === 0 && e.ticket === 0) continue;
      if (unmatchedAddresses.has(e.address)) continue;
      const existing = siteMap.get(e.address) || { labor: 0, materials: 0, gas: 0, other: 0 };
      existing.labor += e.hours * w.rate;
      existing.materials += e.materials;
      existing.gas += e.gas;
      existing.other += e.ticket;
      siteMap.set(e.address, existing);
    }
  }
  const siteTotals: SiteTotal[] = Array.from(siteMap.entries()).map(([address, s]) => ({
    address,
    labor: s.labor,
    materials: s.materials,
    gas: s.gas,
    other: s.other,
    total: s.labor + s.materials + s.gas + s.other,
  }));

  return { workerTotals, siteTotals };
}

// --- Worker entry processing ---

function processWorkerEntries(
  workers: WorkerBlock[],
  dateSheets: { sheet: ExcelJS.Worksheet; range: DateRange }[],
  sheetBlocks: Map<ExcelJS.Worksheet, ProjectBlock[]>,
  dateLabel: string
): { matchedSheets: Set<string>; unmatchedAddresses: Set<string>; rowsAdded: number } {
  const matchedSheets = new Set<string>();
  const unmatchedAddresses = new Set<string>();
  let rowsAdded = 0;

  for (const worker of workers) {
    for (const entry of worker.entries) {
      if (entry.hours === 0 && entry.materials === 0 && entry.gas === 0 && entry.ticket === 0) {
        continue;
      }

      let found = false;

      for (const { sheet } of dateSheets) {
        const blocks = sheetBlocks.get(sheet);
        if (!blocks) continue;
        const block = findMatchingBlock(blocks, entry.address);
        if (block) {
          matchedSheets.add(sheet.name);

          const isMaterialsOnly =
            entry.hours === 0 && entry.gas === 0 && entry.ticket === 0 && entry.materials > 0;

          if (isMaterialsOnly) {
            writeMaterialToBlock(sheet, block, blocks, worker.name, entry.materials, dateLabel);
          } else {
            writeEntryToBlock(sheet, block, blocks, worker, entry, dateLabel);
            rowsAdded++;
          }

          found = true;
          break;
        }
      }

      if (!found) {
        unmatchedAddresses.add(entry.address);
      }
    }
  }

  return { matchedSheets, unmatchedAddresses, rowsAdded };
}

// --- Main processing ---

export async function processFiles(
  inputBuffer: ArrayBuffer,
  outputBuffer: ArrayBuffer,
  dateLabel: string
): Promise<{ result: ProcessResult; outputFile: Buffer }> {
  if (inputBuffer.byteLength > MAX_FILE_SIZE) {
    throw new Error(
      `Input file exceeds ${MAX_FILE_SIZE / 1024 / 1024} MB limit`
    );
  }
  if (outputBuffer.byteLength > MAX_FILE_SIZE) {
    throw new Error(
      `Output file exceeds ${MAX_FILE_SIZE / 1024 / 1024} MB limit`
    );
  }

  const inputWb = await loadWorkbook(inputBuffer);
  const inputWs = inputWb.worksheets[0];

  if (!inputWs) {
    throw new Error("Input file has no worksheets");
  }

  const { workers, warnings } = parseInputFile(inputWs);

  const outputWb = await loadWorkbook(outputBuffer);

  // Get date-range sheets sorted most recent first
  const dateSheets = getDateRangeSheets(outputWb);

  // Break shared formulas in col H before any writes to prevent
  // "Shared Formula master must exist above and or left of clone" errors
  for (const { sheet } of dateSheets) {
    breakSharedFormulasInColH(sheet);
  }

  // Pre-compute blocks for each sheet
  const sheetBlocks = new Map<
    ExcelJS.Worksheet,
    ProjectBlock[]
  >();
  for (const { sheet } of dateSheets) {
    sheetBlocks.set(sheet, findProjectBlocks(sheet));
  }

  const { matchedSheets, unmatchedAddresses, rowsAdded } =
    processWorkerEntries(workers, dateSheets, sheetBlocks, dateLabel);

  const { workerTotals, siteTotals } = computeCostSummaries(workers, unmatchedAddresses);

  const outputFileBuffer = await outputWb.xlsx.writeBuffer();

  return {
    result: {
      workers,
      matchedSheets: Array.from(matchedSheets),
      unmatchedAddresses: Array.from(unmatchedAddresses),
      rowsAdded,
      warnings,
      workerTotals,
      siteTotals,
    },
    outputFile: Buffer.from(outputFileBuffer),
  };
}
