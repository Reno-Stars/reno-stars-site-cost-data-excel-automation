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

export interface ProcessResult {
  workers: WorkerBlock[];
  matchedSheets: string[];
  unmatchedAddresses: string[];
  rowsAdded: number;
}

export const MAX_FILE_SIZE = 50 * 1024 * 1024; // 50 MB

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

export function parseInputFile(worksheet: ExcelJS.Worksheet): WorkerBlock[] {
  const workers: WorkerBlock[] = [];
  let currentWorker: WorkerBlock | null = null;

  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber === 1) return; // skip header

    const bVal = getCellString(row.getCell(2)); // B = Name
    const cVal = getCellNumber(row.getCell(3)); // C = $/hr
    const dVal = getCellString(row.getCell(4)); // D = Address
    const eVal = getCellNumber(row.getCell(5)); // E = Hours
    const gVal = getCellNumber(row.getCell(7)); // G = Material$
    const hVal = getCellNumber(row.getCell(8)); // H = Gas$
    const iVal = getCellNumber(row.getCell(9)); // I = Ticket

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

  return workers;
}

function extractAddressCode(sheetName: string): string {
  const match = sheetName.match(/^(\d+)/);
  return match ? match[1] : "";
}

function findSummaryRows(sheet: ExcelJS.Worksheet): {
  totalExpenseRow: number;
  totalPriceRow: number;
  profitRow: number;
} {
  let totalExpenseRow = -1;
  let totalPriceRow = -1;
  let profitRow = -1;

  sheet.eachRow((row, rowNum) => {
    const aVal = getCellString(row.getCell(1));
    if (aVal === "总开销" && totalExpenseRow === -1) totalExpenseRow = rowNum;
    if (aVal === "总价格" && totalPriceRow === -1) totalPriceRow = rowNum;
    if (aVal === "利润率" && profitRow === -1) profitRow = rowNum;
  });

  return { totalExpenseRow, totalPriceRow, profitRow };
}

function shiftFormulaReferences(formula: string, fromRow: number): string {
  return formula.replace(/([A-Z]+)(\d+)/g, (match, col, rowStr) => {
    const row = parseInt(rowStr, 10);
    if (row >= fromRow) {
      return `${col}${row + 1}`;
    }
    return match;
  });
}

function writeEntryToSheet(
  sheet: ExcelJS.Worksheet,
  worker: WorkerBlock,
  entry: WorkerEntry,
  dateLabel: string
): void {
  let { totalExpenseRow, totalPriceRow, profitRow } = findSummaryRows(sheet);

  // Search from row 2 for the first empty data row (D/E/F/G all empty)
  const searchEnd = totalExpenseRow > 0 ? totalExpenseRow - 1 : sheet.rowCount;
  let insertRow = -1;

  for (let r = 2; r <= searchEnd; r++) {
    const row = sheet.getRow(r);
    const dVal = getCellString(row.getCell(4));
    const eVal = getCellString(row.getCell(5));
    const fVal = getCellNumber(row.getCell(6));
    const gVal = getCellNumber(row.getCell(7));

    if (!dVal && !eVal && fVal === 0 && gVal === 0) {
      insertRow = r;
      break;
    }
  }

  if (insertRow === -1) {
    // No empty row found — insert before summary row
    if (totalExpenseRow > 0) {
      const spliceAt = totalExpenseRow;
      sheet.spliceRows(spliceAt, 0, []);
      insertRow = spliceAt;

      // All summary rows shifted down by 1
      totalExpenseRow++;
      if (totalPriceRow > 0) totalPriceRow++;
      if (profitRow > 0) profitRow++;

      // Update 总开销 formulas with expanded ranges
      const summaryRow = sheet.getRow(totalExpenseRow);
      const dataEnd = totalExpenseRow - 1;
      summaryRow.getCell(8).value = {
        formula: `SUM(H2:H${dataEnd})`,
      } as ExcelJS.CellFormulaValue;
      summaryRow.getCell(9).value = {
        formula: `SUM(I2:I${dataEnd})`,
      } as ExcelJS.CellFormulaValue;
      summaryRow.getCell(11).value = {
        formula: `SUM(K2:K${dataEnd})`,
      } as ExcelJS.CellFormulaValue;
      summaryRow.getCell(2).value = {
        formula: `SUM(H${totalExpenseRow}:K${totalExpenseRow})`,
      } as ExcelJS.CellFormulaValue;

      // Update formula references in 总价格 and 利润率 rows
      for (const rowNum of [totalPriceRow, profitRow]) {
        if (rowNum <= 0) continue;
        const sRow = sheet.getRow(rowNum);
        sRow.eachCell({ includeEmpty: false }, (cell) => {
          const val = cell.value;
          if (val && typeof val === "object" && "formula" in val) {
            cell.value = {
              formula: shiftFormulaReferences(
                (val as ExcelJS.CellFormulaValue).formula,
                spliceAt
              ),
            } as ExcelJS.CellFormulaValue;
          }
        });
      }
    } else {
      insertRow = sheet.rowCount + 1;
    }
  }

  // Write the data row — always include date and worker name for context
  const row = sheet.getRow(insertRow);
  row.getCell(4).value = dateLabel; // D = 日期
  row.getCell(5).value = worker.name; // E = 工人

  if (entry.hours > 0) {
    row.getCell(6).value = entry.hours; // F = 工时
    row.getCell(7).value = worker.rate; // G = hourly rate
    row.getCell(8).value = {
      formula: `F${insertRow}*G${insertRow}`,
    } as ExcelJS.CellFormulaValue;
  }

  const otherExpenses = entry.gas + entry.ticket;
  if (otherExpenses > 0) {
    row.getCell(9).value = otherExpenses; // I = other expenses
  }

  // Add materials — same row if K is empty, otherwise find next empty K before summary
  if (entry.materials > 0) {
    if (getCellNumber(row.getCell(11)) === 0) {
      row.getCell(11).value = entry.materials; // K = material cost
      row.getCell(12).value = `${worker.name}材料`; // L = note
      row.getCell(13).value = dateLabel; // M = date
    } else {
      const materialSearchEnd =
        totalExpenseRow > 0 ? totalExpenseRow - 1 : sheet.rowCount;
      for (let r = insertRow + 1; r <= materialSearchEnd; r++) {
        const mRow = sheet.getRow(r);
        if (getCellNumber(mRow.getCell(11)) === 0) {
          mRow.getCell(11).value = entry.materials;
          mRow.getCell(12).value = `${worker.name}材料`;
          mRow.getCell(13).value = dateLabel;
          break;
        }
      }
    }
  }

  row.commit();
}

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

  const workers = parseInputFile(inputWs);

  const outputWb = await loadWorkbook(outputBuffer);

  // Build address-to-sheet mapping
  const sheetMap = new Map<string, ExcelJS.Worksheet>();
  for (const ws of outputWb.worksheets) {
    const code = extractAddressCode(ws.name);
    if (code) {
      sheetMap.set(code, ws);
    }
  }

  const matchedSheets = new Set<string>();
  const unmatchedAddresses = new Set<string>();
  let rowsAdded = 0;

  for (const worker of workers) {
    for (const entry of worker.entries) {
      // Skip entries with no data at all
      if (
        entry.hours === 0 &&
        entry.materials === 0 &&
        entry.gas === 0 &&
        entry.ticket === 0
      ) {
        continue;
      }

      const sheet = sheetMap.get(entry.address);
      if (!sheet) {
        unmatchedAddresses.add(entry.address);
        continue;
      }

      matchedSheets.add(sheet.name);
      writeEntryToSheet(sheet, worker, entry, dateLabel);
      rowsAdded++;
    }
  }

  const outputFileBuffer = await outputWb.xlsx.writeBuffer();

  return {
    result: {
      workers,
      matchedSheets: Array.from(matchedSheets),
      unmatchedAddresses: Array.from(unmatchedAddresses),
      rowsAdded,
    },
    outputFile: Buffer.from(outputFileBuffer),
  };
}
