/**
 * Generates comprehensive mock input and output xlsx files for manual testing.
 *
 * Run with: npx tsx scripts/generate-test-files.ts
 *
 * Edge cases covered:
 *
 * INPUT:
 *  1. Worker with multiple addresses (hours + materials + gas + ticket)
 *  2. Worker with hours only
 *  3. Worker with materials only (no hours)
 *  4. Worker with gas+ticket only
 *  5. Entry with all zeros (should be skipped)
 *  6. Address that matches in col C of most recent sheet
 *  7. Address that matches in col A (fallback)
 *  8. Address that only exists in an older sheet (tests sheet ordering)
 *  9. Address that matches nothing (unmatched)
 * 10. Short address code that could false-positive (e.g. "53" should NOT match "1530")
 * 11. Multiple entries for the same block (tests sequential empty-row search)
 *
 * OUTPUT:
 *  1. Multiple date-range sheets sorted by recency (Jan 26, Sep 25, Apr 25)
 *  2. Non-date sheets that should be skipped (Balance, T4a)
 *  3. Block with existing data (partially filled)
 *  4. Block with data in col A only (tests fallback matching)
 *  5. Nearly full block (only 1 empty row — second entry triggers insertion)
 *  6. Payment rows with existing material data in K/L/M
 *  7. Shared formulas in col H (=F*G)
 *  8. Block that only exists in older sheet (Sep 25)
 *  9. Empty template block (客户名+项目 placeholder in col C)
 * 10. Block with address "1530" to test false-positive against input "53"
 */

import ExcelJS from "exceljs";
import { writeFileSync } from "fs";
import { join } from "path";

const ROOT = join(__dirname, "..");

// ─── Helper to build a project block ────────────────────────────

interface BlockDef {
  headerRow: number;
  // Data rows: array of { A?, C?, D?, E?, F?, G?, I?, K?, L?, M? }
  dataRows: Record<string, string | number | undefined>[];
  // Payment amounts [first, second, third]
  payments: [number, number, number];
  // Extra material on payment rows: [{row: 0|1|2, K, L, M}]
  paymentMaterials?: { row: number; K: number; L: string; M: string }[];
  // Total buffer rows (empty data rows between payments and 总开销)
  bufferRows: number;
}

function writeBlock(
  sheet: ExcelJS.Worksheet,
  def: BlockDef
): void {
  const h = def.headerRow;

  // Header row
  const header = sheet.getRow(h);
  header.getCell(1).value = "单号";
  header.getCell(2).value = "负责人";
  header.getCell(3).value = "客户名+项目";
  header.getCell(4).value = "日期";
  header.getCell(5).value = "工人";
  header.getCell(6).value = "工时";
  header.getCell(7).value = "hourly rate";
  header.getCell(8).value = "开销";
  header.getCell(9).value = "其他开销";
  header.getCell(10).value = "Note";
  header.getCell(11).value = "材料开销";

  let r = h + 1;

  // Data rows
  for (const data of def.dataRows) {
    const row = sheet.getRow(r);
    if (data.A !== undefined) row.getCell(1).value = data.A;
    if (data.C !== undefined) row.getCell(3).value = data.C;
    if (data.D !== undefined) row.getCell(4).value = data.D;
    if (data.E !== undefined) row.getCell(5).value = data.E;
    if (data.F !== undefined) row.getCell(6).value = data.F as number;
    if (data.G !== undefined) row.getCell(7).value = data.G as number;
    // Col H = F*G formula (shared formula pattern)
    row.getCell(8).value = { formula: `F${r}*G${r}` } as ExcelJS.CellFormulaValue;
    if (data.I !== undefined) row.getCell(9).value = data.I as number;
    if (data.K !== undefined) row.getCell(11).value = data.K as number;
    if (data.L !== undefined) row.getCell(12).value = data.L;
    if (data.M !== undefined) row.getCell(13).value = data.M;
    r++;
  }

  // Payment rows
  const paymentLabels = ["第一笔款", "第二笔款", "第三笔款"];
  const paymentStartRow = r;
  for (let i = 0; i < 3; i++) {
    const row = sheet.getRow(r);
    row.getCell(1).value = paymentLabels[i];
    if (def.payments[i] > 0) row.getCell(2).value = def.payments[i];
    row.getCell(8).value = { formula: `F${r}*G${r}` } as ExcelJS.CellFormulaValue;
    // Payment material data
    if (def.paymentMaterials) {
      const pm = def.paymentMaterials.find((p) => p.row === i);
      if (pm) {
        row.getCell(11).value = pm.K;
        row.getCell(12).value = pm.L;
        row.getCell(13).value = pm.M;
      }
    }
    r++;
  }

  // Buffer rows (empty with =F*G)
  for (let i = 0; i < def.bufferRows; i++) {
    const row = sheet.getRow(r);
    row.getCell(8).value = { formula: `F${r}*G${r}` } as ExcelJS.CellFormulaValue;
    r++;
  }

  const dataStartRow = h + 1;
  const dataEndRow = r - 1;

  // 总开销
  const expRow = sheet.getRow(r);
  expRow.getCell(1).value = "总开销";
  expRow.getCell(2).value = {
    formula: `SUM(H${r}:K${r})`,
  } as ExcelJS.CellFormulaValue;
  expRow.getCell(8).value = {
    formula: `SUM(H${dataStartRow}:H${dataEndRow})`,
  } as ExcelJS.CellFormulaValue;
  expRow.getCell(9).value = {
    formula: `SUM(I${dataStartRow}:I${dataEndRow})`,
  } as ExcelJS.CellFormulaValue;
  expRow.getCell(11).value = {
    formula: `SUM(K${dataStartRow}:K${dataEndRow})`,
  } as ExcelJS.CellFormulaValue;
  r++;

  // 总价格
  const priceRow = sheet.getRow(r);
  priceRow.getCell(1).value = "总价格";
  priceRow.getCell(2).value = {
    formula: `SUM(B${paymentStartRow}:B${paymentStartRow + 2})`,
  } as ExcelJS.CellFormulaValue;
  r++;

  // 利润率
  const profitRow = sheet.getRow(r);
  profitRow.getCell(1).value = "利润率";
  profitRow.getCell(2).value = {
    formula: `(B${r - 1}-B${r - 2})/B${r - 1}`,
  } as ExcelJS.CellFormulaValue;
}

// ─── Generate output template ───────────────────────────────────

async function generateOutputFile(): Promise<void> {
  const wb = new ExcelJS.Workbook();

  // ── Sheet 1: "Jan 26" (most recent) ──
  const s1 = wb.addWorksheet("Jan 26");

  // Block 1: "7217 Bridlewood" — partially filled (3 data rows used, tests normal match on col C)
  writeBlock(s1, {
    headerRow: 1,
    dataRows: [
      { C: "Chi Man", D: "1月上", E: "阿华", F: 49.5, G: 14, K: 6300, L: "电工定金" },
      { C: "7217 Bridlewood", D: "1月上", E: "阿华", F: 50, G: 14, K: 1000, L: "钢梁定金", M: "2026-01-17" },
      { C: "Burnaby", K: 8400, L: "Andy水工80%", M: "1月中" },
      // Row 5 is empty — new data should go here
    ],
    payments: [4000, 0, 0],
    paymentMaterials: [{ row: 0, K: 500, L: "定金材料", M: "1月" }],
    bufferRows: 20,
  });
  // Block 1 occupies rows 1-33

  // Blank separator rows
  // Block 2 starts at row 36: "1530 Oak Street" — tests false-positive protection
  // Address "53" in input should NOT match this block
  writeBlock(s1, {
    headerRow: 36,
    dataRows: [
      { A: "unit 5", C: "1530 Oak Street", D: "2月上", E: "老赵", F: 30, G: 15 },
      { C: "Vancouver" },
    ],
    payments: [2000, 0, 0],
    bufferRows: 8,
  });
  // Block 2: rows 36-52

  // Block 3 starts at row 55: "5880 Dover" — nearly full (only 1 empty row left)
  // Tests: two entries to same block → second triggers block-full insertion
  writeBlock(s1, {
    headerRow: 55,
    dataRows: [
      { C: "5880 Dover", D: "1月上", E: "老赵", F: 69.5, G: 13, K: 92.76, L: "补砖" },
      { C: "Richmond", K: 78.62, L: "老赵材料", M: "1月上" },
      { C: "5880 Dover", D: "1月中", E: "阿华", F: 20, G: 14 },
      { C: "5880 Dover", D: "2月上", E: "老赵", F: 15, G: 13, I: 50 },
      // Only 1 empty data row remains (row 60) before payments
    ],
    payments: [1500, 800, 0],
    paymentMaterials: [
      { row: 0, K: 1260, L: "Aaron 水工", M: "2月19" },
      { row: 1, K: 700, L: "Chris 台面", M: "3月1" },
    ],
    bufferRows: 0, // NO buffer rows — nearly full!
  });
  // Block 3: rows 55-67

  // Block 4 starts at row 70: col A fallback match
  // Col C has generic name, but col A has address "302"
  writeBlock(s1, {
    headerRow: 70,
    dataRows: [
      { A: "302", C: "Marco Project", D: "1月上", E: "阿华", F: 40, G: 14 },
    ],
    payments: [0, 0, 0],
    bufferRows: 15,
  });
  // Block 4: rows 70-92

  // Block 5 starts at row 95: "53 Maple Ave" — tests that "53" matches here, not "1530"
  writeBlock(s1, {
    headerRow: 95,
    dataRows: [
      { C: "53 Maple Ave", D: "2月上", E: "老赵", F: 10, G: 13 },
    ],
    payments: [0, 0, 0],
    bufferRows: 15,
  });
  // Block 5: rows 95-117

  // Block 6 starts at row 120: "4444 Test Lane" — ALL K cells filled (tests material row insertion)
  writeBlock(s1, {
    headerRow: 120,
    dataRows: [
      { C: "4444 Test Lane", D: "1月上", E: "阿华", F: 10, G: 14, K: 100, L: "材料1" },
      { C: "4444 Test Lane", D: "1月中", E: "老赵", F: 5, G: 13, K: 200, L: "材料2" },
    ],
    payments: [1000, 0, 0],
    paymentMaterials: [
      { row: 0, K: 300, L: "定金材料", M: "1月" },
      { row: 1, K: 400, L: "材料3", M: "2月" },
      { row: 2, K: 500, L: "材料4", M: "3月" },
    ],
    bufferRows: 0,
  });
  // Block 6: rows 120-128 (all K cells occupied, 0 buffer)

  // ── Sheet 2: "Sep 25" (older) ──
  const s2 = wb.addWorksheet("Sep 25");

  // Block 1: "9000 Ash Grove" — only in this older sheet (tests sheet ordering)
  writeBlock(s2, {
    headerRow: 1,
    dataRows: [
      { C: "9000 Ash Grove Crescent", D: "10月上", E: "阿华", F: 30, G: 14 },
    ],
    payments: [3000, 0, 0],
    bufferRows: 20,
  });

  // Block 2: "7217 Bridlewood" also here — but should NOT be matched (newer sheet takes priority)
  writeBlock(s2, {
    headerRow: 36,
    dataRows: [
      { C: "7217 Bridlewood", D: "9月上", E: "老赵", F: 20, G: 13 },
    ],
    payments: [1000, 0, 0],
    bufferRows: 10,
  });

  // ── Sheet 3: "Apr 25" (oldest) ──
  const s3 = wb.addWorksheet("Apr 25");

  writeBlock(s3, {
    headerRow: 1,
    dataRows: [
      { C: "6033 Williams Rd", D: "4月上", E: "阿华", F: 25, G: 14 },
    ],
    payments: [0, 0, 0],
    bufferRows: 20,
  });

  // ── Sheet 4: "Balance" (non-date, should be skipped) ──
  const s4 = wb.addWorksheet("Balance");
  s4.getRow(1).getCell(1).value = "Income";
  s4.getRow(1).getCell(6).value = "total";
  s4.getRow(2).getCell(1).value = "项目";
  s4.getRow(2).getCell(2).value = "日期";
  s4.getRow(2).getCell(3).value = "金额";

  // ── Sheet 5: "T4a" (non-date, should be skipped) ──
  const s5 = wb.addWorksheet("T4a");
  s5.getRow(1).getCell(1).value = "所有人的";
  s5.getRow(2).getCell(1).value = "名字";

  const buffer = await wb.xlsx.writeBuffer();
  writeFileSync(join(ROOT, "test-output-template.xlsx"), Buffer.from(buffer));
  console.log("✓ Written test-output-template.xlsx");
}

// ─── Generate input file ────────────────────────────────────────

async function generateInputFile(): Promise<void> {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet("Sheet1");

  // Header
  const hdr = ws.getRow(1);
  hdr.getCell(2).value = "Name";
  hdr.getCell(3).value = "$/hr";
  hdr.getCell(4).value = "Address";
  hdr.getCell(5).value = "Hours";
  hdr.getCell(7).value = "Material$";
  hdr.getCell(8).value = "Gas$";
  hdr.getCell(9).value = "Ticket";

  let r = 2;

  function addRow(vals: {
    B?: string; C?: number; D?: string; E?: number; G?: number; H?: number; I?: number;
  }) {
    const row = ws.getRow(r);
    if (vals.B) row.getCell(2).value = vals.B;
    if (vals.C) row.getCell(3).value = vals.C;
    if (vals.D) row.getCell(4).value = vals.D;
    if (vals.E) row.getCell(5).value = vals.E;
    if (vals.G) row.getCell(7).value = vals.G;
    if (vals.H) row.getCell(8).value = vals.H;
    if (vals.I) row.getCell(9).value = vals.I;
    r++;
  }

  // ── Worker 1: 阿华 ($14/hr) ──
  // Case 1: Normal match — hours + materials + gas + ticket → "7217" matches "7217 Bridlewood"
  // Case 2: Materials only → "7217" again (tests material on separate row)
  // Case 5: All zeros → should be skipped
  // Case 8: Match in older sheet → "9000" matches "9000 Ash Grove Crescent" in "Sep 25"
  // Case 9: No match → "9999" is unmatched
  addRow({ B: "阿华", C: 14, D: "7217", E: 8, G: 200, H: 30, I: 15 });
  addRow({ D: "7217", G: 150 }); // materials only for same address
  addRow({ D: "5880", E: 6 }); // hours only → "5880 Dover"
  addRow({ E: 3 }); // WARNING: has hours but no address (D empty)
  addRow({ D: "9999" }); // all zeros — should be skipped
  addRow({ D: "9000", E: 10, G: 300 }); // match in older sheet
  addRow({ D: "8888", E: 5 }); // unmatched address
  addRow({ B: "Total" });

  // ── Worker 2: 老赵 ($13/hr) ──
  // Case 4: Gas+ticket only → "7217" matches
  // Case 10: Short code "53" should match "53 Maple Ave" NOT "1530 Oak Street"
  // Case 7: Col A fallback → "302" matches block 4 via col A
  // Case 11: Second entry to nearly-full block 5880 → triggers block insertion
  addRow({ B: "老赵", C: 13, D: "7217", H: 25, I: 10 }); // gas+ticket only
  addRow({ D: "53", E: 12, G: 500 }); // short code — must match "53 Maple Ave"
  addRow({ D: "302", E: 20 }); // col A fallback match
  addRow({ D: "5880", E: 15, G: 800, H: 40 }); // second entry to nearly-full block → insertion
  addRow({ B: "Total" });

  // ── Worker 3: Chris ($18/hr) ──
  // Case: match in oldest sheet → "6033" matches "6033 Williams Rd" in "Apr 25"
  // Case: materials-only to block with ALL K cells filled → triggers row insertion
  addRow({ B: "Chris", C: 18, D: "6033", E: 24, G: 1200, H: 50, I: 20 });
  addRow({ D: "4444", G: 600 }); // materials-only → Block 6, all K full → inserts row
  addRow({ B: "Total" });

  const buffer = await wb.xlsx.writeBuffer();
  writeFileSync(join(ROOT, "test-input.xlsx"), Buffer.from(buffer));
  console.log("✓ Written test-input.xlsx");
}

// ─── Expected results documentation ─────────────────────────────

function printExpectedResults(): void {
  console.log(`
╔══════════════════════════════════════════════════════════════════╗
║                    EXPECTED TEST RESULTS                        ║
╠══════════════════════════════════════════════════════════════════╣
║                                                                  ║
║  Workers parsed: 3 (阿华, 老赵, Chris)                          ║
║  Rows added: 8  (materials-only entries don't create rows)      ║
║  Matched sheets: "Jan 26", "Sep 25", "Apr 25"                 ║
║  Unmatched addresses: "8888"                                     ║
║  Warnings: 1 (阿华 has hours but no address at row 5)           ║
║                                                                  ║
║  Skipped entries (all zeros): "9999" (阿华)                      ║
║                                                                  ║
║  NOTE: Col C is NEVER written on new rows.                       ║
║  Materials-only entries write K/L/M to first empty K cell;       ║
║  if no empty K cell exists, inserts a new row before 总开销.     ║
║                                                                  ║
╠══════════════════════════════════════════════════════════════════╣
║  ENTRY-BY-ENTRY BREAKDOWN                                        ║
╠══════════════════════════════════════════════════════════════════╣
║                                                                  ║
║  阿华 @ 7217 (hrs=8, mat=200, gas=30, tkt=15)                   ║
║    → Sheet "Jan 26", Block 1 ("7217 Bridlewood")                ║
║    → Inserted after last worker row 3 → row 4:                   ║
║      D=dateLabel, E=阿华, F=8, G=14, H=F*G, I=45,              ║
║      K=200, L=阿华材料, M=dateLabel                              ║
║                                                                  ║
║  阿华 @ 7217 (mat=150 only)                                     ║
║    → Sheet "Jan 26", Block 1 — materials-only, no new row       ║
║    → Writes K=150, L=阿华材料, M=dateLabel to first empty K     ║
║      cell (第二笔款 row)                                         ║
║                                                                  ║
║  阿华 @ 5880 (hrs=6)                                            ║
║    → Sheet "Jan 26", Block 3 ("5880 Dover")                     ║
║    → Inserted after last worker: D=dateLabel, E=阿华, F=6, G=14 ║
║                                                                  ║
║  阿华 @ 9999 — ALL ZEROS → SKIPPED                              ║
║                                                                  ║
║  阿华 @ 9000 (hrs=10, mat=300)                                  ║
║    → Sheet "Sep 25", Block 1 ("9000 Ash Grove Crescent")       ║
║    → First empty row in that block                               ║
║                                                                  ║
║  阿华 @ 8888 (hrs=5) → UNMATCHED                                ║
║                                                                  ║
║  老赵 @ 7217 (gas=25, tkt=10)                                   ║
║    → Sheet "Jan 26", Block 1 ("7217 Bridlewood")                ║
║    → Inserted after last worker: D=dateLabel, E=老赵, I=35      ║
║                                                                  ║
║  老赵 @ 53 (hrs=12, mat=500)                                    ║
║    → Sheet "Jan 26", Block 5 ("53 Maple Ave")  ← NOT Block 2!  ║
║    → First empty row: D=dateLabel, E=老赵, F=12, G=13,          ║
║      K=500, L=老赵材料                                           ║
║                                                                  ║
║  老赵 @ 302 (hrs=20)                                            ║
║    → Sheet "Jan 26", Block 4 (col A fallback → "Marco Project") ║
║    → First empty row: D=dateLabel, E=老赵, F=20, G=13           ║
║                                                                  ║
║  老赵 @ 5880 (hrs=15, mat=800, gas=40)                          ║
║    → Sheet "Jan 26", Block 3 ("5880 Dover")                     ║
║    → Block full! Inserts after last worker row                   ║
║    → D=dateLabel, E=老赵, F=15, G=13, I=40,                    ║
║      K=800, L=老赵材料                                           ║
║    → 总开销 SUM ranges expanded, 总价格/利润率 refs shifted      ║
║                                                                  ║
║  Chris @ 6033 (hrs=24, mat=1200, gas=50, tkt=20)                ║
║    → Sheet "Apr 25", Block 1 ("6033 Williams Rd")               ║
║    → First empty row: D=dateLabel, E=Chris, F=24, G=18,         ║
║      H=F*G, I=70, K=1200, L=Chris材料                           ║
║                                                                  ║
║  Chris @ 4444 (mat=600 only)                                     ║
║    → Sheet "Jan 26", Block 6 ("4444 Test Lane")                 ║
║    → ALL K cells occupied → inserts row before 总开销            ║
║    → New row: K=600, L=Chris材料, M=dateLabel                    ║
║    → 总开销 shifted from row 126 to row 127                      ║
║                                                                  ║
╠══════════════════════════════════════════════════════════════════╣
║  KEY EDGE CASES TO VERIFY IN EXCEL                               ║
╠══════════════════════════════════════════════════════════════════╣
║                                                                  ║
║  1. "53" matched Block 5 (53 Maple Ave), NOT Block 2 (1530 Oak) ║
║  2. "302" matched via col A fallback → Block 4                   ║
║  3. "9000" matched in older sheet "Sep 25" (not in "Jan 26")   ║
║  4. "7217" matched in "Jan 26" (not "Sep 25" where it exists)  ║
║  5. Block 3 (5880) insertion: new row after last worker, ok      ║
║  6. Balance and T4a sheets untouched                             ║
║  7. "9999" (all zeros) was skipped entirely                      ║
║  8. "8888" appears in unmatchedAddresses                         ║
║  9. Payment rows (第一笔款 etc.) were not overwritten             ║
║ 10. Shared formulas in col H preserved on existing rows          ║
║ 11. Col C is never written on new data rows                      ║
║ 12. Materials-only entries write to empty K cell, no new row     ║
║ 13. Materials-only to full-K block inserts row before 总开销    ║
║ 14. Missing address in input generates warning (not silent skip)║
║ 15. Worker totals exclude unmatched entries, split gas/ticket   ║
║                                                                  ║
╚══════════════════════════════════════════════════════════════════╝
`);
}

async function main() {
  await generateOutputFile();
  await generateInputFile();
  printExpectedResults();
}

main().catch(console.error);
