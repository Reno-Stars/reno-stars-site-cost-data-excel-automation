/**
 * Smoke test: runs processFiles against the generated test files
 * and prints the result summary + inspects key cells in the output.
 *
 * Run with: npx tsx scripts/smoke-test.ts
 */

import { readFileSync, writeFileSync } from "fs";
import { join } from "path";
import ExcelJS from "exceljs";

const ROOT = join(__dirname, "..");

async function main() {
  const { processFiles } = await import("../lib/excel-processor");

  const inputBuf = readFileSync(join(ROOT, "test-input.xlsx"));
  const outputBuf = readFileSync(join(ROOT, "test-output-template.xlsx"));

  const { result, outputFile } = await processFiles(
    inputBuf.buffer.slice(inputBuf.byteOffset, inputBuf.byteOffset + inputBuf.byteLength),
    outputBuf.buffer.slice(outputBuf.byteOffset, outputBuf.byteOffset + outputBuf.byteLength),
    "3月下"
  );

  console.log("=== PROCESS RESULT ===");
  console.log(`Workers parsed: ${result.workers.length}`);
  console.log(
    `  ${result.workers.map((w) => `${w.name} ($${w.rate}/hr, ${w.entries.length} entries)`).join(", ")}`
  );
  console.log(`Rows added: ${result.rowsAdded}`);
  console.log(`Matched sheets: ${result.matchedSheets.join(", ")}`);
  console.log(`Unmatched addresses: ${result.unmatchedAddresses.join(", ") || "(none)"}`);

  writeFileSync(join(ROOT, "test-output-result.xlsx"), outputFile);
  console.log("\n✓ Saved test-output-result.xlsx for manual inspection\n");

  // Load output and verify key cells
  const wb = new ExcelJS.Workbook();
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  await (wb.xlsx as any).load(outputFile);

  const s1 = wb.getWorksheet("Jan 26")!;
  const s2 = wb.getWorksheet("Sep 25")!;
  const s3 = wb.getWorksheet("Apr 25")!;

  function cellVal(sheet: ExcelJS.Worksheet, row: number, col: number): string {
    const cell = sheet.getRow(row).getCell(col);
    const val = cell.value;
    if (val == null) return "(empty)";
    if (typeof val === "object" && "formula" in val) {
      return `={${(val as ExcelJS.CellFormulaValue).formula}}`;
    }
    if (typeof val === "object" && "sharedFormula" in val) {
      return `=shared(${(val as { sharedFormula: string }).sharedFormula})`;
    }
    return String(val);
  }

  function dumpRow(label: string, sheet: ExcelJS.Worksheet, row: number) {
    const cols = [1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13];
    const names = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "K", "L", "M"];
    console.log(`  ${label} (row ${row}):`);
    for (let i = 0; i < cols.length; i++) {
      const v = cellVal(sheet, row, cols[i]);
      if (v !== "(empty)") {
        console.log(`    ${names[i]}: ${v}`);
      }
    }
    console.log();
  }

  // ── Block structure (original template from generate-test-files.ts) ──
  //
  // Block 1 (7217 Bridlewood): header=1, data 2-4, payments 5-7, buffer 8-27, 总开销=28
  //   Row 2: Chi Man (E=阿华), Row 3: 7217 (E=阿华), Row 4: Burnaby (D/E empty, K=8400)
  // Block 2 (1530 Oak): header=36, data 37-38, payments 39-41, buffer 42-49, 总开销=50
  // Block 3 (5880 Dover): header=55, data 56-59, payments 60-62, NO buffer, 总开销=63
  //   Row 56: E=老赵, Row 57: Richmond (D/E empty, K=78.62), Row 58: E=阿华, Row 59: E=老赵
  // Block 4 (302 Marco): header=70, data 71, payments 72-74, buffer 75-89, 总开销=90
  // Block 5 (53 Maple): header=95, data 96, payments 97-99, buffer 100-114, 总开销=115
  //
  // BEHAVIOR:
  //   - Worker data (D/E/F/G/H/I) and material data (K/L/M) are INDEPENDENT.
  //     A row is available for worker data if D and E are both empty — existing
  //     A/B/C/K data on the row does not prevent writing worker data there.
  //   - Payment rows (第一笔款/第二笔款/第三笔款) are available for worker data
  //     if their D and E columns are empty. Payment info (A/B) is independent.
  //   - Materials-only entries (no hours/gas/ticket) do NOT create a new row;
  //     materials are written to the first empty K cell in the block.
  //   - Col C (project name) is NEVER written on new rows.
  //   - Only insert before 总开销 if truly no rows with empty D+E exist in the block.
  //
  // Block 1: Row 4 (Burnaby, D/E empty) → 阿华. Row 5 (第一笔款, D/E empty) → 老赵.
  //   No insertion, no shift. Materials packed: K=200 at row 6, K=150 at row 7.
  // Block 2: Untouched, no shift.
  // Block 3: Row 57 (Richmond, D/E empty) → 阿华. Row 60 (第一笔款, D/E empty) → 老赵.
  //   No insertion, no shift. K=800 packed to row 58.
  // Block 4: Row 72 (第一笔款, D/E empty) → 老赵. No shift.
  // Block 5: K=500 packed to row 96. Row 97 (第一笔款, D/E empty) → 老赵. No shift.
  // Older sheets: Payment row 3 (第一笔款, D/E empty) used. Materials packed to row 2.

  console.log("=== CELL VERIFICATION ===\n");

  // ── Block 1: 7217 Bridlewood ──
  console.log("── Block 1: 7217 Bridlewood (阿华 fills Burnaby row, 老赵 fills 第一笔款 row) ──");
  dumpRow("阿华 fills row 4 (Burnaby, D/E was empty)", s1, 4);
  dumpRow("老赵 fills row 5 (第一笔款, D/E was empty)", s1, 5);
  dumpRow("第二笔款 (mat K=200 packed)", s1, 6);
  dumpRow("第三笔款 (mat K=150 packed)", s1, 7);

  // ── Block 2: untouched ──
  console.log("── Block 2: 1530 Oak Street — untouched (no shift) ──");
  dumpRow("Existing data row 1", s1, 37);

  // ── Block 3: 5880 Dover ──
  console.log("── Block 3: 5880 Dover (阿华 fills Richmond row, 老赵 fills 第一笔款 row) ──");
  dumpRow("阿华 fills row 57 (Richmond, D/E was empty)", s1, 57);
  dumpRow("K=800 packed to row 58", s1, 58);
  dumpRow("老赵 fills row 60 (第一笔款, D/E was empty)", s1, 60);
  dumpRow("总开销 at row 63 (no shift)", s1, 63);

  const expFormula = cellVal(s1, 63, 8);
  console.log(`  总开销 H formula: ${expFormula}`);
  console.log(`  (should be SUM(H56:H62))\n`);

  // ── Block 4: Marco Project ──
  console.log("── Block 4: Marco Project (老赵 fills 第一笔款 row) ──");
  dumpRow("老赵 @ 302 (第一笔款 row)", s1, 72);

  // ── Block 5: 53 Maple Ave ──
  console.log("── Block 5: 53 Maple Ave (K packed, 老赵 fills 第一笔款 row) ──");
  dumpRow("K=500 packed to row 96", s1, 96);
  dumpRow("老赵 @ 53 (第一笔款 row)", s1, 97);

  // ── Older sheets ──
  console.log("── Sheet 'Sep 25': 9000 Ash Grove (第一笔款 row used) ──");
  dumpRow("K=300 packed to row 2", s2, 2);
  dumpRow("阿华 @ 9000 (第一笔款 row)", s2, 3);

  console.log("── Sheet 'Apr 25': 6033 Williams Rd (第一笔款 row used) ──");
  dumpRow("K=1200 packed to row 2", s3, 2);
  dumpRow("Chris @ 6033 (第一笔款 row)", s3, 3);

  // Non-date sheets
  const bal = wb.getWorksheet("Balance")!;
  const t4a = wb.getWorksheet("T4a")!;
  console.log("── Non-date sheets ──");
  console.log(`  Balance row 1 col A: ${cellVal(bal, 1, 1)} (should be "Income")`);
  console.log(`  T4a row 1 col A: ${cellVal(t4a, 1, 1)} (should be "所有人的")\n`);

  // ── PASS/FAIL ──
  console.log("=== PASS/FAIL CHECKS ===");
  const checks: [string, boolean][] = [
    ["Workers parsed = 3", result.workers.length === 3],
    ["Rows added = 8", result.rowsAdded === 8],
    ["Unmatched = [8888]", result.unmatchedAddresses.length === 1 && result.unmatchedAddresses[0] === "8888"],
    ["Matched 3 sheets", result.matchedSheets.length === 3],

    // Block 1: 阿华 fills Burnaby row (D/E was empty), 老赵 fills 第一笔款 row 5.
    // Materials packed: K=200 at row 6 (第二笔款), K=150 at row 7 (第三笔款).
    ["Burnaby C preserved at row 4", cellVal(s1, 4, 3) === "Burnaby"],
    ["阿华@7217 → row 4, D=3月下", cellVal(s1, 4, 4) === "3月下"],
    ["阿华@7217 → row 4, E=阿华", cellVal(s1, 4, 5) === "阿华"],
    ["阿华@7217 → row 4, F=8", cellVal(s1, 4, 6) === "8"],
    ["阿华@7217 → row 4, I=45", cellVal(s1, 4, 9) === "45"],
    ["阿华@7217 mat → 第二笔款 K=200", cellVal(s1, 6, 11) === "200"],
    ["阿华@7217 mat-only → 第三笔款 K=150", cellVal(s1, 7, 11) === "150"],
    ["老赵@7217 → row 5 (第一笔款), E=老赵", cellVal(s1, 5, 5) === "老赵"],
    ["老赵@7217 → row 5, I=35", cellVal(s1, 5, 9) === "35"],
    ["第一笔款 A preserved at row 5", cellVal(s1, 5, 1) === "第一笔款"],
    ["H formula: row 4 H=F4*G4", cellVal(s1, 4, 8) === "={F4*G4}"],
    ["H formula: row 5 H=F5*G5", cellVal(s1, 5, 8) === "={F5*G5}"],

    // Block 2: untouched (no shift)
    ["Block 2 data at row 37", cellVal(s1, 37, 5) === "老赵"],

    // Block 3: 阿华 fills Richmond row (D/E empty), 老赵 fills 第一笔款 row 60.
    // No insertion, no shift. K=800 packed to row 58.
    ["阿华@5880 fills row 57, E=阿华", cellVal(s1, 57, 5) === "阿华"],
    ["阿华@5880 fills row 57, F=6", cellVal(s1, 57, 6) === "6"],
    ["老赵@5880 fills row 60 (第一笔款), E=老赵", cellVal(s1, 60, 5) === "老赵"],
    ["老赵@5880 fills row 60, F=15", cellVal(s1, 60, 6) === "15"],
    ["老赵@5880 K=800 packed to row 58", cellVal(s1, 58, 11) === "800"],
    ["第一笔款 A preserved at row 60", cellVal(s1, 60, 1) === "第一笔款"],
    ["总开销 at row 63 (no shift)", cellVal(s1, 63, 1) === "总开销"],
    ["总开销 H formula = SUM(H56:H62)", expFormula === "={SUM(H56:H62)}"],

    // Block 4: no shift, 老赵 fills 第一笔款 row 72.
    ["302→row 72 (第一笔款), E=老赵", cellVal(s1, 72, 5) === "老赵"],
    ["302→row 72, F=20", cellVal(s1, 72, 6) === "20"],

    // Block 5: no shift. K=500 packed to row 96. 老赵 fills 第一笔款 row 97.
    ["53→row 97 (第一笔款), E=老赵", cellVal(s1, 97, 5) === "老赵"],
    ["53→row 97, F=12", cellVal(s1, 97, 6) === "12"],
    ["53 K=500 packed to row 96", cellVal(s1, 96, 11) === "500"],

    // Older sheets — 第一笔款 row 3 used, materials packed to row 2
    ["9000→'Sep 25' row 3, D=3月下", cellVal(s2, 3, 4) === "3月下"],
    ["9000→'Sep 25' row 3, E=阿华", cellVal(s2, 3, 5) === "阿华"],
    ["6033→'Apr 25' row 3, D=3月下", cellVal(s3, 3, 4) === "3月下"],
    ["6033→'Apr 25' row 3, E=Chris", cellVal(s3, 3, 5) === "Chris"],

    // Non-date sheets
    ["Balance untouched", cellVal(bal, 1, 1) === "Income"],
    ["T4a untouched", cellVal(t4a, 1, 1) === "所有人的"],
  ];

  let passed = 0;
  let failed = 0;
  for (const [name, ok] of checks) {
    console.log(`  ${ok ? "✅" : "❌"} ${name}`);
    if (ok) passed++;
    else failed++;
  }
  console.log(`\n${passed}/${checks.length} checks passed${failed > 0 ? ` (${failed} FAILED)` : ""}`);
}

main().catch(console.error);
