# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Build & Dev Commands

```bash
npm run dev       # Start Next.js dev server (http://localhost:3000)
npm run build     # Production build (also runs TypeScript type-checking)
npm start         # Start production server
```

No test framework is configured. Verify changes by running `npm run build` and testing the API endpoint with curl:
```bash
curl -X POST http://localhost:3000/api/process \
  -F "input=@input-example.xlsx" \
  -F "output=@output-example.xlsx" \
  -F "dateLabel=3月下"
```

### Test Scripts

```bash
npx tsx scripts/generate-test-files.ts  # Generate test-input.xlsx + test-output-template.xlsx
npx tsx scripts/smoke-test.ts           # Run processor and verify 36 edge-case checks
```

Generated test files cover: normal matching, materials-only, hours-only, gas+ticket-only, all-zeros skip, unmatched address, short-code false-positive protection, col A fallback, older sheet matching, newest-sheet priority, block-full insertion with subsequent block shifting, non-date sheet preservation, and payment row filling (worker data written to payment rows with empty D/E).

## Architecture

Next.js 16 app (React 19, TypeScript, Tailwind CSS 4) that migrates construction labor/material cost data from an input Excel sheet into date-range output Excel sheets organized by project blocks.

### Data Flow

1. **Client** (`app/page.tsx`, client component) — upload UI with two file drop zones + date label input, sends FormData to API, displays processing summary (per-entry details, employee totals, site totals, input warnings), triggers file download
2. **API Route** (`app/api/process/route.ts`) — receives input file, output file, and date label; calls processor; returns JSON with metadata + base64-encoded output file
3. **Excel Processor** (`lib/excel-processor.ts`) — core logic using ExcelJS (server-only, marked in `next.config.ts` as `serverExternalPackages`)

### Input File Format (Sheet1)

Worker blocks separated by "Total" rows:
- **B**: Worker name, **C**: $/hr rate, **D**: Address code (e.g. "5793", "7217"), **E**: Hours, **G**: Materials, **H**: Gas, **I**: Ticket

### Output File Format (date-range sheets with project blocks)

Sheets are named by date ranges: `{startMonth}-{endMonth} {2-digit year}` (e.g. "1-3 26" = Jan–Mar 2026, "9-12 25" = Sep–Dec 2025). Non-date sheets (Balance, T4a) are skipped.

Each date-range sheet contains ~11 project blocks. Each block:
```
Row N:         Header (单号, 负责人, 客户名+项目, 日期, 工人, 工时, hourly rate, 开销, 其他开销, Note, 材料开销)
Row N+1..N+k:  Data rows (C=project name, D=date, E=worker, F=hours, G=rate, H=F*G shared formula, I=other, K=materials, L=notes, M=date)
                Payment rows (第一笔款/第二笔款/第三笔款 in col A, B=amount; may also have K/L/M material data)
                Buffer rows (empty, with =F*G shared formula in col H)
Row M:         总开销 (B=SUM(H:K), H/I/K=SUM over data range)
Row M+1:       总价格 (B=SUM of payment amounts)
Row M+2:       利润率 (B=(总价格-总开销)/总价格)
```

Multiple project names appear in column C within a single block (client names, street addresses, cities). Column A may hold unit/address numbers on data rows.

### Processing Logic

1. Parse input into `WorkerBlock[]` with warnings — each block has worker name, rate, and per-address entries; rows with data but missing address or worker context generate warnings
2. Get date-range sheets sorted most recent first; pre-compute project blocks per sheet
3. Skip entries where all four data fields are zero (hours, materials, gas, ticket)
4. For each entry, search sheets (most recent first) for a block whose col C values match the address code (word-boundary match via regex `\b`); fallback to col A values
5. Break shared formulas in col H before processing (converts shared formula clones to explicit `=Fn*Gn` to prevent ExcelJS errors when rows shift)
6. Within matched block, find first row where D (date) and E (worker name) are both empty — payment rows are available (payment info in A/B is independent of worker data in D/E/F/G/H/I); rows with existing dates are project info rows, not available for new workers; material data (K) is independent and does not affect worker placement
7. Materials-only entries (hours=0, gas=0, ticket=0, materials>0) write to the first empty K cell in the block; if no empty K cell exists, insert a new row before 总开销
8. For non-materials-only entries: write col D = date label, col E = worker name (col C is never written)
9. Write hours, rate only when hours > 0; preserve existing =F*G formula in col H
10. Write gas+ticket to col I, materials to col K (search from `dataStartRow` to pack contiguously; if no empty K cell in block, insert a new row before 总开销)
11. If no rows with empty D+E in entire block (including payment and buffer rows) → insert before 总开销, add explicit =F*G to col H, update summary formulas, shift 总价格/利润率 references, and shift row indices for all subsequent blocks on the same sheet

### Validation & Limits

- `MAX_FILE_SIZE` (50 MB) is exported from `lib/excel-processor.ts` and used in both the processor and the API route
- API route validates: `instanceof File`, `.xlsx` extension, zero-byte rejection, file size (returns HTTP 413 for oversized files), and `typeof` check for dateLabel
- Date label is sanitized via `sanitizeDateLabel()` to strip leading `=`, `+`, `-`, `@`, tab, CR characters (formula injection protection)
- Generic error message returned on 500 (internal details logged server-side only)
- Material search is bounded by the block's 总开销 row — will never write into or past summary rows

### Constants (`lib/excel-processor.ts`)

- `IN_COL` — input file column indices: `NAME`(B), `RATE`(C), `ADDRESS`(D), `HOURS`(E), `MATERIALS`(G), `GAS`(H), `TICKET`(I)
- `OUT_COL` — output file column indices: `LABEL`(A), `AMOUNT`(B), `PROJECT`(C), `DATE`(D), `WORKER`(E), `HOURS`(F), `RATE`(G), `COST`(H), `OTHER`(I), `MATERIAL`(K), `NOTES`(L), `MAT_DATE`(M)
- `PAYMENT_LABELS` — Set of payment row labels (第一笔款/第二笔款/第三笔款); used only in `findProjectBlocks` to skip payment rows when collecting project names, NOT for row availability
- `CENTURY_PREFIX` — 2000, used in `parseDateRange` to convert 2-digit year
- `colLetter(index)` — converts 1-based column index to Excel letter (1→A, 8→H, 11→K)
- `COL` — precomputed column letters derived from `OUT_COL`: `HOURS`(F), `RATE`(G), `COST`(H), `OTHER`(I), `MATERIAL`(K); used in formula strings to avoid hardcoded letters

### Key Helpers (`lib/excel-processor.ts`)

- `loadWorkbook(buffer)` — wraps the ExcelJS `any` cast for `xlsx.load()` in one place (TS 6 type mismatch)
- `parseDateRange(name)` — parses sheet name "1-3 26" → `{ startMonth, endMonth, year }` or null
- `getDateRangeSheets(workbook)` — filters to date-range sheets, sorted most recent first
- `findProjectBlocks(sheet)` — scans for 单号 header rows, builds `ProjectBlock[]` with data range, summary rows, and col C/A values
- `matchesAddressCode(text, addressCode)` — word-boundary regex match (prevents "53" matching "1530")
- `findMatchingBlock(blocks, addressCode)` — word-boundary match on col C then col A; returns matched `ProjectBlock` or null
- `isCellEmpty(cell)` — checks if a cell is null, empty string, or zero (used for K-column emptiness)
- `setMaterialCells(row, materials, workerName, dateLabel)` — writes K/L/M columns for a material entry
- `isRowAvailableForWorker(sheet, r)` — checks if D and E are both empty; rows with existing dates are project info rows, material data (K) is independent
- `shiftFormulaReferences(formula, fromRow)` — increments cell references at or after `fromRow` by 1; skips `$`-anchored absolute row refs
- `spliceBlockRow(sheet, block, allBlocks, spliceAt)` — inserts a row, fixes formulas, shifts subsequent blocks, updates summary SUM ranges
- `breakSharedFormulasInColH(sheet)` — converts all shared formula clones/masters in col H to explicit `=Fn*Gn` formulas (prevents ExcelJS errors on row shifts)
- `fixFormulasAfterSplice(sheet, spliceAt)` — shifts all formula references on the sheet after a row insertion at `spliceAt`; handles leftover shared formulas in col H
- `findOrCreateInsertRow(sheet, block, allBlocks)` — searches entire block for rows where D and E are both empty (including payment and buffer rows); inserts before 总开销 only if no empty D+E rows exist
- `writeMaterialToBlock(sheet, block, allBlocks, workerName, materials, dateLabel)` — writes material cost to first empty K cell in block; inserts a new row before 总开销 if no empty K cell exists
- `writeEntryToBlock(sheet, block, allBlocks, worker, entry, dateLabel)` — writes a single worker entry into a project block; shifts subsequent blocks when inserting rows
- `computeCostSummaries(workers, unmatchedAddresses)` — computes per-worker and per-site cost breakdowns; excludes unmatched and all-zero entries; splits gas from ticket; filters out workers with zero matched total
- `processWorkerEntries(workers, dateSheets, sheetBlocks, dateLabel)` — main worker entry processing loop (extracted from `processFiles`); iterates workers/entries, matches to blocks, writes data

### Exported Types (`lib/excel-processor.ts`)

- `WorkerEntry` — per-address entry: address, hours, materials, gas, ticket
- `WorkerBlock` — worker name, rate, and entries array
- `CostSummary` — base: labor, materials, gas, other (ticket), total
- `WorkerTotal extends CostSummary` — adds name and rate
- `SiteTotal extends CostSummary` — adds address
- `ProcessResult` — workers, matchedSheets, unmatchedAddresses, rowsAdded, warnings, workerTotals, siteTotals

## Key Constraints

- ExcelJS types don't align with Node.js Buffer in TS 6 — consolidated into `loadWorkbook()` helper with a single `(wb.xlsx as any).load()` cast
- Output file is returned as base64 in JSON (not streamed) for simplicity
- Formula cells in the input file use cached results (`cell.value.result`) since ExcelJS doesn't evaluate formulas
- Client imports shared types from `lib/excel-processor.ts` via `import type` (erased at compile time, no server code bundled)
