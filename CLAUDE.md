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

## Architecture

Next.js 16 app (React 19, TypeScript, Tailwind CSS 4) that migrates construction labor/material cost data from an input Excel sheet into per-project output Excel sheets.

### Data Flow

1. **Client** (`app/page.tsx`, client component) — upload UI with two file drop zones + date label input, sends FormData to API, displays processing summary, triggers file download
2. **API Route** (`app/api/process/route.ts`) — receives input file, output file, and date label; calls processor; returns JSON with metadata + base64-encoded output file
3. **Excel Processor** (`lib/excel-processor.ts`) — core logic using ExcelJS (server-only, marked in `next.config.ts` as `serverExternalPackages`)

### Input File Format (Sheet1)

Worker blocks separated by "Total" rows:
- **B**: Worker name, **C**: $/hr rate, **D**: Address code (e.g. "5793", "7217"), **E**: Hours, **G**: Materials, **H**: Gas, **I**: Ticket

### Output File Format (one sheet per project)

Sheet names start with address codes (e.g. "7217 Bridlewood", "1171 Jervis Street"). Columns:
- **D**: 日期, **E**: 工人, **F**: 工时, **G**: hourly rate, **H**: =F×G formula, **I**: 其他开销, **K**: 材料开销, **L**: notes

Sheets contain payment rows (第一笔款/第二笔款/第三笔款) and summary rows (总开销/总价格/利润率) with formulas.

### Processing Logic

1. Parse input into `WorkerBlock[]` — each block has worker name, rate, and per-address entries
2. Match address codes to output sheet name prefixes
3. Skip entries where all four data fields are zero (hours, materials, gas, ticket)
4. For each match, find the first empty row where D/E/F/G are all empty (searching from row 2 up to 总开销)
5. Write date label and worker name on every data row (even expense-only or material-only entries)
6. Write hours, rate, and `=F*G` formula only when hours > 0
7. Write gas+ticket to column I, materials to column K (overflow to next empty K row, bounded by summary row)
8. If no empty rows remain, insert before 总开销 and update formulas on all three summary rows (总开销, 总价格, 利润率)

### Validation & Limits

- `MAX_FILE_SIZE` (50 MB) is exported from `lib/excel-processor.ts` and used in both the processor and the API route
- API route validates: `instanceof File`, `.xlsx` extension, file size, and `typeof` check for dateLabel
- Material search is bounded by the 总开销 row — will never write into or past summary rows

### Key Helpers (`lib/excel-processor.ts`)

- `loadWorkbook(buffer)` — wraps the ExcelJS `any` cast for `xlsx.load()` in one place (TS 6 type mismatch)
- `findSummaryRows(sheet)` — locates 总开销, 总价格, 利润率 row numbers
- `shiftFormulaReferences(formula, fromRow)` — increments cell references at or after `fromRow` by 1 (used after `spliceRows`)
- `writeEntryToSheet(sheet, worker, entry, dateLabel)` — writes a single worker entry to the matched output sheet

## Key Constraints

- ExcelJS types don't align with Node.js Buffer in TS 6 — consolidated into `loadWorkbook()` helper with a single `(wb.xlsx as any).load()` cast
- Output file is returned as base64 in JSON (not streamed) for simplicity
- Formula cells in the input file use cached results (`cell.value.result`) since ExcelJS doesn't evaluate formulas
- Client imports shared types from `lib/excel-processor.ts` via `import type` (erased at compile time, no server code bundled)
