import type { NextApiRequest, NextApiResponse } from "next";
import fs from "fs";
import PDFDocument from "pdfkit";
import * as xlsx from "xlsx";
import type { File, Fields } from "formidable";

type PDFKitDocument = InstanceType<typeof PDFDocument>;
type Workbook = xlsx.WorkBook;
type Worksheet = xlsx.WorkSheet;

type SheetSummary = {
  name: string;
  hidden: boolean;
  range: string;
  rowCount: number;
  columnCount: number;
  nonEmptyRowCount: number;
  sampleRows: string[][];
  merges: number;
};

type FileAnalysis = {
  mode: "groq" | "fallback";
  selectedSheets: string[];
  ignoredSheets: string[];
  sheetLayoutMeta: Record<string, "VOUCHER" | "REPORT">;
  reason: string;
  warnings: string[];
};

type RowSection = "HEADER" | "META" | "TABLE_HEADER" | "TABLE_BODY" | "SUMMARY" | "FOOTER";

interface ClassifiedRow {
  section: RowSection;
  rowIndex: number;
  cells: string[];
}

export const config = {
  api: { bodyParser: false },
};

// ─── Helpers ────────────────────────────────────────────────────────────────

const parseMultipart = (req: NextApiRequest): Promise<{ fields: Fields; files: File[] }> => {
  return new Promise(async (resolve, reject) => {
    try {
      const formidableModule = await import("formidable");
      const form = formidableModule.default({ multiples: true, keepExtensions: true, maxFileSize: 200 * 1024 * 1024 });
      form.parse(req, (err, fields, files) => {
        if (err) return reject(err);
        const fileList = Object.values(files ?? {})
          .flatMap((field) => (Array.isArray(field) ? field : field ? [field] : []))
          .filter(Boolean) as File[];
        resolve({ fields: fields as Fields, files: fileList });
      });
    } catch (error) {
      reject(error);
    }
  });
};

let groqClientPromise: Promise<any | null> | null = null;
const getGroqClient = async () => {
  if (!process.env.GROQ_API_KEY) return null;
  if (!groqClientPromise) {
    groqClientPromise = import("groq-sdk").then((m) => new m.default({ apiKey: process.env.GROQ_API_KEY! }));
  }
  return groqClientPromise;
};

const getCellValue = (sheet: xlsx.WorkSheet, row: number, col: number): string => {
  const cell = sheet[xlsx.utils.encode_cell({ r: row, c: col })];
  if (!cell) return "";
  if (cell.w != null) return String(cell.w);
  return String(cell.v ?? "");
};

const normalizeCellValue = (value: unknown) => String(value ?? "").replace(/\s+/g, " ").trim();

const tryParseDate = (value: string): Date | null => {
  const normalized = String(value || "").trim();
  if (!normalized) return null;
  // First try standard JS date parsing (handles ISO, US, EU formats)
  const parsed = Date.parse(normalized);
  if (!Number.isNaN(parsed)) {
    const date = new Date(parsed);
    const year = date.getUTCFullYear();
    // Check if year is within reasonable historical/future bounds
    if (year >= 1900 && year <= 2100) return date;
  }
  // Try Excel serial date format (numeric days since 1899-12-30)
  // Excel serial dates are typically 5-6 digit numbers (e.g., 44927 for 2023-01-01)
  const n = Number(normalized);
  if (!Number.isNaN(n) && n > 0 && n < 60000 && /^[0-9]{5,6}$/.test(normalized)) {
    // Excel epoch: 1899-12-30 (Excel has a leap year bug that treats 1900 as leap year)
    const excelEpoch = new Date(Date.UTC(1899, 11, 30));
    const dateMs = excelEpoch.getTime() + n * 86400000;
    const date = new Date(dateMs);
    const year = date.getUTCFullYear();
    if (year >= 1900 && year <= 2100) return date;
  }
  return null;
};

const isRowEmpty = (row: string[]) => row.every((c) => String(c || "").trim() === "");

const detectDateColumn = (rows: string[][]): number | null => {
  if (rows.length < 2 || rows[0].length === 0) return null;
  const scores = new Array(rows[0].length).fill(0);
  rows[0].forEach((v, i) => { if (/date|time|day|created|updated|due/.test(v.toLowerCase())) scores[i] += 10; });
  rows.slice(1, 25).filter((r) => !isRowEmpty(r)).forEach((row) => {
    row.forEach((v, i) => { if (tryParseDate(v)) scores[i] += 1; });
  });
  let best = 0;
  for (let i = 1; i < scores.length; i++) if (scores[i] > scores[best]) best = i;
  return scores[best] >= 3 ? best : null;
};

const sortRowsByDate = (rows: string[][]) => {
  if (rows.length < 2) return { rows, sortedBy: undefined };
  const dc = detectDateColumn(rows);
  if (dc === null) return { rows, sortedBy: undefined };
  const [header, ...body] = rows;
  const sortedBody = [...body].sort((a, b) => {
    const da = tryParseDate(a[dc]), db = tryParseDate(b[dc]);
    if (da && db) return da.getTime() - db.getTime();
    return da ? -1 : db ? 1 : 0;
  });
  return { rows: [header, ...sortedBody], sortedBy: header[dc] || `column ${dc + 1}` };
};

const getSheetDate = (sheet: xlsx.WorkSheet): Date | null => {
  const range = sheet["!ref"] ? xlsx.utils.decode_range(sheet["!ref"]) : null;
  if (!range) return null;
  for (let r = range.s.r; r <= range.e.r; r++)
    for (let c = range.s.c; c <= range.e.c; c++) {
      const date = tryParseDate(getCellValue(sheet, r, c));
      if (date) return date;
    }
  return null;
};

const getSheetRows = (sheet: xlsx.WorkSheet): string[][] => {
  const range = sheet["!ref"] ? xlsx.utils.decode_range(sheet["!ref"]) : { s: { r: 0, c: 0 }, e: { r: 0, c: 0 } };
  const rows: string[][] = [];
  for (let r = range.s.r; r <= range.e.r; r++) {
    const row: string[] = [];
    for (let c = range.s.c; c <= range.e.c; c++) row.push(getCellValue(sheet, r, c));
    rows.push(row);
  }
  return rows;
};

const getColumnCount = (rows: string[][]) => rows.reduce((max, r) => Math.max(max, r.length), 0);
const toFiniteNumber = (v: number, fallback: number) => (Number.isFinite(v) ? v : fallback);

const getWorkbookSheetMeta = (workbook: Workbook, sheetName: string) => {
  const meta = (workbook.Workbook?.Sheets ?? [])[workbook.SheetNames.indexOf(sheetName)];
  return { hidden: Boolean(meta && meta.Hidden) };
};

const summarizeSheet = (workbook: Workbook, sheetName: string, sheet: Worksheet): SheetSummary => {
  const rows = getSheetRows(sheet);
  const nonEmptyRows = rows.filter((r) => !isRowEmpty(r));
  return {
    name: sheetName,
    hidden: getWorkbookSheetMeta(workbook, sheetName).hidden,
    range: sheet["!ref"] || "A1:A1",
    rowCount: rows.length,
    columnCount: getColumnCount(rows),
    nonEmptyRowCount: nonEmptyRows.length,
    sampleRows: nonEmptyRows.slice(0, 8).map((r) => r.slice(0, 8).map(normalizeCellValue)),
    merges: sheet["!merges"]?.length || 0,
  };
};

const fallbackAnalyzeWorkbook = (fileName: string, workbook: Workbook, summaries: SheetSummary[]): FileAnalysis => {
  const visible = summaries.filter((s) => !s.hidden && s.nonEmptyRowCount > 0);
  const reportLike = visible.filter((s) => /report|summary|invoice|statement|payment|print|final/.test(s.name.toLowerCase()));
  const selected = (reportLike.length > 0 ? reportLike : visible).slice(0, 1).map((s) => s.name);
  return {
    mode: "fallback",
    selectedSheets: selected.length > 0 ? selected : workbook.SheetNames.slice(0, 1),
    ignoredSheets: summaries.map((s) => s.name).filter((n) => !selected.includes(n)),
    sheetLayoutMeta: Object.fromEntries(summaries.map((s) => [s.name, s.merges > s.rowCount * 0.2 ? "VOUCHER" : "REPORT"])),
    reason: `Local fallback selected the most likely visible sheet(s) for "${fileName}".`,
    warnings: process.env.GROQ_API_KEY
      ? ["Groq analysis was unavailable; local sheet selection used."]
      : ["Groq not configured. Set GROQ_API_KEY to enable AI sheet selection."],
  };
};

const extractJsonObject = (value: string) => {
  const start = value.indexOf("{"), end = value.lastIndexOf("}");
  if (start === -1 || end === -1 || end <= start) throw new Error("No JSON object found in Groq response.");
  return JSON.parse(value.slice(start, end + 1));
};

const analyzeWorkbookWithGroq = async (fileName: string, summaries: SheetSummary[]): Promise<FileAnalysis> => {
  const groq = await getGroqClient();
  if (!groq) throw new Error("Groq client not configured.");
  const model = process.env.GROQ_MODEL || "llama-3.3-70b-versatile";
  const prompt = [
    "You are analyzing an Excel workbook before PDF export. Return only JSON.",
    "Pick sheets most likely for final PDF output. Prefer visible, meaningful content sheets.",
    "Select exactly one primary sheet. Categorize as 'VOUCHER' (forms/invoices) or 'REPORT' (data grids).",
    'Schema: {"selectedSheets":string[],"ignoredSheets":string[],"sheetLayoutMeta":Record<string,"VOUCHER"|"REPORT">,"reason":string,"warnings":string[]}',
    `File: ${fileName}`, `Summaries: ${JSON.stringify(summaries)}`,
  ].join("\n");
  const response = await groq.chat.completions.create({
    model, temperature: 0.2,
    messages: [{ role: "system", content: "Return only valid JSON." }, { role: "user", content: prompt }],
  });
  const content = response.choices[0]?.message?.content;
  if (!content) throw new Error("Groq returned empty response.");
  const parsed = extractJsonObject(content) as Partial<FileAnalysis>;
  const isStr = (v: unknown): v is string => typeof v === "string";
  const selected = (Array.isArray(parsed.selectedSheets) ? parsed.selectedSheets.filter(isStr) : []);
  const ignored = (Array.isArray(parsed.ignoredSheets) ? parsed.ignoredSheets.filter(isStr) : []);
  const warnings = (Array.isArray(parsed.warnings) ? parsed.warnings.filter(isStr) : []);
  const available = new Set(summaries.map((s) => s.name));
  const normalizedSelected = selected.filter((n) => available.has(n)).slice(0, 1);
  if (normalizedSelected.length === 0) throw new Error("Groq did not select any valid sheets.");
  return {
    mode: "groq",
    selectedSheets: normalizedSelected,
    ignoredSheets: ignored.filter((n) => available.has(n) && !normalizedSelected.includes(n)),
    sheetLayoutMeta: typeof parsed.sheetLayoutMeta === "object" ? (parsed.sheetLayoutMeta as Record<string, "VOUCHER" | "REPORT">) : {},
    reason: typeof parsed.reason === "string" && parsed.reason.trim() ? parsed.reason.trim() : `Groq selected ${normalizedSelected.join(", ")}.`,
    warnings,
  };
};

const analyzeWorkbook = async (fileName: string, workbook: Workbook) => {
  const summaries = workbook.SheetNames.map((n) => summarizeSheet(workbook, n, workbook.Sheets[n]));
  try {
    return { summaries, analysis: await analyzeWorkbookWithGroq(fileName, summaries) };
  } catch (error) {
    const fallback = fallbackAnalyzeWorkbook(fileName, workbook, summaries);
    fallback.warnings = [...fallback.warnings, `Groq failed: ${error instanceof Error ? error.message : String(error)}`];
    return { summaries, analysis: fallback };
  }
};

const encodeHeaderValue = (value: string) =>
  encodeURIComponent(value.length > 7500 ? value.slice(0, 7500) + "..." : value);

const getRenderableSheetNames = (workbook: Workbook, summaries: SheetSummary[], _analysis: FileAnalysis) =>
  workbook.SheetNames
    .map((name) => ({ name, hidden: summaries.find((s) => s.name === name)?.hidden ?? false, date: getSheetDate(workbook.Sheets[name]) }))
    .filter((s) => !s.hidden)
    .sort((a, b) => (a.date && b.date ? a.date.getTime() - b.date.getTime() : a.date ? -1 : b.date ? 1 : 0))
    .map((s) => s.name);

// ═══════════════════════════════════════════════════════════════════════════
// ROW CLASSIFIER
// ═══════════════════════════════════════════════════════════════════════════

const isTableHeaderRow = (row: string[]) =>
  /particulars|description|item\b|amount|rate|qty|quantity|credit|debit|dr\b|cr\b|sl\.?\s*no/i.test(row.map((v) => v.toLowerCase()).join(" "));

const isMetaRow = (row: string[]) =>
  /^(no|date|through|ca\.?\s*no|mode|instrument|party|a\/c|utr|bank|ifsc|branch|cheque|ch[\.:])[\s]*[:\-]/i.test(
    row.filter((v) => v.trim()).join(" ").trim()
  );

const isTotalRow = (row: string[]) =>
  row.some((v) => /^\s*total\s*$/i.test(v.trim()) || /grand\s*total/i.test(v));

const isAmountWordsRow = (row: string[]) => {
  const joined = row.filter((v) => v.trim()).join(" ");
  return /amount\s*\(?\s*in\s*words?\s*\)?/i.test(joined) || /^(INR|Rs)[\.\s]/i.test(joined.trim());
};

const isSignatureRow = (row: string[]) =>
  /prepared\s*by|checked\s*by|authorized|authorised|signature|receiver|approved|varified|verified/i.test(
    row.filter((v) => v.trim()).join(" ")
  );

const classifyRows = (rows: string[][], merges: any[]): ClassifiedRow[] => {
  const classified: ClassifiedRow[] = [];
  let tableHeaderFound = false, tableEnded = false, footerStarted = false;
  const colCount = getColumnCount(rows);
  const isFullMerge = (r: number) =>
    merges?.some((m) => m.s.r === r && m.e.r === r && m.e.c - m.s.c + 1 >= colCount - 1) ?? false;

  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (isRowEmpty(row)) {
      classified.push({ section: tableHeaderFound && !tableEnded ? "TABLE_BODY" : "META", rowIndex: r, cells: row });
      continue;
    }
    if (footerStarted) { classified.push({ section: "FOOTER", rowIndex: r, cells: row }); continue; }
    if (isSignatureRow(row)) {
      footerStarted = true; tableEnded = true;
      classified.push({ section: "FOOTER", rowIndex: r, cells: row }); continue;
    }
    if (!tableHeaderFound) {
      if (isTableHeaderRow(row)) {
        tableHeaderFound = true;
        classified.push({ section: "TABLE_HEADER", rowIndex: r, cells: row }); continue;
      }
      const nonEmpty = row.filter((v) => v.trim());
      if (r <= 4 && (isFullMerge(r) || (nonEmpty.length === 1 && r < 4))) {
        classified.push({ section: "HEADER", rowIndex: r, cells: row });
      } else if (isMetaRow(row)) {
        classified.push({ section: "META", rowIndex: r, cells: row });
      } else if (r <= 4) {
        classified.push({ section: "HEADER", rowIndex: r, cells: row });
      } else {
        classified.push({ section: "META", rowIndex: r, cells: row });
      }
    } else {
      if (isTotalRow(row)) { classified.push({ section: "SUMMARY", rowIndex: r, cells: row }); continue; }
      if (isAmountWordsRow(row)) {
        tableEnded = true;
        classified.push({ section: "SUMMARY", rowIndex: r, cells: row }); continue;
      }
      classified.push({ section: tableEnded ? (isMetaRow(row) ? "META" : "SUMMARY") : "TABLE_BODY", rowIndex: r, cells: row });
    }
  }
  return classified;
};

// ═══════════════════════════════════════════════════════════════════════════
// VOUCHER RENDERER
// Produces output identical to the Associated Traders Cash Payment Voucher
// ═══════════════════════════════════════════════════════════════════════════

// ── Page geometry (US Letter 612 × 792 pts) ──────────────────────────────
const PW = 612;
const PH = 792;
const OUTER_M = 36;   // outer border inset from page edge
const TEXT_X = 48;    // text/content left edge
const TEXT_W = PW - TEXT_X * 2;   // 516 pts wide

// ── Table column widths ──────────────────────────────────────────────────
// Particulars ~60%, Amount ~20%, Remarks ~20%  (must sum to TEXT_W = 516)
const COL_P = Math.round(TEXT_W * 0.595);  // ~307
const COL_A = Math.round(TEXT_W * 0.210);  // ~108
const COL_R = TEXT_W - COL_P - COL_A;     // ~101

// Derived X positions
const X_P = TEXT_X;
const X_A = X_P + COL_P;
const X_R = X_A + COL_A;
const X_END = X_P + TEXT_W;

// ── Fonts ────────────────────────────────────────────────────────────────
const FR = "Helvetica";
const FB = "Helvetica-Bold";

// ── Colors ───────────────────────────────────────────────────────────────
const C_TEXT   = "#111827";  // Charcoal Gray
const C_GRAY   = "#4B5563";  // Medium Gray
const C_BORDER = "#D1D5DB";  // Light Border
const C_BG     = "#F9FAFB";  // Subtle BG

// ── Font sizes ───────────────────────────────────────────────────────────
const SZ_CO = 18;   // company name
const SZ_AD = 10;   // address
const SZ_TI = 14;   // voucher title
const SZ_ME = 10;   // meta (No., Date, Paid to)
const SZ_TH = 10;   // table header
const SZ_TB = 9;    // table body
const SZ_TO = 10;   // total row
const SZ_WD = 9;    // amount in words
const SZ_FT = 9;    // footer

// ── Line weights ─────────────────────────────────────────────────────────
const LW_OUTER = 0.8;
const LW_TABLE = 0.4;

// ── Layout constants ─────────────────────────────────────────────────────
const TABLE_BODY_MIN_H = 300;  // minimum body area height (creates the big empty space)
const HEADER_ROW_H    = 18;
const SUB_LABEL_H     = 14;
const TOTAL_ROW_H     = 16;
const SIG_ROW_H       = 45;
const FOOTER_ROW_H    = 20;

interface Renderer {
  render(sheetName: string, rows: string[][], merges: any[]): void;
}

class BaseRenderer {
  protected doc: PDFKitDocument;
  protected y = TEXT_X;

  constructor(doc: PDFKitDocument) {
    this.doc = doc;
  }

  // ── Branding & Borders (Repeatable on every page) ────────────────────
  protected renderPageDecorations(companyName: string, supportLines: string[], docTitle: string) {
    this.outerBorder();
    this.y = TEXT_X + 12;

    // Company Header
    this.txt(companyName, X_P, this.y, TEXT_W, FB, SZ_CO, "center");
    this.y += 24;

    for (const line of supportLines) {
      this.txt(line, X_P, this.y, TEXT_W, FR, SZ_AD, "center", false, C_GRAY);
      this.y += 14;
    }

    this.y += 5;

    // Voucher Title Box
    this.hr(this.y, LW_TABLE, X_P, X_END, C_BORDER);
    this.y += 4;
    this.txt(docTitle, X_P, this.y, TEXT_W, FB, SZ_TI, "center");
    this.y += 18;
    this.hr(this.y, LW_TABLE, X_P, X_END, C_BORDER);
    this.y += 15;
  }

  protected outerBorder() {
    this.doc.save()
      .lineWidth(LW_OUTER).strokeColor(C_BORDER)
      .rect(OUTER_M, OUTER_M, PW - OUTER_M * 2, PH - OUTER_M * 2)
      .stroke().restore();
  }

  protected hr(y: number, lw = LW_TABLE, xStart = X_P, xEnd = X_END, color = C_BORDER) {
    this.doc.save()
      .lineWidth(lw).strokeColor(color)
      .moveTo(xStart, y).lineTo(xEnd, y)
      .stroke().restore();
  }

  protected vl(x: number, y1: number, y2: number, lw = LW_TABLE, color = C_BORDER) {
    this.doc.save()
      .lineWidth(lw).strokeColor(color)
      .moveTo(x, y1).lineTo(x, y2)
      .stroke().restore();
  }

  protected box(x: number, y: number, w: number, h: number, lw = LW_TABLE, color = C_BORDER) {
    this.doc.save()
      .lineWidth(lw).strokeColor(color)
      .rect(x, y, w, h).stroke().restore();
  }

  protected txt(
    s: string, x: number, y: number, w: number,
    font = FR, size = SZ_TB,
    align: "left" | "center" | "right" = "left",
    lineBreak = false,
    color = C_TEXT
  ) {
    this.doc.font(font).fontSize(size).fillColor(color)
      .text(s, x, y, { width: w, align, lineBreak });
  }

  protected textH(s: string, w: number, font = FR, size = SZ_TB): number {
    this.doc.font(font).fontSize(size);
    return toFiniteNumber(this.doc.heightOfString(s, { width: w }), size);
  }
}

class VoucherRenderer extends BaseRenderer implements Renderer {
  constructor(doc: PDFKitDocument) {
    super(doc);
  }

  // ── Primitives ───────────────────────────────────────────────────────


  // ── Parse header rows → companyName, supportLines, documentTitle ─────
  private parseHeader(rows: ClassifiedRow[], sheetName: string) {
    const texts = rows.map((r) => r.cells.filter((c) => c.trim()).join("  ").trim()).filter(Boolean);
    const titleIdx = texts.findIndex((t) =>
      /voucher|invoice|receipt|payment|journal|credit.?note|debit.?note|bill/i.test(t)
    );
    const documentTitle = titleIdx >= 0 ? texts[titleIdx] : "Cash Payment Voucher";
    const nonTitle = texts.filter((_, i) => i !== titleIdx);
    const companyName = [...nonTitle].sort((a, b) => b.length - a.length)[0] || sheetName;
    const supportLines = nonTitle.filter((t) => t !== companyName);
    return { companyName, supportLines, documentTitle };
  }

  // ── Find a value from META rows matching any of the given regexes ────
  private findMetaVal(rows: ClassifiedRow[], patterns: RegExp[]): string {
    for (const row of rows) {
      const joined = row.cells.filter((c) => c.trim()).join("  ").trim();
      for (const pat of patterns) {
        const m = joined.match(pat);
        if (m) return (m[1] ?? "").trim();
      }
    }
    return "";
  }

  private fmtDate(v: string): string {
    const parsed = tryParseDate(v.trim());
    if (!parsed) return v.trim();
    const day = String(parsed.getDate()).padStart(2, "0");
    const month = String(parsed.getMonth() + 1).padStart(2, "0");
    const year = parsed.getFullYear();
    return `${day}-${month}-${year}`;
  }

  // ── Parse No., Date, Paid to, and extra meta detail lines ───────────
  private parseVoucherMeta(rows: ClassifiedRow[]) {
    // First, gather all meta text to search from
    let allMetaText = rows.map((r) => r.cells.join(" ")).join(" | ");
    
    // Extract Date - look for "Date" followed by a date value (handles same-line dates)
    let rawDate = "";
    const dateMatches = allMetaText.match(/\bdate\s*[:\-]?\s*([^\|]+?)(?=\s*(?:through|mode|instrument|paid|received|no:|voucher|details?|\|)|$)/i);
    if (dateMatches) {
      rawDate = dateMatches[1].trim().split(/\s+/)[0]; // take only first token (the date value)
    }
    const voucherDate = this.fmtDate(rawDate);

    // Extract Voucher No - only the numeric part (not including dates/text after)
    let voucherNo = "";
    const noMatch = allMetaText.match(/(?:no\.?|voucher\s*no\.?)\s*[:\-]?\s*(\d+|\w+)/i);
    if (noMatch) {
      voucherNo = noMatch[1].trim();
    }

    // Extract Party Value
    const partyValue = this.findMetaVal(rows, [
      /^paid\s*to\s*[:\-]?\s*(.+)$/i,
      /^received\s*from\s*[:\-]?\s*(.+)$/i,
      /^party'?s?\s*name\s*[:\-]\s*(.+)$/i,
    ]).trim();

    const partyLabel = /received\s*from/i.test(allMetaText) ? "Received from:-" : "Paid to:-";

    // Collect extra detail rows (Through, Mode, Instrument, UTR etc.)
    const extraLines: string[] = [];
    for (const row of rows) {
      const nonEmpty = row.cells.map((c) => c.trim()).filter(Boolean);
      if (nonEmpty.length === 0) continue;
      const joined = nonEmpty.join("  ");
      
      // Skip main meta fields
      if (/^(no\.?|voucher\s*no\.?|date|paid\s*to|received\s*from|party'?s?\s*name)\s*[:\-]/i.test(joined)) continue;
      // Skip if contains extracted values
      if (voucherNo && joined.includes(voucherNo)) continue;
      if (rawDate && joined.includes(rawDate)) continue;
      // Skip standalone date patterns
      if (/^\d{1,2}[-\/]\d{1,2}[-\/]\d{2,4}$/.test(joined) || /^\d{4}[-\/]\d{1,2}[-\/]\d{1,2}$/.test(joined)) continue;
      if (/^\d{5,6}$/.test(joined)) continue;
      
      extraLines.push(joined);
    }

    return { voucherNo, voucherDate, partyLabel, partyValue, extraLines };
  }

  // ── Detect which column index holds amounts ──────────────────────────
  private amountColIdx(headerRow: ClassifiedRow | undefined, bodyRows: ClassifiedRow[]): number {
    // First try header-based detection
    if (headerRow) {
      const idx = headerRow.cells.findIndex((c) => /amount|debit|credit|dr\b|cr\b|total|value|price|cost/i.test(c));
      if (idx >= 0) return idx;
    }
    
    // Pattern-based fallback
    let bestIdx = -1, bestScore = 0;
    const maxCols = Math.max(...bodyRows.slice(0, 25).map((r) => r.cells.length), 0);
    for (let c = 0; c < maxCols; c++) {
      // Count numeric columns - match patterns like: 1000, 1,000, 1000.00, 1,000.00, -500
      const score = bodyRows.slice(0, 25).filter((r) => {
        const cell = (r.cells[c] || "").trim();
        // Numeric pattern: optional minus, digits with optional commas, optional decimal part
        return /^-?\d{1,3}(,\d{3})*(\.\d{1,2})?$/.test(cell) || /^-?\d+(\.\d{1,2})?$/.test(cell);
      }).length;
      if (score > bestScore) { bestScore = score; bestIdx = c; }
    }
    return bestIdx;
  }

  // ── Format amount with proper currency/decimal handling ──────────────
  private formatAmount(value: string): string {
    if (!value || !value.trim()) return "";
    
    const trimmed = value.trim();
    // Remove currency symbols but keep the number
    let num = trimmed.replace(/[^\d,.\-]/g, "").trim();
    
    // Handle different decimal separators
    // If comma comes after 3 digits from right, it's thousands separator
    if (/,\d{3}(\.\d+)?$/.test(num)) {
      // Already in US format: 1,000.00
      return num;
    }
    
    // If period comes after 3 digits from right, it might be thousands separator (European)
    if (/\.\d{3}(,\d+)?$/.test(num)) {
      // European format: 1.000,00 -> convert to US: 1000.00
      num = num.replace(/\./g, "").replace(/,/, ".");
      return num;
    }
    
    // If there's a comma with 2 decimal places, it's decimal separator (European)
    if (/,\d{2}$/.test(num)) {
      num = num.replace(/\./g, "").replace(",", ".");
      return num;
    }
    
    return num;
  }

  // ── Clean and normalize amounts for display ───────────────────────────
  private normalizeAmount(value: string): string {
    const formatted = this.formatAmount(value);
    if (!formatted) return "";
    
    // Parse to number and back for consistency
    const parsed = parseFloat(formatted.replace(/,/g, ""));
    if (isNaN(parsed)) return formatted;
    
    // Format with 2 decimals and thousands separator
    return parsed.toLocaleString("en-IN", { 
      minimumFractionDigits: 2, 
      maximumFractionDigits: 2 
    });
  }

  // ── Validate amounts using pattern analysis ─────────────────────────
  private isValidAmount(value: string): boolean {
    if (!value || !value.trim()) return true; // Empty is valid (optional field)
    
    const formatted = this.formatAmount(value);
    if (!formatted) return false;
    
    // Check if it parses to a valid number
    const parsed = parseFloat(formatted.replace(/,/g, ""));
    return !isNaN(parsed) && isFinite(parsed) && parsed >= 0;
  }

  // ── Smart amount extraction from mixed content ──────────────────────
  private extractAmountValue(cellContent: string): string {
    if (!cellContent || !cellContent.trim()) return "";
    
    const content = cellContent.trim();
    
    // Match patterns like: 1000, 1,000, 1000.00, ₹1000, $1000, etc.
    const amountMatch = content.match(/(-?\d{1,3}(?:,\d{3})*(?:\.\d{1,2})?|\d+\.?\d*)/);
    if (amountMatch) {
      return amountMatch[0];
    }
    
    // If no match, return original (might be formatted differently)
    return content;
  }

  // ════════════════════════════════════════════════════════════════════
  // PUBLIC RENDER — called once per sheet/page
  // ════════════════════════════════════════════════════════════════════

  public render(sheetName: string, rows: string[][], merges: any[]) {
    const classified = classifyRows(rows, merges);
    const thIdx      = classified.findIndex((r) => r.section === "TABLE_HEADER");
    const thRowIdx   = thIdx >= 0 ? classified[thIdx].rowIndex : Infinity;

    const headerRows  = classified.filter((r) => r.section === "HEADER");
    const metaBefore  = classified.filter((r) => r.section === "META" && r.rowIndex < thRowIdx);
    const tableHeader = thIdx >= 0 ? classified[thIdx] : undefined;
    const bodyRows    = classified.filter((r) => r.section === "TABLE_BODY");
    const summaryRows = classified.filter((r) => r.section === "SUMMARY");
    const footerRows  = classified.filter((r) => r.section === "FOOTER");

    const { companyName, supportLines, documentTitle } = this.parseHeader(headerRows, sheetName);

    this.renderPageDecorations(companyName, supportLines, documentTitle);
    this.renderMeta(metaBefore);
    this.renderTable(tableHeader, bodyRows, summaryRows, companyName, supportLines, documentTitle);
    this.renderAmountWords(summaryRows);
    this.renderFooter(footerRows);
  }

  // ── 2. VOUCHER META (No., Date, Paid to, Details) ────────────────────
  private renderMeta(rows: ClassifiedRow[]) {
    const { voucherNo, voucherDate, partyLabel, partyValue, extraLines } = this.parseVoucherMeta(rows);

    // No.: XXXXX  (left)          Date: DD-MM-YYYY (right) — same line
    this.txt(`No.:  ${voucherNo || "-"}`, X_P, this.y, 240, FB, SZ_ME, "left");
    this.txt(`Date:  ${voucherDate || "-"}`, X_P + 240, this.y, TEXT_W - 240, FB, SZ_ME, "right");
    this.y += 18;

    // Paid to:- party name (with multi-line support)
    const partyFull = `${partyLabel} ${partyValue || ""}`;
    const partyH = this.textH(partyFull, TEXT_W, FR, SZ_ME);
    this.txt(partyFull, X_P, this.y, TEXT_W, FR, SZ_ME, "left", true);
    this.y += partyH + 4;

    // Extra detail lines (Through:, Mode of Receipt:, Instrument Details:, etc.)
    if (extraLines.length > 0) {
      this.y += 3;
      // "DETAILS" small caps label
      this.txt("DETAILS", X_P, this.y, TEXT_W, FB, 7);
      this.y += 10;
      for (const line of extraLines) {
        const lineH = this.textH(line, TEXT_W, FR, SZ_ME);
        this.txt(line, X_P, this.y, TEXT_W, FR, SZ_ME, "left", true);
        this.y += lineH + 3;
      }
    }

    this.y += 4;
  }

  // ── 3. TABLE ─────────────────────────────────────────────────────────
  //
  //  ┌───────────────────────────┬────────────┬──────────────┐  ← full borders, bold header
  //  │        Particulars        │   Amount   │   Remarks    │
  //  ├───────────────────────────┼────────────┼──────────────┤
  //  │ On Account of :           │            │              │  ← sub-label (col dividers only)
  //  │ DESCRIPTION TEXT          │    3000.00 │              │  ← data row (col dividers + sides)
  //  │   ...empty rows...        │            │              │  ← padding to MIN height
  //  ├───────────────────────────┼────────────┼──────────────┤  ← bold top border
  //  │           Total           │    3000.00 │              │  ← full borders, bold
  //  └───────────────────────────┴────────────┴──────────────┘
  //
  private renderTable(
    headerRow: ClassifiedRow | undefined,
    bodyRows: ClassifiedRow[],
    summaryRows: ClassifiedRow[],
    companyName: string,
    supportLines: string[],
    documentTitle: string
  ) {
    // ── Column header row ──────────────────────────────────────────────
    // Extract headers from Excel, with smart fallbacks
    const extractHeaders = (header: ClassifiedRow | undefined): string[] => {
      if (!header) return ["Particulars", "Amount", "Remarks"];
      
      const cells = header.cells.map((c) => c.trim()).filter(Boolean);
      if (cells.length >= 3) {
        // Use first 3 non-empty cells
        return cells.slice(0, 3);
      }
      
      if (cells.length === 2) {
        // If only 2 headers, use them + default for 3rd
        return [...cells, "Remarks"];
      }
      
      if (cells.length === 1) {
        // If only 1 header, could be just "Particulars"
        return [cells[0], "Amount", "Remarks"];
      }
      
      // Fallback to defaults if no headers found
      return ["Particulars", "Amount", "Remarks"];
    };

    const labels = extractHeaders(headerRow);

    this.box(X_P, this.y, COL_P, HEADER_ROW_H, LW_OUTER);
    this.box(X_A, this.y, COL_A, HEADER_ROW_H, LW_OUTER);
    this.box(X_R, this.y, COL_R, HEADER_ROW_H, LW_OUTER);

    this.txt(labels[0], X_P + 2, this.y + 5, COL_P - 4, FB, SZ_TH, "center");
    this.txt(labels[1], X_A + 2, this.y + 5, COL_A - 4, FB, SZ_TH, "center");
    this.txt(labels[2], X_R + 2, this.y + 5, COL_R - 4, FB, SZ_TH, "center");
    this.y += HEADER_ROW_H;

    // ── "On Account of :" sub-label ────────────────────────────────────
    // Top border + left/right edges + column dividers
    this.hr(this.y, LW_TABLE);  // top border line
    this.vl(X_P, this.y, this.y + SUB_LABEL_H);
    this.vl(X_A, this.y, this.y + SUB_LABEL_H);
    this.vl(X_R, this.y, this.y + SUB_LABEL_H);
    this.vl(X_END, this.y, this.y + SUB_LABEL_H);
    this.txt("On Account of :", X_P + 4, this.y + 3, COL_P - 8, FB, 8);
    this.y += SUB_LABEL_H;
    this.hr(this.y, LW_TABLE);  // bottom border line
    let bodyStartY = this.y;

    // ── Body rows ──────────────────────────────────────────────────────
    // ── Body rows ──────────────────────────────────────────────────────
    const amtIdx = this.amountColIdx(headerRow, bodyRows);

    for (const row of bodyRows.filter((r) => !isRowEmpty(r.cells))) {
      const cells = row.cells.map((c) => c.trim());
      const amountRaw = amtIdx >= 0 ? (cells[amtIdx] || "") : "";
      const amountExtracted = this.extractAmountValue(amountRaw);
      const amountTxt = this.normalizeAmount(amountExtracted);
      const descTxt   = cells.filter((_, i) => i !== amtIdx).join("  ").trim();

      const descH = this.textH(descTxt, COL_P - 8);
      const rowH  = Math.max(16, descH + 6);

      // Pagination Check (Vouchers usually single-page, but handle for safety)
      if (this.y + rowH > PH - OUTER_M - 40) {
        // Pad the current page table before closing
        const prevBodyUsed = this.y - bodyStartY;
        if (prevBodyUsed < TABLE_BODY_MIN_H) {
          const padH = TABLE_BODY_MIN_H - prevBodyUsed;
          this.vl(X_P,   this.y, this.y + padH);
          this.vl(X_A,   this.y, this.y + padH);
          this.vl(X_R,   this.y, this.y + padH);
          this.vl(X_END, this.y, this.y + padH);
          this.y += padH;
        }
        this.hr(this.y, LW_TABLE); // Close current page table with line
        
        this.doc.addPage({ size: [PW, PH], margin: OUTER_M });
        this.renderPageDecorations(companyName, supportLines, documentTitle);
        
        // Repeat column headers on new page
        this.box(X_P, this.y, COL_P, HEADER_ROW_H, LW_OUTER);
        this.box(X_A, this.y, COL_A, HEADER_ROW_H, LW_OUTER);
        this.box(X_R, this.y, COL_R, HEADER_ROW_H, LW_OUTER);
        this.txt(labels[0], X_P + 2, this.y + 5, COL_P - 4, FB, SZ_TH, "center");
        this.txt(labels[1], X_A + 2, this.y + 5, COL_A - 4, FB, SZ_TH, "center");
        this.txt(labels[2], X_R + 2, this.y + 5, COL_R - 4, FB, SZ_TH, "center");
        this.y += HEADER_ROW_H;
        
        // Repeat sub-label row to match first page identically
        this.hr(this.y, LW_TABLE);
        this.vl(X_P, this.y, this.y + SUB_LABEL_H);
        this.vl(X_A, this.y, this.y + SUB_LABEL_H);
        this.vl(X_R, this.y, this.y + SUB_LABEL_H);
        this.vl(X_END, this.y, this.y + SUB_LABEL_H);
        this.txt("Continued :", X_P + 4, this.y + 3, COL_P - 8, FB, 8);
        this.y += SUB_LABEL_H;
        this.hr(this.y, LW_TABLE);
        
        bodyStartY = this.y;
      }

      // Sidebar borders + column dividers
      this.vl(X_P,   this.y, this.y + rowH);
      this.vl(X_A,   this.y, this.y + rowH);
      this.vl(X_R,   this.y, this.y + rowH);
      this.vl(X_END, this.y, this.y + rowH);

      if (descTxt)   this.txt(descTxt,   X_P + 4, this.y + 3, COL_P - 8, FR, SZ_TB, "left", true);
      if (amountTxt) this.txt(amountTxt, X_A + 2, this.y + 3, COL_A - 6, FR, SZ_TB, "right");

      this.y += rowH;
    }

    // ── Pad body to minimum height (creates the large empty area) ──────
    const bodyUsed = this.y - bodyStartY;
    if (bodyUsed < TABLE_BODY_MIN_H) {
      const padH = TABLE_BODY_MIN_H - bodyUsed;
      this.vl(X_P,   this.y, this.y + padH);
      this.vl(X_A,   this.y, this.y + padH);
      this.vl(X_R,   this.y, this.y + padH);
      this.vl(X_END, this.y, this.y + padH);
      this.y += padH;
    }

    // ── Total row — full borders, bold top line, centered "Total" ──────
    const totalAmtRaw = (() => {
      for (const r of summaryRows) {
        if (isTotalRow(r.cells)) return [...r.cells].reverse().find((c) => c.trim())?.trim() || "";
      }
      return "";
    })();
    const totalAmtExtracted = this.extractAmountValue(totalAmtRaw);  // Extract core amount
    const totalAmt = this.normalizeAmount(totalAmtExtracted);       // Format intelligently

    this.hr(this.y, LW_OUTER);  // bold top border before total row
    this.box(X_P, this.y, COL_P, TOTAL_ROW_H, LW_TABLE);
    this.box(X_A, this.y, COL_A, TOTAL_ROW_H, LW_TABLE);
    this.box(X_R, this.y, COL_R, TOTAL_ROW_H, LW_TABLE);

    this.txt("Total",   X_P + 2, this.y + 4, COL_P - 4, FB, SZ_TO, "center");
    if (totalAmt) this.txt(totalAmt, X_A + 2, this.y + 4, COL_A - 6, FB, SZ_TO, "right");
    this.y += TOTAL_ROW_H;
  }

  // ── 4. AMOUNT IN WORDS ───────────────────────────────────────────────
  // Rendered in a thin-bordered full-width block (matches original exactly)
  private renderAmountWords(summaryRows: ClassifiedRow[]) {
    const lines: Array<{ bold: boolean; text: string }> = [];
    for (const row of summaryRows) {
      if (isTotalRow(row.cells)) continue;
      const text = row.cells.filter((v) => v.trim()).join("  ").trim();
      if (!text) continue;
      lines.push({ bold: isAmountWordsRow(row.cells), text });
    }
    if (lines.length === 0) return;

    // Measure total block height
    const innerPad = 6;
    const blockH = lines.reduce((h, { bold, text }) =>
      h + this.textH(text, TEXT_W - 8, bold ? FB : FR, SZ_WD) + 4, innerPad);

    this.box(X_P, this.y, TEXT_W, blockH);
    let ty = this.y + innerPad / 2;
    for (const { bold, text } of lines) {
      this.txt(text, X_P + 4, ty, TEXT_W - 8, bold ? FB : FR, SZ_WD, "left", true);
      ty += this.textH(text, TEXT_W - 8, bold ? FB : FR, SZ_WD) + 4;
    }
    this.y += blockH;
  }

  // ── 5. FOOTER ────────────────────────────────────────────────────────
  //
  //  ┌────────────────────────────────────────────────────────┐
  //  │ Receiver's contact &                  Authorised      │
  //  │ Signature:                            Signatory       │
  //  │                                                        │
  //  ├─────────────────────┬──────────────────┬──────────────┤
  //  │ Prepared by:        │  Checked by:     │ Verified by: │
  //  └─────────────────────┴──────────────────┴──────────────┘
  //
  private renderFooter(footerRows: ClassifiedRow[]) {
    // ── Receiver & Authorised Signatory bordered block ─────────────────
    this.box(X_P, this.y, TEXT_W, SIG_ROW_H);
    this.txt("Receiver's contact &\nSignature:", X_P + 4, this.y + 6, COL_P - 8, FR, SZ_FT, "left", true);
    this.txt("Authorised Signatory", X_P + 4, this.y + 10, TEXT_W - 8, FR, SZ_FT, "right");
    this.y += SIG_ROW_H;

    // ── Three-column footer row: Prepared by / Checked by / Verified by ─
    let labels = ["Prepared by:", "Checked by:", "Verified by:"];
    for (const row of footerRows) {
      if (isSignatureRow(row.cells)) {
        const parts = row.cells.map((c) => c.trim()).filter(Boolean);
        if (parts.length >= 2) {
          labels = parts.slice(0, 3);
          while (labels.length < 3) labels.push("");
        }
        break;
      }
    }

    const col3W = TEXT_W / 3;
    this.box(X_P,              this.y, col3W, FOOTER_ROW_H);
    this.box(X_P + col3W,      this.y, col3W, FOOTER_ROW_H);
    this.box(X_P + col3W * 2,  this.y, col3W, FOOTER_ROW_H);

    for (let i = 0; i < 3; i++) {
      this.txt(labels[i] || "", X_P + i * col3W + 4, this.y + 6, col3W - 8, FR, SZ_FT);
    }
    this.y += FOOTER_ROW_H;
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// REPORT RENDERER
// Handles generic data grids with dynamic column widths
// ═══════════════════════════════════════════════════════════════════════════

class ReportRenderer extends BaseRenderer implements Renderer {
  constructor(doc: PDFKitDocument) {
    super(doc);
  }

  public render(sheetName: string, rows: string[][], _merges: any[]) {
    if (rows.length === 0) return;

    // Extract headers for branding
    const companyName = sheetName; // Reports use sheetName as title
    const supportLines: string[] = [];
    const documentTitle = "Data Report";

    this.renderPageDecorations(companyName, supportLines, documentTitle);

    // ── 2. COLUMN WIDTHS ────────────────────────────────────────────────
    const colCount = getColumnCount(rows);
    const colW = TEXT_W / colCount;

    // ── 3. TABLE ────────────────────────────────────────────────────────
    const [header, ...body] = rows;

    const renderTableHeaders = (yPos: number) => {
      this.doc.save().rect(X_P, yPos, TEXT_W, HEADER_ROW_H).fill(C_BG);
      this.box(X_P, yPos, TEXT_W, HEADER_ROW_H, LW_OUTER);
      header.forEach((val, i) => {
        const x = X_P + i * colW;
        this.vl(x, yPos, yPos + HEADER_ROW_H);
        this.txt(val, x + 2, yPos + 5, colW - 4, FB, 10, "center");
      });
      this.vl(X_END, yPos, yPos + HEADER_ROW_H);
      return yPos + HEADER_ROW_H;
    };

    // Initial Headers
    this.y = renderTableHeaders(this.y);

    // Render Body
    body.forEach((row, rowIndex) => {
      const maxH = row.reduce((max, val) => Math.max(max, this.textH(val, colW - 4, FR, 8)), 14) + 6;

      // Pagination check (keep safe margin)
      if (this.y + maxH > PH - OUTER_M - 20) {
        this.hr(this.y, 0.1); // Close table on current page
        this.doc.addPage({ size: [PW, PH], margin: OUTER_M });
        this.y = TEXT_X;
        this.renderPageDecorations(companyName, supportLines, documentTitle);
        // Re-render headers on new page
        this.y = renderTableHeaders(this.y);
      }
      
      // Zebra stripe
      if (rowIndex % 2 === 1) {
        this.doc.save().rect(X_P, this.y, TEXT_W, maxH).fill(C_BG);
      }

      row.forEach((val, i) => {
        const x = X_P + i * colW;
        this.vl(x, this.y, this.y + maxH);
        this.txt(val, x + 2, this.y + 3, colW - 4, FR, 8, "left", true);
      });
      this.vl(X_END, this.y, this.y + maxH);
      this.hr(this.y + maxH, 0.1); // thin grid line
      this.y += maxH;
    });

    this.hr(this.y, LW_OUTER);
  }
}

// ═══════════════════════════════════════════════════════════════════════════
// SHEET RENDERING ENTRY POINT
// ═══════════════════════════════════════════════════════════════════════════

const renderSheet = (
  doc: PDFKitDocument,
  _fileName: string,
  sheetName: string,
  sheet: xlsx.WorkSheet,
  analysis: FileAnalysis,
  summaries: SheetSummary[]
) => {
  const isVoucher =
    (analysis.sheetLayoutMeta[sheetName] ??
     ((summaries.find((s) => s.name === sheetName)?.merges ?? 0) > 10 ? "VOUCHER" : "REPORT")) === "VOUCHER";

  let rows = getSheetRows(sheet);
  if (!isVoucher) {
    const { rows: sorted } = sortRowsByDate(rows);
    rows = sorted;
  }

  // US Letter portrait
  doc.addPage({ size: [PW, PH], margin: OUTER_M, layout: "portrait" });

  const renderer: Renderer = isVoucher ? new VoucherRenderer(doc) : new ReportRenderer(doc);
  renderer.render(sheetName, rows, sheet["!merges"] || []);
};

// ═══════════════════════════════════════════════════════════════════════════
// API HANDLER
// ═══════════════════════════════════════════════════════════════════════════

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method !== "POST") {
    res.setHeader("Allow", ["POST"]);
    return res.status(405).send("Method not allowed");
  }

  let filesToCleanup: File[] = [];
  try {
    const { fields, files } = await parseMultipart(req);
    filesToCleanup = files;
    const requestedName = Array.isArray(fields.pdfName) ? fields.pdfName[0] : fields.pdfName;
    const sanitizedName = typeof requestedName === "string"
      ? requestedName.trim().replace(/[\\/?%*:|"<>]/g, "-").trim() : "";
    const outputFileName = `${sanitizedName || "combined-excel"}.pdf`;

    if (!files || files.length === 0) return res.status(400).send("No Excel files were uploaded.");

    const responseMeta: Array<{ fileName: string; analysis: FileAnalysis }> = [];

    const pdfBuffer = await (async () => {
      const chunks: Buffer[] = [];
      const doc = new PDFDocument({ autoFirstPage: false });
      const bufferPromise = new Promise<Buffer>((resolve, reject) => {
        doc.on("data", (c) => chunks.push(c));
        doc.on("end",  () => resolve(Buffer.concat(chunks)));
        doc.on("error", reject);
      });

      for (const file of files) {
        const workbook = xlsx.read(fs.readFileSync(file.filepath), { type: "buffer" });
        const { summaries, analysis } = await analyzeWorkbook(file.originalFilename ?? "unknown", workbook);
        responseMeta.push({ fileName: file.originalFilename ?? "unknown", analysis });

        for (const sheetName of getRenderableSheetNames(workbook, summaries, analysis)) {
          try {
            renderSheet(doc, file.originalFilename ?? "unknown", sheetName, workbook.Sheets[sheetName], analysis, summaries);
          } catch (err) {
            throw new Error(
              `Failed rendering "${file.originalFilename}" / "${sheetName}": ${err instanceof Error ? err.message : err}`
            );
          }
        }
      }

      doc.end();
      return bufferPromise;
    })();

    res.setHeader("Content-Type", "application/pdf");
    res.setHeader("Content-Disposition", `attachment; filename="${outputFileName}"`);
    res.setHeader(
      "X-Workbook-Analysis",
      encodeHeaderValue(JSON.stringify(
        responseMeta.map(({ fileName, analysis }) => ({
          fileName, mode: analysis.mode,
          selectedSheets: analysis.selectedSheets, ignoredSheets: analysis.ignoredSheets,
          reason: analysis.reason, warnings: analysis.warnings,
        }))
      ))
    );
    res.status(200).send(pdfBuffer);
  } catch (error) {
    console.error("API conversion error:", error);
    if (!res.writableEnded)
      res.status(500).send(`Failed to convert Excel files to PDF. ${error instanceof Error ? error.message : ""}`);
  } finally {
    filesToCleanup.forEach((file) => {
      if (fs.existsSync(file.filepath)) try { fs.unlinkSync(file.filepath); } catch {}
    });
  }
}