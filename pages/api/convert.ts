import type { NextApiRequest, NextApiResponse } from "next";
import fs from "fs";
import PDFDocument from "pdfkit";
import * as xlsx from "xlsx";
import type { File, Fields } from "formidable";
import { formidable } from "formidable";
import { Groq } from "groq-sdk";

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
  hasVernacular: boolean;
};

type FileAnalysis = {
  mode: "groq" | "fallback";
  selectedSheets: string[];
  ignoredSheets: string[];
  sheetLayoutMeta: Record<string, "VOUCHER" | "REPORT">;
  sheetDesigns?: Record<string, {
    headers: [string, string, string];
    columnRatios: [number, number, number];
    typography: { baseFontSize: number; headingFontSize: number; titleFontSize: number; lineSpacing: number };
  }>;
  reason: string;
  warnings: string[];
  verifiedAmounts?: Record<string, { total: number; words: string }>;
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
  return new Promise((resolve, reject) => {
    try {
      const form = formidable({ multiples: true, keepExtensions: true, maxFileSize: 200 * 1024 * 1024 });
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

let _groq: Groq | null = null;
const getGroqClient = async () => {
  if (!process.env.GROQ_API_KEY) return null;
  if (!_groq) {
    try {
      _groq = new Groq({ apiKey: process.env.GROQ_API_KEY! });
    } catch (e) {
      console.error("Failed to init Groq:", e);
      return null;
    }
  }
  return _groq;
};

const getCellValue = (sheet: xlsx.WorkSheet, row: number, col: number): string => {
  const cell = sheet[xlsx.utils.encode_cell({ r: row, c: col })];
  if (!cell) return "";
  if (cell.t === "d" || (cell.t === "n" && cell.z && /[mdy]/i.test(cell.z))) {
    const d = xlsx.utils.format_cell(cell);
    if (d) return d;
  }
  if (cell.w != null) return String(cell.w).trim();
  return String(cell.v ?? "").trim();
};

const normalizeCellValue = (value: unknown) => String(value ?? "").replace(/\s+/g, " ").trim();

const hasNonLatin = (s: string) => /[^\u0000-\u007F\u00A0-\u00FF]/.test(s);

const tryParseDate = (value: string): Date | null => {
  const s = String(value || "").trim();
  if (!s) return null;
  const dmyMatch = s.match(/^(\d{1,2})[-/.](\d{1,2})[-/.](\d{2,4})$/);
  if (dmyMatch) {
    const d = parseInt(dmyMatch[1], 10), m = parseInt(dmyMatch[2], 10);
    let y = parseInt(dmyMatch[3], 10);
    if (y < 100) y += 2000;
    const date = new Date(y, m - 1, d);
    if (date.getFullYear() === y && date.getMonth() === m - 1 && date.getDate() === d) return date;
  }
  const n = Number(s);
  if (!Number.isNaN(n) && n > 20000 && n < 60000 && /^\d+(\.\d+)?$/.test(s)) {
    const date = new Date(Math.round((n - 25569) * 86400 * 1000));
    if (date.getFullYear() >= 1900 && date.getFullYear() <= 2100) return date;
  }
  const ts = Date.parse(s);
  if (!isNaN(ts)) {
    const date = new Date(ts);
    if (date.getFullYear() >= 1900 && date.getFullYear() <= 2100) return date;
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

const numberToWords = (num: number): string => {
  const n = Math.abs(num);
  if (n === 0) return "Zero Only";
  const single = ["", "One", "Two", "Three", "Four", "Five", "Six", "Seven", "Eight", "Nine"];
  const double = ["Ten", "Eleven", "Twelve", "Thirteen", "Fourteen", "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"];
  const tens = ["", "", "Twenty", "Thirty", "Forty", "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"];
  const formatT = (v: number) => v < 10 ? single[v] : v < 20 ? double[v - 10] : tens[Math.floor(v / 10)] + (v % 10 !== 0 ? " " + single[v % 10] : "");
  const formatHundred = (v: number) => {
    let res = "";
    if (v >= 100) { res += single[Math.floor(v / 100)] + " Hundred"; v %= 100; if (v > 0) res += " and "; }
    if (v > 0) res += formatT(v);
    return res;
  };
  const integerPart = Math.floor(n);
  const decimalPart = Math.round((n - integerPart) * 100);
  let result = "";
  let temp = integerPart;
  if (temp >= 10000000) { result += formatHundred(Math.floor(temp / 10000000)) + " Crore "; temp %= 10000000; }
  if (temp >= 100000) { result += formatHundred(Math.floor(temp / 100000)) + " Lakh "; temp %= 100000; }
  if (temp >= 1000) { result += formatHundred(Math.floor(temp / 1000)) + " Thousand "; temp %= 1000; }
  if (temp > 0) result += formatHundred(temp);
  result = result.trim() + " Only";
  if (decimalPart > 0) { result = result.replace(" Only", "") + " and " + formatT(decimalPart) + " Paise Only"; }
  return result;
};

const summarizeSheet = (workbook: Workbook, sheetName: string, sheet: Worksheet): SheetSummary => {
  const rows = getSheetRows(sheet);
  const nonEmptyRows = rows.filter((r) => !isRowEmpty(r));
  return {
    name: sheetName, hidden: getWorkbookSheetMeta(workbook, sheetName).hidden,
    range: sheet["!ref"] || "A1:A1",
    rowCount: rows.length, columnCount: getColumnCount(rows),
    nonEmptyRowCount: nonEmptyRows.length,
    sampleRows: nonEmptyRows.slice(0, 8).map((r) => r.slice(0, 8).map(normalizeCellValue)),
    merges: sheet["!merges"]?.length || 0,
    hasVernacular: nonEmptyRows.some(r => r.some(c => hasNonLatin(c))),
  };
};

const fallbackAnalyzeWorkbook = (fileName: string, workbook: Workbook, summaries: SheetSummary[]): FileAnalysis => {
  const visible = summaries.filter((s) => !s.hidden && s.nonEmptyRowCount > 0);
  const reportLike = visible.filter((s) => /report|summary|invoice|statement|payment|print|final/.test(s.name.toLowerCase()));
  const selected = (reportLike.length > 0 ? reportLike : visible).slice(0, 1).map((s) => s.name);
  return {
    mode: "fallback", selectedSheets: selected.length > 0 ? selected : workbook.SheetNames.slice(0, 1),
    ignoredSheets: summaries.map((s) => s.name).filter((n) => !selected.includes(n)),
    sheetLayoutMeta: Object.fromEntries(summaries.map((s) => [s.name, s.merges > s.rowCount * 0.2 ? "VOUCHER" : "REPORT"])),
    sheetDesigns: Object.fromEntries(summaries.map((s) => [s.name, { headers: ["Particulars", "Amount", "Remarks"], columnRatios: [0.6, 0.2, 0.2], typography: { baseFontSize: 9, headingFontSize: 10, titleFontSize: 14, lineSpacing: 1.2 } }])),
    reason: `Local fallback selected the most likely visible sheet(s) for "${fileName}".`,
    warnings: process.env.GROQ_API_KEY ? ["Groq analysis unavailable; local selection used."] : ["Groq not configured."],
  };
};

const extractJsonObject = (value: string) => {
  const start = value.indexOf("{"), end = value.lastIndexOf("}");
  if (start === -1 || end === -1 || end <= start) throw new Error("No JSON object found.");
  return JSON.parse(value.slice(start, end + 1));
};

const analyzeWorkbookWithGroq = async (fileName: string, summaries: SheetSummary[]): Promise<FileAnalysis> => {
  const groq = await getGroqClient();
  if (!groq) throw new Error("Groq client not configured.");
  const prompt = [
    "You are analyzing an Excel workbook before PDF export. Return only JSON.",
    "Pick sheets most likely for final PDF output. Categorize as 'VOUCHER' or 'REPORT'.",
    'Schema: {"selectedSheets":string[],"ignoredSheets":string[],"sheetLayoutMeta":Record<string,"VOUCHER"|"REPORT">,"reason":string,"warnings":string[]}',
    `File: ${fileName}`, `Summaries: ${JSON.stringify(summaries)}`,
  ].join("\n");
  const response = await groq.chat.completions.create({
    model: process.env.GROQ_MODEL || "llama-3.3-70b-versatile", temperature: 0.2,
    messages: [{ role: "system", content: "Return only JSON." }, { role: "user", content: prompt }],
  });
  const content = response.choices[0]?.message?.content;
  if (!content) throw new Error("Groq empty response.");
  const parsed = extractJsonObject(content) as Partial<FileAnalysis>;
  const isStr = (v: unknown): v is string => typeof v === "string";
  const selected = (Array.isArray(parsed.selectedSheets) ? parsed.selectedSheets.filter(isStr) : []);
  const available = new Set(summaries.map((s) => s.name));
  const normSelected = selected.filter((n) => available.has(n));
  if (normSelected.length === 0) throw new Error("No valid sheets selected.");
  return {
    mode: "groq", selectedSheets: normSelected,
    ignoredSheets: (Array.isArray(parsed.ignoredSheets) ? parsed.ignoredSheets.filter(isStr) : []).filter(n => available.has(n) && !normSelected.includes(n)),
    sheetLayoutMeta: typeof parsed.sheetLayoutMeta === "object" ? parsed.sheetLayoutMeta as any : {},
    sheetDesigns: typeof parsed.sheetDesigns === "object" ? parsed.sheetDesigns : {},
    reason: typeof parsed.reason === "string" ? parsed.reason : `Groq selected ${normSelected.join(", ")}.`,
    warnings: Array.isArray(parsed.warnings) ? parsed.warnings.filter(isStr) : [],
  };
};

const analyzeWorkbook = async (fileName: string, workbook: Workbook) => {
  const summaries = workbook.SheetNames.map((n) => summarizeSheet(workbook, n, workbook.Sheets[n]));
  try {
    const analysis = await analyzeWorkbookWithGroq(fileName, summaries);
    return { summaries, analysis };
  }
  catch (e) { return { summaries, analysis: fallbackAnalyzeWorkbook(fileName, workbook, summaries) }; }
};

const encodeHeaderValue = (v: string) => encodeURIComponent(v.length > 7500 ? v.slice(0, 7500) + "..." : v);

const getRenderableSheetNames = (workbook: Workbook, summaries: SheetSummary[], analysis: FileAnalysis) => {
  const selected = new Set(analysis.selectedSheets);
  return workbook.SheetNames
    .map((name) => ({ name, hidden: summaries.find((s) => s.name === name)?.hidden ?? false, date: getSheetDate(workbook.Sheets[name]) }))
    .filter((s) => !s.hidden && (selected.size === 0 || selected.has(s.name)))
    .sort((a, b) => (a.date && b.date ? a.date.getTime() - b.date.getTime() : a.date ? -1 : b.date ? 1 : 0))
    .map((s) => s.name);
};

const isTableHeaderRow = (row: string[]) => /particulars|description|item\b|amount|rate|qty|quantity|credit|debit|sl\.?\s*no|details\s*of\s*payment/i.test(row.join(" "));
const isMetaRow = (row: string[]) => /^(no|date|through|ca\.?\s*no|mode|party|a\/c|bank|ifsc|branch|cheque|utr)[\s]*[:\-]/i.test(row.filter(v => v.trim()).join(" "));
const isTotalRow = (row: string[]) => {
  const joined = row.join(" ").toLowerCase();
  if (/brought\s*forward|carried\s*forward|\bb\/?f\b|\bc\/?f\b/i.test(joined)) return false;
  const hasTotalKey = (joined.includes("total") || joined.includes("grand") || /^\s*net\s*total/i.test(joined) || joined.includes("payable"));
  if (!hasTotalKey) return false;
  return row.some(c => /^-?[\d,]+(?:\.\d+)?$/.test(c.trim()));
};

const isAmountWordsRow = (row: string[]) => {
  const j = row.filter(v => v.trim()).join(" ").toLowerCase();
  const hasWordsKey = (j.includes("amount in words") || j.includes("rupees") || j.includes("sum of"));
  const endsWithOnly = j.trim().endsWith("only");
  return (hasWordsKey || (endsWithOnly && j.length > 20));
};
const isSignatureRow = (row: string[]) => /prepared\s*by|checked\s*by|authorized|signature|receiver|approved/i.test(row.join(" "));

const classifyRows = (rows: string[][], _merges: any[]): ClassifiedRow[] => {
  const classified: ClassifiedRow[] = [];
  let tableHeaderFound = false, tableEnded = false, footerStarted = false;
  for (let r = 0; r < rows.length; r++) {
    const row = rows[r];
    if (isRowEmpty(row)) { 
      if (tableHeaderFound && !tableEnded) classified.push({ section: "TABLE_BODY", rowIndex: r, cells: row });
      else classified.push({ section: "META", rowIndex: r, cells: row });
      continue; 
    }
    if (footerStarted) { classified.push({ section: "FOOTER", rowIndex: r, cells: row }); continue; }
    if (isSignatureRow(row)) { footerStarted = true; tableEnded = true; classified.push({ section: "FOOTER", rowIndex: r, cells: row }); continue; }
    if (!tableHeaderFound) {
      if (isTableHeaderRow(row)) { tableHeaderFound = true; classified.push({ section: "TABLE_HEADER", rowIndex: r, cells: row }); continue; }
      if (r <= 8) classified.push({ section: "HEADER", rowIndex: r, cells: row });
      else classified.push({ section: "META", rowIndex: r, cells: row });
    } else {
      const it = isTotalRow(row), iaw = isAmountWordsRow(row);
      if (it || iaw) { tableEnded = true; classified.push({ section: "SUMMARY", rowIndex: r, cells: row }); continue; }
      classified.push({ section: tableEnded ? "SUMMARY" : "TABLE_BODY", rowIndex: r, cells: row });
    }
  }
  return classified;
};

const PW_P = 612, PH_P = 792, OUTER_M = 36, TEXT_X = 48;
const FR = "Helvetica", FB = "Helvetica-Bold", C_TEXT = "#111827", C_GRAY = "#4B5563", C_BORDER = "#D1D5DB", C_BG = "#F9FAFB";
const LW_OUTER = 0.8, LW_TABLE = 0.4, TABLE_BODY_MIN_H = 300, SIG_ROW_H = 45, FOOTER_ROW_H = 20;

interface Renderer { render(sheetName: string, rows: string[][], merges: any[]): void; }

class BaseRenderer {
  protected doc: PDFKitDocument; protected y = TEXT_X; protected design: any; protected dyn: Record<string, number> = {};
  constructor(doc: PDFKitDocument, design?: any) {
    this.doc = doc;
    this.design = design || { headers: ["Particulars", "Amount", "Remarks"], columnRatios: [0.6, 0.2, 0.2], typography: { baseFontSize: 10, headingFontSize: 11, titleFontSize: 16, lineSpacing: 1.4 } };
    this.updateDyn();
  }
  protected updateDyn() {
    const t = this.design.typography, cr = this.design.columnRatios;
    const pw = this.doc.page.width, tw = pw - TEXT_X * 2;
    this.dyn.PW = pw; this.dyn.PH = this.doc.page.height; this.dyn.TEXT_W = tw;
    this.dyn.COL_P = Math.round(tw * cr[0]); this.dyn.COL_A = Math.round(tw * cr[1]); this.dyn.COL_R = tw - this.dyn.COL_P - this.dyn.COL_A;
    this.dyn.X_P = TEXT_X; this.dyn.X_A = this.dyn.X_P + this.dyn.COL_P; this.dyn.X_R = this.dyn.X_A + this.dyn.COL_A; this.dyn.X_END = this.dyn.X_P + tw;
    this.dyn.SZ_CO = t.titleFontSize + 4; this.dyn.SZ_AD = t.baseFontSize + 1; this.dyn.SZ_TI = t.titleFontSize; this.dyn.SZ_ME = t.baseFontSize + 1;
    this.dyn.SZ_TH = t.headingFontSize; this.dyn.SZ_TB = t.baseFontSize; this.dyn.SZ_TO = t.headingFontSize; this.dyn.SZ_WD = t.baseFontSize;
    this.dyn.HEADER_ROW_H = t.headingFontSize + 12; this.dyn.SUB_LABEL_H = t.baseFontSize + 8; this.dyn.TOTAL_ROW_H = t.headingFontSize + 10;
    this.dyn.SZ_FT = t.baseFontSize - 2;
  }
  protected renderPageDecorations(companyName: string, supportLines: string[], docTitle: string) {
    this.doc.save().lineWidth(LW_OUTER).strokeColor(C_BORDER).rect(OUTER_M, OUTER_M, this.dyn.PW - OUTER_M * 2, this.dyn.PH - OUTER_M * 2).stroke().restore();
    this.y = TEXT_X + 12; this.txt(companyName, this.dyn.X_P, this.y, this.dyn.TEXT_W, FB, this.dyn.SZ_CO, "center"); this.y += 24;
    for (const l of supportLines) { this.txt(l, this.dyn.X_P, this.y, this.dyn.TEXT_W, FR, this.dyn.SZ_AD, "center", false, C_GRAY); this.y += 14; }
    this.y += 5; this.hr(this.y, LW_TABLE); this.y += 4; this.txt(docTitle, this.dyn.X_P, this.y, this.dyn.TEXT_W, FB, this.dyn.SZ_TI, "center"); this.y += 18; this.hr(this.y, LW_TABLE); this.y += 15;
  }
  protected hr(y: number, lw = LW_TABLE) { this.doc.save().lineWidth(lw).strokeColor(C_BORDER).moveTo(this.dyn.X_P, y).lineTo(this.dyn.X_END, y).stroke().restore(); }
  protected vl(x: number, y1: number, y2: number, lw = LW_TABLE) { this.doc.save().lineWidth(lw).strokeColor(C_BORDER).moveTo(x, y1).lineTo(x, y2).stroke().restore(); }
  protected box(x: number, y: number, w: number, h: number, lw = LW_TABLE) { this.doc.save().lineWidth(lw).strokeColor(C_BORDER).rect(x, y, w, h).stroke().restore(); }
  protected txt(s: string, x: number, y: number, w: number, font = FR, size?: number, align: any = "left", lineBreak = false, color = C_TEXT) {
    const sz = size ?? this.dyn.SZ_TB;
    this.doc.font(font).fontSize(sz).fillColor(color).text(s, x, y, { width: w, align, lineBreak, lineGap: 3 });
  }
  protected textH(s: string, w: number, font = FR, size?: number): number {
    const sz = size ?? this.dyn.SZ_TB;
    this.doc.font(font).fontSize(sz); return toFiniteNumber(this.doc.heightOfString(s, { width: w, lineGap: 3 }), sz);
  }
}

class VoucherRenderer extends BaseRenderer implements Renderer {
  private numericTotal = 0;
  constructor(doc: PDFKitDocument, design?: any) { super(doc, design); }
  private parseHeader(rows: ClassifiedRow[], sheetName: string) {
    const texts = rows.map(r => r.cells.filter(c => c.trim()).join("  ").trim()).filter(Boolean);
    const titleIdx = texts.findIndex(t => /voucher|invoice|receipt|payment/i.test(t));
    const documentTitle = titleIdx >= 0 ? texts[titleIdx] : "Cash Payment Voucher";
    const nonTitle = texts.filter((_, i) => i !== titleIdx);
    const companyName = [...nonTitle].sort((a, b) => b.length - a.length)[0] || sheetName;
    return { companyName, supportLines: nonTitle.filter(t => t !== companyName), documentTitle };
  }
  private findMetaVal(rows: ClassifiedRow[], patterns: RegExp[]): string {
    for (const r of rows) { const j = r.cells.filter(c => c.trim()).join("  "); for (const p of patterns) { const m = j.match(p); if (m) return m[1].trim(); } }
    return "";
  }
  private fmtDate(v: string) { const p = tryParseDate(v); if (!p) return v.trim(); return `${String(p.getDate()).padStart(2, "0")}-${String(p.getMonth() + 1).padStart(2, "0")}-${p.getFullYear()}`; }
  private parseVoucherMeta(rows: ClassifiedRow[]) {
    const text = rows.map(r => r.cells.join(" ")).join(" | ");
    let vDate = "", vNo = "";
    const dm = text.match(/\bdate\s*[:\-]?\s*([^\|]+?)(?=\s*(?:through|mode|paid|received|no:|\|)|$)/i);
    if (dm) vDate = this.fmtDate(dm[1].trim().split(/\s+/)[0]);
    const nm = text.match(/(?:no\.?|voucher\s*no\.?)\s*[:\-]?\s*(\d+|\w+)/i);
    if (nm) vNo = nm[1].trim();
    const party = this.findMetaVal(rows, [/^paid\s*to\s*[:\-]?\s*(.+)$/i, /^received\s*from\s*[:\-]?\s*(.+)$/i, /^party\s*name\s*[:\-]\s*(.+)$/i]);
    const label = /received\s*from/i.test(text) ? "Received from:-" : "Paid to:-";
    const extra = rows.map(r => r.cells.map(c => c.trim()).filter(Boolean).join("  ")).filter(j => j && !/^(no\.?|vouch|date|paid|recei|party)/i.test(j));
    return { vNo, vDate, label, party, extra };
  }
  private amountColIdx(headerRow: ClassifiedRow | undefined, bodyRows: ClassifiedRow[], targetTotal?: number): number {
    if (headerRow) { 
      const idx = headerRow.cells.findIndex(c => /amount|debit|dr|cr|credit|value|sum|total|price|fee|charges|payment|net|gross/i.test(c.toLowerCase())); 
      if (idx >= 0) return idx; 
    }
    let bestIdx = -1, minError = 999999;
    const sampleRows = bodyRows.slice(0, 50).filter(r => !isRowEmpty(r.cells));
    const maxCols = Math.max(...sampleRows.map(r => r.cells.length), 0);
    for (let c = 0; c < maxCols; c++) {
      let colSum = 0, numericCount = 0, dateCount = 0;
      for (const r of sampleRows) {
        const cVal = r.cells[c] || "";
        if (tryParseDate(cVal)) { dateCount++; continue; }
        const val = parseFloat(cVal.replace(/[^\d.\-]/g, "")) || 0;
        if (val !== 0) { colSum += val; numericCount++; }
      }
      if (dateCount > numericCount) continue;
      if (numericCount > 0) {
        if (targetTotal !== undefined && targetTotal > 0) {
          const err = Math.abs(colSum - targetTotal);
          if (err < 0.1 && err < minError) { minError = err; bestIdx = c; }
        } else if (numericCount > (bestIdx === -1 ? 0 : 5)) {
          if (bestIdx === -1) bestIdx = c;
        }
      }
    }
    return bestIdx >= 0 ? bestIdx : (maxCols > 2 ? 1 : maxCols - 1);
  }
  private normalizeAmount(value: string): string {
    const s = value.trim();
    if (!s) return "0.00";
    let numStr = s.replace(/[^\d.\-()]/g, "");
    let isNeg = false;
    if (numStr.startsWith("(") && numStr.endsWith(")")) { isNeg = true; numStr = numStr.slice(1, -1); }
    else if (/\b(?:cr|cr\.)\b/i.test(s)) { isNeg = true; }
    const parsed = parseFloat(numStr);
    if (isNaN(parsed)) return s;
    const finalVal = isNeg ? -Math.abs(parsed) : parsed;
    return finalVal.toLocaleString("en-IN", { minimumFractionDigits: finalVal % 1 !== 0 ? 2 : 0, maximumFractionDigits: 2 });
  }
  private extractAmountValue(s: string) {
    if (!s) return "0";
    const matches = s.match(/-?[\d,]+(?:\.\d+)?/g);
    if (!matches) return "0";
    let maxVal = -1, bestMatch = "0";
    for (const m of matches) {
      const v = Math.abs(parseFloat(m.replace(/,/g, "")));
      if (v > maxVal) { maxVal = v; bestMatch = m; }
    }
    return bestMatch;
  }
  public render(sheetName: string, rows: string[][], merges: any[]) {
    const classified = classifyRows(rows, merges);
    const thIdx = classified.findIndex(r => r.section === "TABLE_HEADER");
    const hRows = classified.filter(r => r.section === "HEADER"), mRows = classified.filter(r => r.section === "META" && (thIdx < 0 || r.rowIndex < classified[thIdx].rowIndex));
    const thRow = thIdx >= 0 ? classified[thIdx] : undefined, bRows = classified.filter(r => r.section === "TABLE_BODY"), sRows = classified.filter(r => r.section === "SUMMARY"), fRows = classified.filter(r => r.section === "FOOTER");
    const { companyName, supportLines, documentTitle } = this.parseHeader(hRows, sheetName);
    this.renderPageDecorations(companyName, supportLines, documentTitle);
    this.renderMeta(mRows);
    this.renderTable(thRow, bRows, sRows, companyName, supportLines, documentTitle);
    this.renderAmountWords(sRows);
    this.renderFooter(fRows);
  }
  private renderMeta(rows: ClassifiedRow[]) {
    const { vNo, vDate, label, party, extra } = this.parseVoucherMeta(rows);
    this.txt(`No.:  ${vNo || "-"}`, this.dyn.X_P, this.y, 240, FB, this.dyn.SZ_ME);
    this.txt(`Date:  ${vDate || "-"}`, this.dyn.X_P + 240, this.y, this.dyn.TEXT_W - 240, FB, this.dyn.SZ_ME, "right");
    this.y += 18;
    const pf = `${label} ${party || ""}`; this.txt(pf, this.dyn.X_P, this.y, this.dyn.TEXT_W, FR, this.dyn.SZ_ME, "left", true);
    this.y += this.textH(pf, this.dyn.TEXT_W, FR, this.dyn.SZ_ME) + 4;
    if (extra.length > 0) { this.y += 3; this.txt("DETAILS", this.dyn.X_P, this.y, this.dyn.TEXT_W, FB, 7); this.y += 10; for (const l of extra) { const lh = this.textH(l, this.dyn.TEXT_W, FR, this.dyn.SZ_ME); this.txt(l, this.dyn.X_P, this.y, this.dyn.TEXT_W, FR, this.dyn.SZ_ME, "left", true); this.y += lh + 3; } }
    this.y += 4;
  }
  private renderTable(th: ClassifiedRow | undefined, br: ClassifiedRow[], sr: ClassifiedRow[], cn: string, sl: string[], dt: string) {
    this.renderTableHeader(["Particulars", "Amount", "Remarks"]);
    this.renderSubHeader();
    let excelTotalValue = 0, excelTotalText = "0", totalRow = sr.find(r => isTotalRow(r.cells));
    if (totalRow) {
      const nums = totalRow.cells.map(c => this.extractAmountValue(c).replace(/,/g, "")).map(v => parseFloat(v)).filter(v => !isNaN(v));
      if (nums.length > 0) { excelTotalValue = Math.max(...nums); excelTotalText = excelTotalValue.toLocaleString("en-IN", { minimumFractionDigits: excelTotalValue % 1 !== 0 ? 2 : 0, maximumFractionDigits: 2 }); }
    }
    let bodyStartY = this.y, amtIdx = this.amountColIdx(th, br, excelTotalValue), rowIndex = 1;
    const processedRows: any[] = [];
    for (const r of br.filter(x => !isRowEmpty(x.cells))) {
      const cells = r.cells.map(c => c.trim());
      let descBase = cells.filter((_, i) => i !== amtIdx).join("  ").trim();
      let rawAmt = amtIdx >= 0 ? (cells[amtIdx] || "") : "";
      if (!rawAmt && descBase) { const matches = [...descBase.matchAll(/\b-?(\d{1,3}(?:,\d{3})+|\d{4,}|\d+)(?:\.\d{2,})?\b/g)]; if (matches.length > 0) { rawAmt = matches[matches.length - 1][0]; descBase = descBase.replace(rawAmt, "").trim(); } }
      const val = parseFloat(this.extractAmountValue(rawAmt).replace(/,/g, "")) || 0;
      const isMajor = /details\s*of\s*payment|party(?:'?s?)?\s*bank\s*details|bank\s*details/i.test(descBase);
      const isField = /\b(?:a\/c\.?\s*(?:name|no\.?)|ifsc|utr|bank\s*name|branch|name\s*of\s*the\s*bank)\b/i.test(descBase);
      const desc = (!isMajor && !isField && descBase) ? `${rowIndex}. ${descBase}` : isMajor ? descBase.toUpperCase() : descBase;
      if (!isMajor && !isField && descBase) rowIndex++;
      processedRows.push({ desc, amtTxt: this.normalizeAmount(String(val)), val, isHeader: isMajor, isField, padding: isMajor ? 16 : isField ? 12 : 10 });
    }
    for (const r of processedRows) {
      const rh = Math.max(16, this.textH(r.desc, this.dyn.COL_P - 8) + r.padding);
      if (this.y + rh > this.dyn.PH - OUTER_M - 60) {
        if (this.y - bodyStartY < TABLE_BODY_MIN_H) { const p = TABLE_BODY_MIN_H - (this.y - bodyStartY); this.vl(this.dyn.X_P, this.y, this.y + p); this.vl(this.dyn.X_A, this.y, this.y + p); this.vl(this.dyn.X_R, this.y, this.y + p); this.vl(this.dyn.X_END, this.y, this.y + p); this.y += p; }
        this.hr(this.y); this.doc.addPage({ size: [this.dyn.PW, this.dyn.PH], margin: OUTER_M }); this.updateDyn(); this.renderPageDecorations(cn, sl, dt); this.y += 10; this.renderTableHeader(["Particulars", "Amount", "Remarks"]); this.renderSubHeader(); bodyStartY = this.y;
      }
      this.vl(this.dyn.X_P, this.y, this.y + rh); this.vl(this.dyn.X_A, this.y, this.y + rh); this.vl(this.dyn.X_R, this.y, this.y + rh); this.vl(this.dyn.X_END, this.y, this.y + rh);
      const ind = r.isField ? 16 : 4; this.txt(r.desc, this.dyn.X_P + ind, this.y + (r.padding / 2), this.dyn.COL_P - (ind + 4), r.isHeader ? FB : FR, r.isHeader ? this.dyn.SZ_TH : this.dyn.SZ_TB, "left", true);
      if (r.val !== 0) this.txt(r.amtTxt, this.dyn.X_A + 2, this.y + (r.padding / 2), this.dyn.COL_A - 6, r.isHeader ? FB : FR, this.dyn.SZ_TB, "right");
      this.y += rh;
    }
    const used = this.y - bodyStartY; if (used < TABLE_BODY_MIN_H) { const p = TABLE_BODY_MIN_H - used; this.vl(this.dyn.X_P, this.y, this.y + p); this.vl(this.dyn.X_A, this.y, this.y + p); this.vl(this.dyn.X_R, this.y, this.y + p); this.vl(this.dyn.X_END, this.y, this.y + p); this.y += p; }
    this.numericTotal = excelTotalValue; this.hr(this.y, LW_OUTER); this.box(this.dyn.X_P, this.y, this.dyn.COL_P, this.dyn.TOTAL_ROW_H); this.box(this.dyn.X_A, this.y, this.dyn.COL_A, this.dyn.TOTAL_ROW_H); this.box(this.dyn.X_R, this.y, this.dyn.COL_R, this.dyn.TOTAL_ROW_H);
    this.txt("Total", this.dyn.X_P + 2, this.y + 4, this.dyn.COL_P - 4, FB, this.dyn.SZ_TO, "center"); this.txt(excelTotalText, this.dyn.X_A + 2, this.y + 4, this.dyn.COL_A - 6, FB, this.dyn.SZ_TO, "right"); this.y += this.dyn.TOTAL_ROW_H;
  }
  private renderTableHeader(lbs: string[]) {
    for (let i = 0; i < 3; i++) { const w = i === 0 ? this.dyn.COL_P : i === 1 ? this.dyn.COL_A : this.dyn.COL_R; const x = i === 0 ? this.dyn.X_P : i === 1 ? this.dyn.X_A : this.dyn.X_R; this.box(x, this.y, w, this.dyn.HEADER_ROW_H, LW_OUTER); this.txt(lbs[i], x + 2, this.y + 5, w - 4, FB, this.dyn.SZ_TH, "center"); }
    this.y += this.dyn.HEADER_ROW_H;
  }
  private renderSubHeader() {
    this.hr(this.y); this.vl(this.dyn.X_P, this.y, this.y + this.dyn.SUB_LABEL_H); this.vl(this.dyn.X_A, this.y, this.y + this.dyn.SUB_LABEL_H); this.vl(this.dyn.X_R, this.y, this.y + this.dyn.SUB_LABEL_H); this.vl(this.dyn.X_END, this.y, this.y + this.dyn.SUB_LABEL_H);
    this.txt("On Account of :", this.dyn.X_P + 4, this.y + 3, this.dyn.COL_P - 8, FB, 8); this.y += this.dyn.SUB_LABEL_H; this.hr(this.y);
  }
  private renderAmountWords(sr: ClassifiedRow[]) {
    const wordsRow = sr.find(r => isAmountWordsRow(r.cells));
    let clean = wordsRow ? wordsRow.cells.join(" ").replace(/amount\s*\(?in\s*words\)?\s*[:\-]?\s*/i, "").trim() : numberToWords(this.numericTotal);
    if (!clean.toLowerCase().startsWith("rupees")) clean = clean.charAt(0).toUpperCase() + clean.slice(1);
    const text = "Amount in words: " + clean; const bh = this.textH(text, this.dyn.TEXT_W - 8, FB, this.dyn.SZ_WD) + 8;
    this.box(this.dyn.X_P, this.y, this.dyn.TEXT_W, bh); this.txt(text, this.dyn.X_P + 4, this.y + 4, this.dyn.TEXT_W - 8, FB, this.dyn.SZ_WD, "left", true); this.y += bh;
  }
  private renderFooter(rows: ClassifiedRow[]) {
    this.y = Math.max(this.y, this.dyn.PH - OUTER_M - 110);
    const act = rows.filter(r => r.cells.some(c => c.trim().length > 3));
    if (act.length > 0) { this.box(this.dyn.X_P, this.y, this.dyn.TEXT_W, 10 + act.length * 12); act.forEach((r, i) => this.txt(r.cells.join("  "), this.dyn.X_P + 4, this.y + 4 + i * 12, this.dyn.TEXT_W - 8, FR, this.dyn.SZ_FT)); this.y += 15 + act.length * 12; }
    else {
      this.box(this.dyn.X_P, this.y, this.dyn.TEXT_W, SIG_ROW_H); this.txt("Receiver's Signature:", this.dyn.X_P + 4, this.y + 6, this.dyn.COL_P, FR, this.dyn.SZ_FT); this.txt("Authorised Signatory", this.dyn.X_A, this.y + 10, this.dyn.COL_A + this.dyn.COL_R - 8, FR, this.dyn.SZ_FT, "right"); this.y += SIG_ROW_H;
      const c3 = this.dyn.TEXT_W / 3, lbs = ["Prepared by:", "Checked by:", "Verified by:"];
      for (let i = 0; i < 3; i++) { this.box(this.dyn.X_P + i * c3, this.y, c3, FOOTER_ROW_H); this.txt(lbs[i], this.dyn.X_P + i * c3 + 4, this.y + 6, c3 - 8, FR, this.dyn.SZ_FT); }
    }
  }
}

class ReportRenderer extends BaseRenderer implements Renderer {
  public render(sheetName: string, rows: string[][]) {
    if (rows.length === 0) return;
    this.renderPageDecorations(sheetName, [], "Data Report");
    const count = getColumnCount(rows), cw = this.dyn.TEXT_W / count, [h, ...b] = rows;
    const rh = (y: number) => { this.doc.save().rect(this.dyn.X_P, y, this.dyn.TEXT_W, this.dyn.HEADER_ROW_H).fill(C_BG); this.box(this.dyn.X_P, y, this.dyn.TEXT_W, this.dyn.HEADER_ROW_H, LW_OUTER); h.forEach((v, i) => { const x = this.dyn.X_P + i * cw; this.vl(x, y, y + this.dyn.HEADER_ROW_H); this.txt(v, x + 2, y + 5, cw - 4, FB, 10, "center"); }); this.vl(this.dyn.X_END, y, y + this.dyn.HEADER_ROW_H); return y + this.dyn.HEADER_ROW_H; };
    this.y = rh(this.y);
    b.forEach((r, i) => {
      const mh = r.reduce((m, v) => Math.max(m, this.textH(v, cw - 4, FR, 8)), 14) + 6;
      if (this.y + mh > this.dyn.PH - OUTER_M - 20) { this.hr(this.y, 0.1); this.doc.addPage({ size: [this.dyn.PW, this.dyn.PH], margin: OUTER_M }); this.updateDyn(); this.y = TEXT_X; this.renderPageDecorations(sheetName, [], "Data Report"); this.y = rh(this.y); }
      if (i % 2 === 1) this.doc.save().rect(this.dyn.X_P, this.y, this.dyn.TEXT_W, mh).fill(C_BG);
      r.forEach((v, idx) => { const x = this.dyn.X_P + idx * cw; this.vl(x, this.y, this.y + mh); this.txt(v, x + 2, this.y + 3, cw - 4, FR, 8, "left", true); }); this.vl(this.dyn.X_END, this.y, this.y + mh); this.y += mh;
    });
    this.hr(this.y, LW_OUTER);
  }
}

const renderSheet = (doc: PDFKitDocument, sn: string, s: Worksheet, a: FileAnalysis) => {
  const rows = getSheetRows(s), { rows: sorted } = sortRowsByDate(rows), classified = classifyRows(sorted, s["!merges"] || []);
  if (!classified.some(r => r.section === "TABLE_BODY" && !isRowEmpty(r.cells)) && !classified.some(r => r.section === "META" && !isRowEmpty(r.cells))) return;
  const isV = a.sheetLayoutMeta[sn] === "VOUCHER", colCount = getColumnCount(rows);
  doc.addPage({ size: [PW_P, PH_P], margin: OUTER_M, layout: (!isV && colCount > 7) ? "landscape" : "portrait" });
  const ren: Renderer = isV ? new VoucherRenderer(doc, a.sheetDesigns?.[sn]) : new ReportRenderer(doc, a.sheetDesigns?.[sn]);
  ren.render(sn, sorted, s["!merges"] || []);
};

export default async function handler(req: NextApiRequest, res: NextApiResponse) {
  if (req.method !== "POST") return res.status(405).send("Method not allowed");
  let list: File[] = [];
  try {
    const { fields, files } = await parseMultipart(req); list = files;
    const rName = Array.isArray(fields.pdfName) ? fields.pdfName[0] : fields.pdfName;
    const sName = typeof rName === "string" ? rName.trim().replace(/[\\/?%*:|"<>]/g, "-").trim() : "";
    const out = `${sName || "combined-excel"}.pdf`, meta: any[] = [];
    const buf = await (async () => {
      const chunks: Buffer[] = [], doc = new PDFDocument({ autoFirstPage: false });
      const p = new Promise<Buffer>((resolve, reject) => { doc.on("data", c => chunks.push(c)); doc.on("end", () => resolve(Buffer.concat(chunks))); doc.on("error", reject); });
      for (const f of files) {
        const wb = xlsx.read(fs.readFileSync(f.filepath), { type: "buffer" });
        const { summaries, analysis } = await analyzeWorkbook(f.originalFilename ?? "unknown", wb);
        meta.push({ fileName: f.originalFilename ?? "unknown", analysis });
        for (const sn of getRenderableSheetNames(wb, summaries, analysis)) renderSheet(doc, sn, wb.Sheets[sn], analysis);
      }
      doc.end(); return p;
    })();
    res.setHeader("Content-Type", "application/pdf"); res.setHeader("Content-Disposition", `attachment; filename="${out}"`);
    res.setHeader("X-Workbook-Analysis", encodeHeaderValue(JSON.stringify(meta.map(m => ({ fileName: m.fileName, mode: m.analysis.mode, selectedSheets: m.analysis.selectedSheets, ignoredSheets: m.analysis.ignoredSheets, reason: m.analysis.reason, warnings: m.analysis.warnings })))));
    res.status(200).send(buf);
  } catch (e) { console.error(e); if (!res.writableEnded) res.status(500).send(`Error: ${e instanceof Error ? e.message : e}`); }
  finally { list.forEach(f => { if (fs.existsSync(f.filepath)) try { fs.unlinkSync(f.filepath); } catch { } }); }
}