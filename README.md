# Excel to PDF — AI-Assisted Spreadsheet Export

A Next.js application that accepts Excel workbooks and generates formatted PDF documents. It uses SheetJS to inspect workbook data, PDFKit to create PDFs, and an optional Groq integration to help classify and lay out spreadsheet content. A deterministic fallback path is available when the AI service is not configured.

## Features

- Upload one or more `.xls` or `.xlsx` workbooks
- Inspect worksheets, ranges, rows, columns and merged cells
- Select and classify sheets for PDF output
- Generate formatted PDFs with PDFKit
- Optional Groq-assisted layout analysis with a fallback mode
- Light and dark interface themes
- Drag-and-drop upload interface

## Stack

- Next.js 14 (Pages Router)
- React 18 and TypeScript
- Chakra UI
- SheetJS (`xlsx`)
- PDFKit
- Formidable for multipart uploads
- Groq SDK (optional)

## Requirements

- Node.js 20+
- npm
- Groq API key only if you want AI-assisted analysis; fallback processing is implemented for requests without a key

## Run locally

```bash
git clone https://github.com/Nashid-k/excel-to-pdf.git
cd excel-to-pdf
npm install
cp .env.example .env.local
npm run dev
```

Open [http://localhost:3000](http://localhost:3000).

## Environment

| Variable | Purpose |
| --- | --- |
| `GROQ_API_KEY` | Optional API key for AI-assisted workbook analysis |

Do not expose provider keys in client-side code or commit local environment files.

## API

The conversion endpoint is implemented as a Next.js API route at `/api/convert`. It parses multipart form uploads and has a 200 MB per-file parser limit. Keep deployment request-size limits, execution timeouts and storage constraints in mind; the application should be load-tested before handling large or sensitive workbooks in production.

## Scripts

- `npm run dev` — start the development server
- `npm run build` — create the production build
- `npm start` — start the production server

## Security and data handling

Uploaded workbooks may contain personal or financial information. Validate uploads, avoid logging workbook contents, restrict upload size at the hosting layer, and define a retention policy for temporary files. The optional Groq integration sends selected workbook summaries to the configured provider; review that data flow before processing confidential documents.
