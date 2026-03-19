// convert-to-text.mjs — xlsx/pptx/pdf → txt 一括変換
// Usage: node convert-to-text.mjs [--xlsx] [--pptx] [--pdf] [--force]

import fs from 'fs/promises';
import path from 'path';
import { createRequire } from 'module';

const require = createRequire(import.meta.url);
const XLSX = require('xlsx');
const pdfParse = require('pdf-parse');
const AdmZip = require('adm-zip');

// プロジェクトルートを自動検出（スクリプトの親ディレクトリ）
const SCRIPT_DIR = path.dirname(new URL(import.meta.url).pathname.replace(/^\/([A-Z]:)/, '$1'));
const BASE_DIR = path.resolve(SCRIPT_DIR, '..');
const DRIVE_DIR = path.join(BASE_DIR, 'drive');
const NOTION_DIR = path.join(BASE_DIR, 'notion');

// ============================================================
// Utility: Recursive file finder by extension
// ============================================================
async function findFiles(dir, ext) {
  const results = [];
  try {
    const entries = await fs.readdir(dir, { withFileTypes: true });
    for (const entry of entries) {
      const fullPath = path.join(dir, entry.name);
      if (entry.isDirectory()) {
        results.push(...await findFiles(fullPath, ext));
      } else if (entry.name.toLowerCase().endsWith(ext)) {
        results.push(fullPath);
      }
    }
  } catch (err) {
    if (err.code !== 'ENOENT') throw err;
    console.warn(`[WARN] Directory not found: ${dir}`);
  }
  return results;
}

// ============================================================
// Skip check: does .txt already exist?
// ============================================================
function toTxtPath(filePath) {
  const parsed = path.parse(filePath);
  return path.join(parsed.dir, parsed.name + '.txt');
}

async function shouldSkip(filePath, force) {
  if (force) return false;
  try {
    await fs.access(toTxtPath(filePath));
    return true;
  } catch {
    return false;
  }
}

// ============================================================
// XLSX Converter
// ============================================================
function convertXlsx(filePath) {
  const workbook = XLSX.readFile(filePath);
  const lines = [];
  for (const sheetName of workbook.SheetNames) {
    lines.push(`\n===== Sheet: ${sheetName} =====\n`);
    const sheet = workbook.Sheets[sheetName];
    const text = XLSX.utils.sheet_to_csv(sheet, { FS: '\t' });
    lines.push(text);
  }
  return lines.join('\n');
}

// ============================================================
// PPTX Converter (ZIP + XML parsing)
// ============================================================
function convertPptx(filePath) {
  const zip = new AdmZip(filePath);
  const entries = zip.getEntries();

  const slideEntries = entries
    .filter(e => /^ppt\/slides\/slide\d+\.xml$/i.test(e.entryName))
    .sort((a, b) => {
      const numA = parseInt(a.entryName.match(/slide(\d+)/)[1]);
      const numB = parseInt(b.entryName.match(/slide(\d+)/)[1]);
      return numA - numB;
    });

  const lines = [];
  for (const entry of slideEntries) {
    const slideNum = entry.entryName.match(/slide(\d+)/)[1];
    lines.push(`\n===== Slide ${slideNum} =====\n`);

    const xml = entry.getData().toString('utf8');
    const paragraphs = xml.split(/<\/a:p>/);
    const paraTexts = [];
    for (const para of paragraphs) {
      const texts = para.match(/<a:t[^>]*>([^<]*)<\/a:t>/g);
      if (texts) {
        const paraText = texts.map(t => t.replace(/<[^>]+>/g, '')).join('');
        if (paraText.trim()) {
          paraTexts.push(paraText);
        }
      }
    }
    lines.push(paraTexts.join('\n'));
  }
  return lines.join('\n');
}

// ============================================================
// PDF Converter
// ============================================================
async function convertPdf(filePath) {
  const buffer = await fs.readFile(filePath);
  const data = await pdfParse(buffer);
  return data.text;
}

// ============================================================
// Main
// ============================================================
async function main() {
  const args = process.argv.slice(2);
  const xlsxOnly = args.includes('--xlsx');
  const pptxOnly = args.includes('--pptx');
  const pdfOnly = args.includes('--pdf');
  const force = args.includes('--force');
  const all = !xlsxOnly && !pptxOnly && !pdfOnly;

  console.log(`Base directory: ${BASE_DIR}`);

  const results = { success: 0, skipped: 0, failed: 0, errors: [] };

  async function processFile(filePath, converter) {
    if (await shouldSkip(filePath, force)) {
      console.log(`[SKIP] ${path.relative(BASE_DIR, filePath)} (txt exists)`);
      results.skipped++;
      return;
    }
    try {
      console.log(`[CONVERTING] ${path.relative(BASE_DIR, filePath)}`);
      const text = await converter(filePath);
      const outPath = toTxtPath(filePath);
      await fs.writeFile(outPath, text, 'utf-8');
      console.log(`[OK] -> ${path.relative(BASE_DIR, outPath)}`);
      results.success++;
    } catch (err) {
      console.error(`[FAIL] ${path.relative(BASE_DIR, filePath)}: ${err.message}`);
      results.failed++;
      results.errors.push({ file: path.relative(BASE_DIR, filePath), error: err.message });
    }
  }

  if (all || xlsxOnly) {
    console.log('\n--- XLSX ---');
    const files = await findFiles(DRIVE_DIR, '.xlsx');
    for (const f of files) await processFile(f, convertXlsx);
  }

  if (all || pptxOnly) {
    console.log('\n--- PPTX ---');
    const files = await findFiles(DRIVE_DIR, '.pptx');
    for (const f of files) await processFile(f, convertPptx);
  }

  if (all || pdfOnly) {
    console.log('\n--- PDF ---');
    const files = await findFiles(NOTION_DIR, '.pdf');
    for (const f of files) await processFile(f, convertPdf);
  }

  console.log('\n========== SUMMARY ==========');
  console.log(`Success: ${results.success}`);
  console.log(`Skipped: ${results.skipped}`);
  console.log(`Failed:  ${results.failed}`);
  if (results.errors.length > 0) {
    console.log('\nFailed files:');
    results.errors.forEach(e => console.log(`  - ${e.file}: ${e.error}`));
  }
}

main().catch(console.error);
