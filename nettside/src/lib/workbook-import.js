import { normalizeText, parseNumber } from './formatting.js';
import { validateWorkbookFile } from './report-engine.js';
import { parseReportDate } from './date-utils.js';

export const maxUploadBytes = 15 * 1024 * 1024;
export const maxWorkbookRows = 25_000;
export const maxWorkbookColumns = 250;
export const maxBatchFiles = 16;
const maxHeaderRows = 20;
const maxExpandedBytes = 100 * 1024 * 1024;
const labels = { fastlønn: 'Fastlønn', overtid: 'Overtid', vakttillegg: 'Vakttillegg', foreldrepermisjon: 'Foreldrepermisjon' };
const requiredColumns = {
  fastlønn: ['Etternavn', 'Fornavn', 'Stillingsgruppe betegnelse', 'Deltids-% 0007', 'Medarbeidergruppe'],
  overtid: ['Etternavn, fornavn', 'Utbetalingsdato', 'Beløp'],
  vakttillegg: ['Etternavn, fornavn', 'Utbetalingsdato', 'Beløp'],
  foreldrepermisjon: ['Fornavn', 'Etternavn', 'Kjønn', 'Start', 'Slutt']
};
const canonicalHeaders = new Map([
  ...new Set(Object.values(requiredColumns).flat()), 'Nøkkel for kjønn', '107A - Individuell lønn årsbe',
  '1006-Årslønn lederlønnstab.', 'Ansattnummer', 'Ansattnr.', 'Ansattnr - navn', 'Rapportgruppe', 'Personalansvar',
  'Frav.type', 'Frav.tekst', 'Frav.dager', 'Arbeidsførhet'
].map((name) => [normalizeText(name), name]));

export function parseSnapshotMeta(fileName) {
  const text = normalizeText(fileName);
  const dated = text.match(/per(?:\s*den)?\s*(\d{1,2})[.\-/](\d{1,2})[.\-/](\d{4}|\d{2})(?!\d)/);
  const compact = text.match(/per(?:\s*den)?\s*(\d{2})(\d{2})(\d{4}|\d{2})(?!\d)/);
  const iso = text.match(/\b(20\d{2})-(\d{2})-(\d{2})\b/);
  const match = dated || compact;
  const year = match ? (match[3].length === 2 ? `20${match[3]}` : match[3]) : iso?.[1];
  const month = match?.[2] || iso?.[2];
  const day = match?.[1] || iso?.[3];
  const date = year && parseReportDate(`${year}-${month}-${day}`);
  if (!date) return { snapshotKey: 'udatert', snapshotLabel: 'Udatert uttrekk', fileYear: text.match(/\b(20\d{2})\b/)?.[1] || '' };
  return { snapshotKey: date.toISOString().slice(0, 10), snapshotLabel: date.toLocaleDateString('nb-NO', { timeZone: 'UTC' }), fileYear: year };
}

export function detectFileMeta(fileName, header, expectedRole) {
  const lower = normalizeText(fileName);
  const headers = new Set(header.map(normalizeText));
  const has = (name) => headers.has(normalizeText(name));
  const source = lower.includes('bearbeidet') ? 'bearbeidet' : lower.includes('utregningsskjema') ? 'referanse' : 'rådata';
  let role = /fastl[oø]nn/.test(lower) ? 'fastlønn' : lower.includes('overtid') ? 'overtid'
    : lower.includes('vakt') ? 'vakttillegg' : lower.includes('permisjon') ? 'foreldrepermisjon' : '';
  if (!role && has('Stillingsgruppe betegnelse')) role = 'fastlønn';
  if (!role && has('Start') && has('Slutt') && has('Kjønn')) role = 'foreldrepermisjon';
  if (!role && expectedRole) role = expectedRole;
  return { role: role || 'ukjent', source, label: labels[role] || 'Ukjent fil', ...parseSnapshotMeta(fileName) };
}

export function compareSnapshotKeys(left, right) {
  if (left === right) return 0;
  if (!left || left === 'udatert') return 1;
  if (!right || right === 'udatert') return -1;
  return left.localeCompare(right);
}

export function deriveFileYear(role, rows, fallbackYear = '') {
  return fallbackYear || yearsInRows(role, rows)[0] || '';
}

function yearsInRows(role, rows) {
  const years = new Set();
  for (const row of rows) {
    const start = parseReportDate(role === 'foreldrepermisjon' ? row.Start : row.Utbetalingsdato);
    const end = role === 'foreldrepermisjon' ? parseReportDate(row.Slutt) : start;
    if (!start || !end) continue;
    for (let year = start.getUTCFullYear(); year <= Math.min(end.getUTCFullYear(), start.getUTCFullYear() + 100); year++) years.add(String(year));
  }
  return [...years].sort();
}

export function fileCoversYear(file, year) {
  return !year || (file.years || [file.fileYear]).includes(String(year));
}

// Inspect directory sizes before SheetJS expands the ZIP container. Never parse
// disguised HTML/legacy XLS files, macro workbooks or oversized archives.
export function validateWorkbookArchive(buffer) {
  const bytes = new Uint8Array(buffer);
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  if (bytes.length < 22 || view.getUint32(0, true) !== 0x04034b50) throw new Error('Filen er ikke en gyldig XLSX-arbeidsbok.');
  let footer = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65557); i--) {
    if (view.getUint32(i, true) === 0x06054b50 && i + 22 + view.getUint16(i + 20, true) === bytes.length) { footer = i; break; }
  }
  if (footer < 0) throw new Error('XLSX-arkivet er skadet. Lagre filen på nytt fra Excel.');
  const count = view.getUint16(footer + 10, true);
  let offset = view.getUint32(footer + 16, true);
  if (count > 4096 || view.getUint16(footer + 4, true) || view.getUint16(footer + 6, true)) throw new Error('Arbeidsboken har et format eller en størrelse som ikke støttes.');
  let expanded = 0;
  let hasWorkbook = false;
  for (let i = 0; i < count; i++) {
    if (offset + 46 > footer || view.getUint32(offset, true) !== 0x02014b50) throw new Error('XLSX-arkivet er skadet.');
    const nameLength = view.getUint16(offset + 28, true);
    const next = offset + 46 + nameLength + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true);
    if (next > footer) throw new Error('XLSX-arkivet er skadet.');
    if (view.getUint16(offset + 8, true) & 1) throw new Error('Passordbeskyttede arbeidsbøker støttes ikke.');
    const name = new TextDecoder().decode(bytes.subarray(offset + 46, offset + 46 + nameLength));
    if (/vbaProject\.bin$/i.test(name)) throw new Error('Arbeidsbøker med makroer støttes ikke.');
    hasWorkbook ||= name === 'xl/workbook.xml';
    expanded += view.getUint32(offset + 24, true);
    if (expanded > maxExpandedBytes) throw new Error('Arbeidsboken er for stor når den pakkes ut (maks 100 MB).');
    offset = next;
  }
  if (!hasWorkbook) throw new Error('Filen inneholder ingen XLSX-arbeidsbok.');
}

function makeHeaders(row) {
  const seen = new Map();
  return row.map((cell, index) => {
    const raw = String(cell ?? '').trim();
    const name = canonicalHeaders.get(normalizeText(raw)) || raw || `Kolonne ${index + 1}`;
    if (name.startsWith('__') || ['__proto__', 'constructor', 'prototype'].includes(name)) throw new Error('Arbeidsboken inneholder reserverte kolonnenavn.');
    const occurrence = (seen.get(name) || 0) + 1;
    seen.set(name, occurrence);
    if (occurrence > 1 && canonicalHeaders.has(normalizeText(name))) throw new Error(`Kolonnen «${name}» forekommer flere ganger.`);
    return occurrence > 1 ? `${name} (${occurrence})` : name;
  });
}

export function validateWorkbookColumns(role, header) {
  const required = requiredColumns[role];
  if (!required) throw new Error('Filtypen kunne ikke gjenkjennes. Bruk filkortet for riktig rådatatype.');
  const missing = required.filter((column) => !header.includes(column));
  if (role === 'fastlønn') {
    if (!header.includes('Nøkkel for kjønn') && !header.includes('Kjønn')) missing.push('Nøkkel for kjønn');
    if (!header.includes('107A - Individuell lønn årsbe') && !header.includes('1006-Årslønn lederlønnstab.')) missing.push('årslønn');
  }
  if (missing.length) throw new Error(`Mangler påkrevde kolonner: ${missing.join(', ')}.`);
}

function validateRows(role, rows) {
  for (const row of rows) {
    const named = role === 'fastlønn' || role === 'foreldrepermisjon' ? row.Fornavn || row.Etternavn : row['Etternavn, fornavn'];
    if (!named) continue; // SAP subtotal rows are not employees or payments.
    if (role === 'overtid' || role === 'vakttillegg') {
      if (!parseReportDate(row.Utbetalingsdato)) throw new Error(`Rad ${row.__sourceRow} mangler en gyldig utbetalingsdato.`);
      if (parseNumber(row.Beløp) === null) throw new Error(`Rad ${row.__sourceRow} mangler et gyldig beløp.`);
    }
    if (role === 'foreldrepermisjon') {
      const start = parseReportDate(row.Start), end = parseReportDate(row.Slutt);
      if (!row.Fornavn || !row.Etternavn || !row.Kjønn) throw new Error(`Rad ${row.__sourceRow} mangler navn eller kjønn.`);
      if (!start || !end || start > end) throw new Error(`Rad ${row.__sourceRow} har en ugyldig permisjonsperiode.`);
    }
  }
}

export async function parseExcelFiles(files, expectedRole) {
  if (!files.length || files.length > maxBatchFiles) throw new Error(`Velg mellom 1 og ${maxBatchFiles} filer om gangen.`);
  for (const file of files) validateWorkbookFile(file, maxUploadBytes);
  if (files.reduce((sum, file) => sum + file.size, 0) > 60 * 1024 * 1024) throw new Error('Velg færre filer. Samlet grense er 60 MB per opplasting.');
  const XLSX = await import('xlsx');
  const parsed = [];
  for (const file of files) {
    try {
      const buffer = await file.arrayBuffer();
      validateWorkbookArchive(buffer);
      const workbook = XLSX.read(buffer, { type: 'array', dense: true, cellDates: true, sheets: 0, sheetRows: maxWorkbookRows + maxHeaderRows + 1 });
      const sheetName = workbook.SheetNames[0];
      if (!sheetName) throw new Error('Arbeidsboken inneholder ingen ark.');
      const sheet = workbook.Sheets[sheetName];
      const fullRange = XLSX.utils.decode_range(sheet['!fullref'] || sheet['!ref'] || 'A1');
      if (fullRange.e.r + 1 > maxWorkbookRows + maxHeaderRows) throw new Error(`Arbeidsboken overskrider grensen på ${maxWorkbookRows.toLocaleString('nb-NO')} datarader.`);
      if (fullRange.e.c + 1 > maxWorkbookColumns) throw new Error(`Arbeidsboken har mer enn ${maxWorkbookColumns} kolonner.`);
      // Excel dates are calendar dates without a timezone. SheetJS creates local
      // Date objects; normalize their wall-clock date before UTC calculations.
      const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '', blankrows: true, range: 0 })
        .map((row) => row.map((value) => value instanceof Date
          ? new Date(Date.UTC(value.getFullYear(), value.getMonth(), value.getDate())) : value));
      let headerIndex = -1;
      let meta, header;
      for (let i = 0; i < Math.min(matrix.length, maxHeaderRows); i++) {
        const candidate = matrix[i].map((cell) => canonicalHeaders.get(normalizeText(cell)) || String(cell).trim());
        const detection = detectFileMeta(file.name, candidate, expectedRole);
        if ((requiredColumns[detection.role] || []).filter((key) => candidate.includes(key)).length >= 3) {
          headerIndex = i; meta = detection; header = makeHeaders(matrix[i]); break;
        }
      }
      if (headerIndex < 0) throw new Error('Fant ingen gyldig kolonneoverskrift i de første 20 radene. Kontroller at første ark inneholder rådata.');
      if (meta.source !== 'rådata') throw new Error('Velg rådatauttrekket. Bearbeidede filer og utregningsskjema skal ikke brukes som datagrunnlag.');
      if (expectedRole && meta.role !== expectedRole) throw new Error(`Filen tilhører ${meta.label.toLowerCase()}, ikke ${labels[expectedRole].toLowerCase()}.`);
      validateWorkbookColumns(meta.role, header);
      if (meta.role === 'fastlønn' && !['05-01', '12-31'].includes(meta.snapshotKey.slice(5))) throw new Error('Fastlønn må ha eksakt uttrekksdato 01.05 eller 31.12 i filnavnet, for eksempel «Fastlønn per 31.12.2025.xlsx».');
      const rows = matrix.slice(headerIndex + 1).map((row, index) => ({ row, number: headerIndex + index + 2 }))
        .filter(({ row }) => row.some((value) => String(value).trim() !== ''))
        .map(({ row, number }) => ({ ...Object.fromEntries(header.map((key, index) => [key, row[index] ?? ''])), __sourceRow: number }));
      if (rows.length > maxWorkbookRows) throw new Error(`Arbeidsboken har mer enn ${maxWorkbookRows.toLocaleString('nb-NO')} datarader.`);
      validateRows(meta.role, rows);
      const years = meta.role === 'fastlønn' ? [meta.fileYear] : yearsInRows(meta.role, rows);
      if (!years.length && meta.fileYear) years.push(meta.fileYear);
      if (!years.length) throw new Error('Rapportåret kunne ikke bestemmes. Bruk årstall i filnavnet for et tomt uttrekk.');
      const warnings = workbook.SheetNames.length > 1 ? [`Bare første ark («${sheetName}») er brukt. Arbeidsboken har ${workbook.SheetNames.length} ark.`] : [];
      parsed.push({
        fileName: file.name, label: meta.label, role: meta.role, source: meta.source,
        snapshotKey: meta.role === 'fastlønn' ? meta.snapshotKey : '', snapshotLabel: meta.role === 'fastlønn' ? meta.snapshotLabel : '',
        fileYear: deriveFileYear(meta.role, rows, meta.fileYear), years,
        rowCount: rows.length, columnCount: header.length, header,
        preview: rows.slice(0, 5).map((row) => header.map((column) => row[column])), rows, warnings
      });
    } catch (error) {
      throw new Error(`${file.name}: ${error.message || 'Kunne ikke lese arbeidsboken.'}`);
    }
  }
  for (let i = 0; i < parsed.length; i++) {
    if (parsed.slice(i + 1).some((file) => filesOverlap(parsed[i], file))) throw new Error('Flere filer dekker samme datatype og rapportår. Velg ett samlet rådatauttrekk per datatype og år.');
  }
  return parsed;
}

export function fileIdentity(file) {
  return `${file.role}::${file.source}::${file.snapshotKey || (file.years || [file.fileYear]).join(',')}`;
}

function filesOverlap(left, right) {
  if (left.role !== right.role || left.source !== right.source) return false;
  if (left.role === 'fastlønn') return left.snapshotKey === right.snapshotKey;
  return (left.years || [left.fileYear]).some((year) => fileCoversYear(right, year));
}

export function mergeFiles(existingFiles, newFiles) {
  return [...existingFiles.filter((existing) => !newFiles.some((file) => filesOverlap(existing, file))), ...newFiles];
}
