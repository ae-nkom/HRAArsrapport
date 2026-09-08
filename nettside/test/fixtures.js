// Entirely fictional records. Never put local HR source records in this file.
import * as XLSX from 'xlsx';

export const salaryRow = (overrides = {}) => ({
  Fornavn: 'Kari', Etternavn: 'Eksempel', 'Nøkkel for kjønn': 'Kvinne',
  'Stillingsgruppe betegnelse': '1364 Seniorrådgiver',
  '107A - Individuell lønn årsbe': 800000, 'Deltids-% 0007': 100,
  Medarbeidergruppe: 'Fast ansatte', ...overrides
});

export const payRow = (overrides = {}) => ({
  'Etternavn, fornavn': 'EKSEMPEL KARI', Utbetalingsdato: '2025-06-15', Beløp: 1000, ...overrides
});

export const leaveRow = (overrides = {}) => ({
  Fornavn: 'Kari', Etternavn: 'Eksempel', Kjønn: 'Kvinne', Start: '2025-01-02', Slutt: '2025-01-03', ...overrides
});

export function reportFiles(year = '2025', period = '12-31') {
  const salary = [salaryRow(), salaryRow({ Fornavn: 'Per', 'Nøkkel for kjønn': 'Mann', '107A - Individuell lønn årsbe': 1000000 })];
  return [
    { role: 'fastlønn', label: 'Fastlønn', rows: salary, snapshotKey: `${year}-${period}`, snapshotLabel: `${period === '05-01' ? '01.05' : '31.12'}.${year}` },
    { role: 'overtid', label: 'Overtid', rows: [payRow(), payRow({ 'Etternavn, fornavn': 'EKSEMPEL PER', Beløp: 2000 })] },
    { role: 'vakttillegg', label: 'Vakttillegg', rows: [payRow({ Beløp: 5000 })] },
    { role: 'foreldrepermisjon', label: 'Foreldrepermisjon', rows: [leaveRow()] }
  ].map((file) => ({
    source: 'rådata', fileYear: year, years: [year], ...file,
    fileName: file.role === 'fastlønn' ? `Fastlønn per ${period === '05-01' ? '01.05' : '31.12'}.${year}.xlsx` : `${file.label} ${year}.xlsx`,
    rowCount: file.rows.length, columnCount: Object.keys(file.rows[0]).length, header: Object.keys(file.rows[0])
  }));
}

export function xlsxFile(name, rows, options = {}) {
  const workbook = XLSX.utils.book_new();
  const sheet = options.matrix ? XLSX.utils.aoa_to_sheet(rows) : XLSX.utils.json_to_sheet(rows);
  XLSX.utils.book_append_sheet(workbook, sheet, 'Rådata');
  if (options.mutate) options.mutate(workbook, sheet);
  return new File([XLSX.write(workbook, { type: 'buffer', bookType: 'xlsx' })], name);
}

export async function uploadPayload(file) {
  return { name: file.name, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', buffer: Buffer.from(await file.arrayBuffer()) };
}
