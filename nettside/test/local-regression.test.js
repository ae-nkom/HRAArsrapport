import test from 'node:test';
import assert from 'node:assert/strict';
import { readdir, readFile } from 'node:fs/promises';
import * as XLSX from 'xlsx';
import { parseExcelFiles } from '../src/lib/workbook-import.js';
import { buildReport, buildVariablePayReconciliation } from '../src/lib/report-engine.js';

const dataDir = new URL('../../data/', import.meta.url);
let local;
try {
  const names = await readdir(dataDir);
  const raw = names.filter((name) => name.endsWith('Raadata.xlsx'));
  const processedSalary = names.find((name) => name.includes('fastlonn') && name.includes('Bearbeidet') && name.endsWith('.xlsx'));
  const processedLeave = names.find((name) => name.includes('foreldrepermisjoner') && name.includes('Bearbeidet') && name.endsWith('.xlsx'));
  const processedPay = ['overtid', 'vakttillegg'].map((role) => names.find((name) => name.includes(role) && name.includes('Bearbeidet') && name.endsWith('.xlsx')));
  if (raw.length === 4 && processedSalary && processedLeave && processedPay.every(Boolean)) {
    const config = JSON.parse(await readFile(new URL('local-regression-config.json', dataDir), 'utf8'));
    const files = await parseExcelFiles(await Promise.all(raw.map(async (name) => new File([await readFile(new URL(name, dataDir))], name))));
    const reference = async (name) => {
      const workbook = XLSX.read(await readFile(new URL(name, dataDir)), { type: 'buffer' });
      return workbook.Sheets[workbook.SheetNames[0]];
    };
    const [year, month, day] = config.snapshotKey.split('-');
    local = { files, report: buildReport(files, year, `${month}-${day}`, config.assignments), salary: await reference(processedSalary), leave: await reference(processedLeave), pay: await Promise.all(processedPay.map(reference)) };
  }
} catch (error) {
  if (error.code !== 'ENOENT') throw error;
}
const options = { skip: !local && 'Lokale rådata, bearbeidede fasitfiler eller data/local-regression-config.json mangler. Ingen HR-data følger med prosjektet.' };
const near = (actual, expected, label) => assert.ok(Math.abs(actual - expected) < 1e-6, `${label}: avvik fra lokal fasit`);

test('lokal fasit: kjønnsfordeling og gruppevalg stemmer med bearbeidet fastlønn', options, () => {
  for (let row = 208; row <= 213; row++) {
    const group = String(local.salary[`I${row}`].v).trim().replace(/\*+$/, '');
    const actual = local.report.fastlonn.genderBalance.find((item) => item.group === group);
    assert.ok(actual, 'Fasitgruppen skal finnes');
    assert.equal(actual.women, local.salary[`J${row}`].v, `Kvinneantall i ${group}`);
    assert.equal(actual.men, local.salary[`L${row}`].v, `Mannsantall i ${group}`);
  }
  assert.equal(local.report.fastlonn.unknownPositionCodes.length, 0);
});

test('lokal fasit: lønnsgjennomsnitt samsvarer med de kontrollerte regnearkcellene', options, () => {
  const total = local.report.fastlonn.fastlonn[0];
  near(total.womenAvg, local.salary.J223.v, 'Gjennomsnitt kvinner');
  near(total.menAvg, local.salary.K223.v, 'Gjennomsnitt menn');
  near(total.totalAvg, local.salary.L223.v, 'Gjennomsnitt totalt');
  near(total.womenPctOfMen, local.salary.M223.v, 'Lønnsforhold');
});

test('lokal fasit: permisjonsdager og uker samsvarer med det manuelt kontrollerte arket', options, () => {
  const result = local.report.foreldrepermisjon;
  assert.equal(result.womenCount, local.leave.S6.v);
  assert.equal(result.menCount, local.leave.T6.v);
  assert.equal(result.womenDays, local.leave.U42.v);
  assert.equal(result.menDays, local.leave.V42.v);
  near(result.womenAvgWeeks, local.leave.U45.v, 'Permisjonsuker kvinner');
  near(result.menAvgWeeks, local.leave.V45.v, 'Permisjonsuker menn');
});

test('lokal fasit: kildetotaler avstemmes mot bearbeidede utbetalingsfiler', options, () => {
  ['overtid', 'vakttillegg'].forEach((role, index) => {
    const rows = XLSX.utils.sheet_to_json(local.pay[index], { defval: '' });
    const records = rows.filter((row) => row['Etternavn, fornavn'] && typeof row.Beløp === 'number' && (role !== 'overtid' || row.Utbetalingsdato));
    const expectedAmount = records.reduce((sum, row) => sum + row.Beløp, 0);
    const raw = local.files.find((file) => file.role === role).rows;
    const reconciliation = buildVariablePayReconciliation(raw, local.report.fastlonn.employeeIndex);
    near(reconciliation.sourceAmount, expectedAmount, `${role}: kildetotal`);
    near(reconciliation.sourceAmount, reconciliation.matchedAmount + reconciliation.unmatchedAmount, `${role}: avstemming`);
    assert.ok(reconciliation.unmatchedCount > 0, 'Kjente mangler mot årssluttuttrekket skal fortsatt vises');
    assert.ok(local.report.notes.some((note) => note.includes(`ansatte med ${role} er ikke`)));
  });
});

test('lokalt avvik: deltid følger rådata og overstyres ikke av historisk rapporttekst', options, () => {
  // The 2025 report says two men; the uploaded salary percentages contain none.
  // Keep the data result explicit instead of hardcoding the reported count.
  assert.deepEqual(local.report.employment.partTime, { women: 2, men: 0 });
  assert.deepEqual(local.report.employment.temporary, { women: 4, men: 5 });
});
