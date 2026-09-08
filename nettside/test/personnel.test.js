import test from 'node:test';
import assert from 'node:assert/strict';
import { buildEmployeeIndex, employeeGroup, employeeKey, excludedGroup, unassignedGroup, matchEmployee, summarizeFastlonn, summarizeEmployment, buildReport } from '../src/lib/report-engine.js';
import { normalizePersonName, parseNumber, formatCurrency } from '../src/lib/formatting.js';
import { salaryRow, reportFiles } from './fixtures.js';

test('underdirektører krever oppgitt personalansvar eller et lokalt gruppevalg', () => {
  const row = salaryRow({ 'Stillingsgruppe betegnelse': '1059 Underdirektør' });
  assert.equal(employeeGroup(row), unassignedGroup);
  assert.equal(employeeGroup({ ...row, Personalansvar: 'Ja' }), 'Direktørgruppen');
  assert.equal(employeeGroup({ ...row, Personalansvar: 'Nei' }), 'Fagsjefgruppen');
  assert.equal(employeeGroup(row, { [employeeKey(row)]: 'Direktørgruppen' }), 'Direktørgruppen');
  assert.equal(employeeGroup(row), unassignedGroup, 'Valget må ikke endre rådataene');
});

test('ukjente grupper beholdes i totaltallet inntil en eksplisitt avgrensning er gjort', () => {
  const row = salaryRow({ 'Stillingsgruppe betegnelse': '9999 Ukjent stilling' });
  const before = summarizeFastlonn([row]);
  assert.equal(before.genderBalance[0].total, 1);
  assert.equal(before.unknownPositionCodes[0].code, '9999');
  const after = summarizeFastlonn([row], null, { [employeeKey(row)]: excludedGroup });
  assert.equal(after.genderBalance[0].total, 0);
  assert.equal(after.employeeIndex.excludedEmployees.length, 1);
});

test('navnelike personer med forskjellige ansattnumre holdes atskilt', () => {
  const index = buildEmployeeIndex([salaryRow({ Ansattnummer: '0010' }), salaryRow({ Ansattnummer: '0020', '107A - Individuell lønn årsbe': 900000 })]);
  assert.equal(index.employees.length, 2);
  assert.equal(matchEmployee({ 'Etternavn, fornavn': 'Eksempel Kari' }, index), null);
  assert.equal(matchEmployee({ Ansattnummer: 20, 'Etternavn, fornavn': 'Eksempel Kari' }, index).salary, 900000);
  assert.ok(index.ambiguousNames.size);
});

test('motstridende personrader uten entydig ID blir ikke vilkårlig valgt', () => {
  const index = buildEmployeeIndex([salaryRow(), salaryRow({ '107A - Individuell lønn årsbe': 900000 })]);
  assert.equal(index.employees.length, 0);
  assert.equal(index.duplicatePeople.size, 1);
  assert.equal(matchEmployee({ 'Etternavn, fornavn': 'Eksempel Kari' }, index), null);
  assert.equal(buildEmployeeIndex([salaryRow(), salaryRow()]).employees.length, 1);
});

test('navnenormalisering bevarer norske bokstaver og skiller ulike navn', () => {
  assert.equal(normalizePersonName('  SØR, Åse  '), 'sør åse');
  assert.notEqual(normalizePersonName('Sør Åse'), normalizePersonName('Sor Ase'));
});

test('manglende lønn er utilgjengelig og blir aldri gjennomsnittslønn på null', () => {
  const result = summarizeFastlonn([salaryRow({ '107A - Individuell lønn årsbe': '' })]);
  assert.equal(result.fastlonn[0].totalAvg, null);
  assert.equal(result.fastlonn[0].womenPctOfMen, null);
  assert.equal(formatCurrency(null), '—');
});

test('norske tallstrenger og null prosent stilling håndteres uten standardverdier', () => {
  assert.equal(parseNumber('1 234,50'), 1234.5);
  assert.equal(parseNumber('1.234.567,89'), 1234567.89);
  assert.equal(parseNumber(''), null);
  assert.equal(parseNumber('feil'), null);
  const index = buildEmployeeIndex([salaryRow({ 'Deltids-% 0007': 0 }), salaryRow({ Fornavn: 'Liv', 'Deltids-% 0007': '' })]);
  assert.equal(index.employees[0].partTimePercent, 0);
  assert.equal(index.employees[1].partTimePercent, null);
  assert.equal(summarizeEmployment(index.employees).partTime.women, 1);
});

test('ansettelsesdato filtreres mot valgt uttrekksdato', () => {
  const rows = [salaryRow({ Ansettelsesdato: '2025-06-01' })];
  assert.equal(buildEmployeeIndex(rows, new Date('2025-05-01')).employees.length, 0);
  assert.equal(buildEmployeeIndex(rows, new Date('2025-12-31')).employees.length, 1);
});

test('rapporten krever et eksakt uttrekk og omtaler mangler', () => {
  const result = buildReport(reportFiles(), '2025', '05-01');
  assert.equal(result.fastlonn, null);
  assert.ok(result.notes.some((note) => note.includes('Mangler fastlønnfil')));
});
