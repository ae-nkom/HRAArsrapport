import test from 'node:test';
import assert from 'node:assert/strict';
import { parseReportDate } from '../src/lib/date-utils.js';
import { buildParentalLeaveEmployeesByYear, buildParentalLeaveGroupedEmployees, buildEmployeeIndex } from '../src/lib/report-engine.js';
import { salaryRow, leaveRow } from './fixtures.js';

test('ugyldige datoer rulles ikke videre til neste måned', () => {
  for (const date of ['31.02.2025', '2025-02-29', '2025-13-01', '00.01.2025', 60]) assert.equal(parseReportDate(date), null);
  assert.equal(parseReportDate('29.02.2024').toISOString().slice(0, 10), '2024-02-29');
  assert.equal(parseReportDate(45659).toISOString().slice(0, 10), '2025-01-02');
});

test('overlappende permisjonsperioder teller hver arbeidsdag bare én gang', () => {
  const result = buildParentalLeaveEmployeesByYear([leaveRow({ Start: '2025-01-02', Slutt: '2025-01-06' }), leaveRow({ Start: '2025-01-03', Slutt: '2025-01-07' })], '2025');
  assert.equal(result[0].days, 4);
  assert.equal(result[0].overlappingDays, 2);
});

test('gruppeoversikten og totaloversikten utelater permisjon uten arbeidsdager', () => {
  const rows = [leaveRow({ Start: '2025-01-04', Slutt: '2025-01-05' })];
  const index = buildEmployeeIndex([salaryRow()]);
  assert.equal(buildParentalLeaveEmployeesByYear(rows, '2025').length, 0);
  assert.equal(buildParentalLeaveGroupedEmployees(rows, index, '2025').length, 0);
});
