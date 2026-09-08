import test from 'node:test';
import assert from 'node:assert/strict';
import { buildEmployeeIndex, buildVariablePayParticipants, buildVariablePayReconciliation, summarizeVariablePay } from '../src/lib/report-engine.js';
import { buildVariablePayMonthlyRows, buildVariablePayGroupMetrics } from '../src/lib/report-analytics.js';
import { salaryRow, payRow } from './fixtures.js';

test('mottakere, gjennomsnitt og avstemming bruker samme positive årsbeløp', () => {
  const index = buildEmployeeIndex([salaryRow(), salaryRow({ Fornavn: 'Per', 'Nøkkel for kjønn': 'Mann' }), salaryRow({ Fornavn: 'Liv' })]);
  const rows = [payRow({ Beløp: 1200 }), payRow({ Beløp: -200 }), payRow({ 'Etternavn, fornavn': 'Eksempel Per', Beløp: 0 }), payRow({ 'Etternavn, fornavn': 'Eksempel Liv', Beløp: -100 }), payRow({ 'Etternavn, fornavn': 'Ukjent Person', Beløp: 500 })];
  const participants = buildVariablePayParticipants(rows, index);
  const totals = buildVariablePayReconciliation(rows, index);
  assert.equal(participants.length, 1);
  assert.equal(participants[0].amount, 1000);
  assert.equal(totals.sourceCount, 2);
  assert.equal(totals.sourceAmount, 1500);
  assert.equal(totals.netSourceAmount, 1400);
  assert.equal(totals.correctionAmount, -100);
  assert.equal(totals.unmatchedAmount, 500);
  assert.equal(totals.sourceAmount, totals.matchedAmount + totals.unmatchedAmount);
  const summary = summarizeVariablePay(rows, index)[0];
  assert.equal(summary.totalAvg, 1000);
  assert.equal(summary.menAvg, null);
  assert.equal(summary.womenPctOfMen, null);
  assert.ok(Math.abs(buildVariablePayGroupMetrics(participants, index.employees)[0].participationShare - 100 / 3) < 1e-10);
});

test('SAP-tellere i ansattnummerkolonnen slår ikke sammen mottakere', () => {
  const index = buildEmployeeIndex([salaryRow(), salaryRow({ Fornavn: 'Per', 'Nøkkel for kjønn': 'Mann' })]);
  const rows = [payRow({ Ansattnummer: 1 }), payRow({ Ansattnummer: 1, 'Etternavn, fornavn': 'Eksempel Per', Beløp: 2000 })];
  assert.equal(buildVariablePayParticipants(rows, index).length, 2);
  assert.equal(buildVariablePayReconciliation(rows, index).matchedAmount, 3000);
});

test('ulike kilde-ID-er med samme navn blir ikke koblet vilkårlig', () => {
  const index = buildEmployeeIndex([salaryRow()]);
  const rows = [payRow({ Ansattnummer: 10 }), payRow({ Ansattnummer: 20 })];
  assert.equal(buildVariablePayParticipants(rows, index).length, 0);
  assert.equal(buildVariablePayReconciliation(rows, index).unmatchedCount, 2);
});

test('månedsserier summerer til samme beløp som mottakertabellen og bevarer korreksjoner', () => {
  const index = buildEmployeeIndex([salaryRow()]);
  const rows = [payRow({ Utbetalingsdato: '2025-01-10', Beløp: 1200 }), payRow({ Utbetalingsdato: '2025-02-10', Beløp: -200 })];
  const monthly = buildVariablePayMonthlyRows(rows, index);
  assert.equal(monthly.length, 12);
  assert.equal(monthly[1].totalt, -200);
  assert.equal(monthly.reduce((sum, row) => sum + row.totalt, 0), 1000);
  assert.equal(monthly[11].totalt, 0);
});
