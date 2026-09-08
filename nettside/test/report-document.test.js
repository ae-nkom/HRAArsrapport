import test from 'node:test';
import assert from 'node:assert/strict';
import { Packer } from 'docx';
import { unzipSync, strFromU8 } from 'fflate';
import { buildReport } from '../src/lib/report-engine.js';
import { standardReportNarratives, validateManualField, normalizeManualReportInputs, reportHistoryYears } from '../src/lib/report-text.js';
import { createReportDocument } from '../src/lib/report-document.js';
import { reportFiles } from './fixtures.js';

test('mai-uttrekk omtales med riktig dato i bemanning, kjønn og deltid', () => {
  const report = buildReport(reportFiles('2025', '05-01'), '2025', '05-01');
  const texts = standardReportNarratives(report, '2025');
  for (const field of ['introduction', 'genderBalance', 'employment']) {
    assert.match(texts[field], /01\.05\.2025/);
    assert.doesNotMatch(texts[field], /ved slutten av året|Ved utgangen av|ved utgangen av/);
  }
  assert.match(texts.overtid, /blant koblede mottakere/);
  assert.doesNotMatch(texts.vakttillegg, /tre beredskapsvaktordninger/);
});

test('manuelle tallfelt validerer prosent, heltall og tomme verdier', () => {
  assert.equal(validateManualField('nyansatteCount', '0'), '');
  for (const value of ['-1', '1,5', '10%', 'abc']) assert.ok(validateManualField('nyansatteCount', value));
  for (const value of ['-1', '101', 'NaN', '']) assert.ok(validateManualField('samletSykefravaer', value));
  assert.equal(validateManualField('samletSykefravaer', '7,5 %'), '');
  assert.deepEqual(normalizeManualReportInputs(null), normalizeManualReportInputs({ ukjent: 'tekst', nyansatteCount: {} }));
});

test('Word-eksporten inneholder riktige tabeller, dato og redigerbare felt uten personrader', async () => {
  const files = reportFiles('2025', '05-01');
  const summary = buildReport(files, '2025', '05-01');
  const bytes = await Packer.toBuffer(createReportDocument(summary, { year: '2025', files }));
  const zip = unzipSync(bytes);
  const xml = strFromU8(zip['word/document.xml']);
  assert.match(xml, /01\.05\.2025/);
  assert.match(xml, /Rapportutkast/);
  assert.match(xml, /\[MANUELL:/);
  assert.match(xml, /w:tblHeader/);
  assert.match(xml, /w:cantSplit/);
  assert.match(xml, /Mottakere i valgt uttrekk/);
  assert.doesNotMatch(xml, /Kari|Eksempel|undefined|NaN|Infinity/);
  assert.doesNotMatch(xml, /to underdirektører|ved slutten av året/);
  assert.doesNotMatch(xml, /w:documentProtection/);
});

test('Word-eksport avviser ugyldige tall og manglende fastlønnsgrunnlag', () => {
  const summary = buildReport(reportFiles(), '2025', '12-31');
  assert.throws(() => createReportDocument(summary, { manualInputs: { samletSykefravaer: '125' } }), /manuelle tallfeltene/);
  assert.throws(() => createReportDocument({}), /fastlønnsuttrekk/);
});


test('historikktabellen følger malens seks år og bruker tydelige andelsoverskrifter', async () => {
  assert.deepEqual(reportHistoryYears('2025'), ['2020', '2021', '2022', '2023', '2024', '2025']);
  const files = reportFiles();
  const xml = strFromU8(unzipSync(await Packer.toBuffer(createReportDocument(buildReport(files, '2025', '12-31'), { year: '2025', files })))['word/document.xml']);
  assert.match(xml, /Andel kvinner/);
  assert.match(xml, /Andel menn/);
  const history = xml.match(/<w:tbl>.*?<\/w:tbl>/gs).find(table => table.includes('Prosentoppnåelse'));
  assert.ok(history.includes('2020'));
  for (const row of history.match(/<w:tr>.*?<\/w:tr>/gs)) assert.equal((row.match(/<w:tc>/g) || []).length, 7);
});
