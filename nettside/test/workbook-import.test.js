import test from 'node:test';
import assert from 'node:assert/strict';
import { parseExcelFiles, parseSnapshotMeta, validateWorkbookArchive, mergeFiles, fileCoversYear } from '../src/lib/workbook-import.js';
import { xlsxFile, payRow, salaryRow } from './fixtures.js';

test('første datatabell finnes selv med tittel, blanke rader og mellomrom i kolonnenavn', async () => {
  const file = xlsxFile('Overtid 2025.xlsx', [['Årsuttrekk'], [], [' Etternavn, fornavn ', ' Utbetalingsdato ', ' Beløp '], ['Eksempel Kari', '2025-06-15', '1 250,50']], { matrix: true });
  const [parsed] = await parseExcelFiles([file]);
  assert.equal(parsed.rows.length, 1);
  assert.equal(parsed.rows[0].Beløp, '1 250,50');
  assert.equal(parsed.rows[0].__sourceRow, 4);
  assert.equal(parsed.header[0], 'Etternavn, fornavn');
});

test('filkortet bestemmer typen når overtidsfilen har et generisk navn', async () => {
  const file = xlsxFile('Uttrekk 2025.xlsx', [payRow()]);
  await assert.rejects(parseExcelFiles([file]), /kolonneoverskrift/);
  assert.equal((await parseExcelFiles([file], 'overtid'))[0].role, 'overtid');
});

test('filnavnet kan ikke overstyre manglende påkrevde kolonner', async () => {
  const row = salaryRow();
  delete row['107A - Individuell lønn årsbe'];
  await assert.rejects(parseExcelFiles([xlsxFile('Fastlønn per 31.12.2025.xlsx', [row])]), /årslønn/);
  await assert.rejects(parseExcelFiles([xlsxFile('Overtid 2025.xlsx', [payRow()])], 'vakttillegg'), /tilhører overtid/);
});

test('bearbeidede ark avvises og feil dato eller beløp blir ikke stille borte', async () => {
  await assert.rejects(parseExcelFiles([xlsxFile('Overtid bearbeidet 2025.xlsx', [payRow()])]), /rådatauttrekket/);
  await assert.rejects(parseExcelFiles([xlsxFile('Overtid 2025.xlsx', [payRow({ Utbetalingsdato: '31.02.2025' })])]), /gyldig utbetalingsdato/);
  await assert.rejects(parseExcelFiles([xlsxFile('Overtid 2025.xlsx', [payRow({ Beløp: '' })])]), /gyldig beløp/);
});

test('XLSX-signatur, arkivstørrelse og doble nøkkelkolonner kontrolleres', async () => {
  await assert.rejects(parseExcelFiles([new File(['<html>Hei</html>'], 'Overtid.xlsx')]), /gyldig XLSX/);
  const duplicate = xlsxFile('Overtid.xlsx', [['Etternavn, fornavn', 'Utbetalingsdato', 'Beløp', 'Beløp'], ['Eksempel Kari', '2025-01-02', 100, 200]], { matrix: true });
  await assert.rejects(parseExcelFiles([duplicate]), /forekommer flere ganger/);
  const bytes = new Uint8Array(await xlsxFile('Overtid.xlsx', [payRow()]).arrayBuffer());
  const view = new DataView(bytes.buffer);
  for (let i = 0; i < bytes.length - 46; i++) {
    if (view.getUint32(i, true) === 0x02014b50) { view.setUint32(i + 24, 101 * 1024 * 1024, true); break; }
  }
  assert.throws(() => validateWorkbookArchive(bytes), /100 MB/);
});

test('store arbeidsbøker avvises uten å avkorte datagrunnlaget', async () => {
  const file = xlsxFile('Overtid 2025.xlsx', Array.from({ length: 25001 }, () => payRow()));
  await assert.rejects(parseExcelFiles([file]), /25 000 datarader/);
});

test('uttakstidspunkt må være eksakt og gyldig', async () => {
  assert.equal(parseSnapshotMeta('Fastlønn per 311225.xlsx').snapshotKey, '2025-12-31');
  assert.equal(parseSnapshotMeta('Fastlønn per 01.05.2025.xlsx').snapshotKey, '2025-05-01');
  assert.equal(parseSnapshotMeta('Fastlønn 2025-12-31.xlsx').snapshotKey, '2025-12-31');
  await assert.rejects(parseExcelFiles([xlsxFile('Fastlønn per 31.02.2025.xlsx', [salaryRow()])]), /eksakt uttrekksdato/);
});

test('Excel-datoer er uavhengige av tidssone og støtter 1904-kalenderen', async () => {
  const original = process.env.TZ;
  try {
    for (const date1904 of [false, true]) {
      const file = xlsxFile('Overtid 2025.xlsx', [payRow()], { mutate(workbook, sheet) {
        workbook.Workbook = { WBProps: { date1904 } };
        sheet.B2 = { t: 'n', v: 45659 - (date1904 ? 1462 : 0), z: 'dd.mm.yyyy' };
      } });
      for (const timezone of ['Europe/Oslo', 'America/Los_Angeles', 'Pacific/Auckland']) {
        process.env.TZ = timezone;
        const [parsed] = await parseExcelFiles([file]);
        assert.equal(parsed.rows[0].Utbetalingsdato.toISOString().slice(0, 10), '2025-01-02', `${timezone}, 1904=${date1904}`);
      }
    }
  } finally {
    if (original === undefined) delete process.env.TZ; else process.env.TZ = original;
  }
});

test('flere års data vises i de riktige årene og overlappende filer erstattes', async () => {
  const [file] = await parseExcelFiles([xlsxFile('Overtid.xlsx', [payRow({ Utbetalingsdato: '2024-12-01' }), payRow()])]);
  assert.deepEqual(file.years, ['2024', '2025']);
  assert.equal(fileCoversYear(file, '2025'), true);
  const [replacement] = await parseExcelFiles([xlsxFile('Overtid 2025.xlsx', [payRow()])]);
  assert.deepEqual(mergeFiles([file], [replacement]), [replacement]);
  await assert.rejects(parseExcelFiles([xlsxFile('Overtid 2025.xlsx', [payRow()]), xlsxFile('Overtid ny 2025.xlsx', [payRow()])]), /Flere filer dekker/);
});
