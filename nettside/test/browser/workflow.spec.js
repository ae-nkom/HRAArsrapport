import { test, expect } from '@playwright/test';
import { readFile } from 'node:fs/promises';
import { unzipSync, strFromU8 } from 'fflate';
import { reportFiles, salaryRow, xlsxFile, uploadPayload } from '../fixtures.js';

async function upload(page, files = reportFiles()) {
  const payloads = await Promise.all(files.map((file) => uploadPayload(xlsxFile(file.fileName, file.rows))));
  await expect(page.locator('#batch-upload')).toBeEnabled();
  await page.locator('#batch-upload').setInputFiles(payloads);
  await expect(page.getByRole('status')).toHaveCount(0);
  await expect(page.getByRole('alert')).toHaveCount(0);
  await expect(page.getByRole('combobox', { name: 'Rapportår', exact: true })).toBeVisible();
}

test('hele arbeidsflyten foregår lokalt og gir en redigerbar Word-fil', async ({ page }) => {
  const errors = [], requests = [];
  page.on('pageerror', (error) => errors.push(error.message));
  page.on('request', (request) => requests.push({ url: request.url(), method: request.method() }));
  await page.goto('./');
  await upload(page);
  for (const name of ['Fastlønn', 'Overtid', 'Vakttillegg', 'Foreldrepermisjon']) {
    await page.getByRole('button', { name, exact: true }).click();
    await expect(page.locator('.local-chart svg').first()).toBeVisible();
    expect(await page.locator('main').innerText()).not.toMatch(/NaN|undefined|Infinity/);
  }
  await page.getByRole('button', { name: 'Årsrapport', exact: true }).click();
  await page.getByLabel('Antall nyansatte', { exact: true }).fill('3');
  await page.getByLabel('Samlet sykefravær %', { exact: true }).fill('4,5');
  await page.getByLabel('Legemeldt sykefravær kvinner %', { exact: true }).fill('5');
  await page.getByLabel('Legemeldt sykefravær menn %', { exact: true }).fill('3');
  const event = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Last ned rapportutkast (DOCX)', exact: true }).click();
  const download = await event;
  expect(download.suggestedFilename()).toBe('hr-arsrapport-2025.docx');
  const xml = strFromU8(unzipSync(await readFile(await download.path()))['word/document.xml']);
  expect(xml).toContain('31.12.2025');
  expect(xml).not.toMatch(/\[MANUELL:|Kari|Eksempel|undefined|NaN/);
  const storage = await page.evaluate(() => JSON.stringify(localStorage));
  expect(storage).not.toMatch(/Kari|Eksempel|xlsx|800000/);
  expect(requests.filter((request) => /^https?:/.test(request.url)).every((request) => request.method === 'GET' && new URL(request.url).origin === new URL(page.url()).origin)).toBe(true);
  expect(errors).toEqual([]);
  await page.reload();
  await expect(page.getByRole('combobox', { name: 'Rapportår', exact: true })).toHaveCount(0);
  await upload(page);
  await page.getByRole('button', { name: 'Årsrapport', exact: true }).click();
  await expect(page.getByLabel('Antall nyansatte', { exact: true })).toHaveValue('3');
});

test('lokale gruppevalg beregnes på nytt og nullstilles når uttrekket erstattes', async ({ page }) => {
  await page.goto('./');
  const files = reportFiles();
  files[0].rows[0] = salaryRow({ 'Stillingsgruppe betegnelse': '1059 Underdirektør' });
  await upload(page, files);
  const choice = page.getByLabel('Rapportgruppe for Eksempel Kari', { exact: true });
  await expect(choice).toHaveValue('Uavklart rapportgruppe');
  await choice.selectOption('Direktørgruppen');
  await expect(page.getByText('Alle gruppevalg er avklart.', { exact: true })).toBeVisible();
  await page.getByRole('button', { name: 'Fastlønn', exact: true }).click();
  await expect(page.locator('main')).not.toContainText('Uavklart rapportgruppe');
  await page.getByRole('button', { name: 'Datagrunnlag', exact: true }).click();
  await upload(page, [files[0]]);
  await expect(choice).toHaveValue('Uavklart rapportgruppe');
});

test('år og uttrekksdato kan velges og gir riktig eksportdato', async ({ page }) => {
  await page.goto('./');
  await upload(page, [reportFiles('2025', '12-31')[0], reportFiles('2025', '05-01')[0], reportFiles('2024', '12-31')[0]]);
  await page.getByRole('combobox', { name: 'Rapportår', exact: true }).selectOption('2024');
  await expect(page.getByRole('button', { name: '01.05', exact: true })).toBeDisabled();
  await page.getByRole('combobox', { name: 'Rapportår', exact: true }).selectOption('2025');
  await page.getByRole('button', { name: '01.05', exact: true }).click();
  await page.getByRole('button', { name: 'Årsrapport', exact: true }).click();
  const event = page.waitForEvent('download');
  await page.getByRole('button', { name: 'Last ned rapportutkast (DOCX)', exact: true }).click();
  const downloaded = await event;
  const xml = strFromU8(unzipSync(await readFile(await downloaded.path()))['word/document.xml']);
  expect(xml).toContain('Per 01.05.2025');
  expect(xml).not.toContain('ved slutten av året');
});

test('feil opplasting bevarer gjeldende data og gir en forståelig feilmelding', async ({ page }) => {
  await page.goto('./');
  await upload(page);
  await page.locator('#batch-upload').setInputFiles({ name: 'Ødelagt.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from('<html>Feil</html>') });
  await expect(page.getByRole('alert')).toContainText('gyldig XLSX');
  await expect(page.getByRole('combobox', { name: 'Rapportår', exact: true })).toHaveValue('2025');
  await expect(page.getByRole('alert')).not.toContainText('æ, ø eller å');
});

test('ugyldige manuelle tall stopper eksport og kan rettes eller slettes', async ({ page }) => {
  await page.goto('./');
  await upload(page);
  await page.getByRole('button', { name: 'Årsrapport', exact: true }).click();
  await page.getByLabel('Samlet sykefravær %', { exact: true }).fill('101');
  await expect(page.getByRole('button', { name: 'Last ned rapportutkast (DOCX)', exact: true })).toBeDisabled();
  await page.getByRole('button', { name: 'Slett lagrede rapportfelt for 2025', exact: true }).click();
  await expect(page.getByLabel('Samlet sykefravær %', { exact: true })).toHaveValue('');
  await expect(page.getByRole('button', { name: 'Last ned rapportutkast (DOCX)', exact: true })).toBeEnabled();
});

test('app og meny fungerer når lokal lagring er blokkert', async ({ page }) => {
  const errors = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.addInitScript(() => Object.defineProperty(window, 'localStorage', { get() { throw new DOMException('Blocked', 'SecurityError'); } }));
  await page.goto('./');
  await upload(page);
  await page.getByRole('button', { name: 'Skjul tekst i meny', exact: true }).click();
  await page.getByRole('button', { name: 'Fastlønn', exact: true }).click();
  await expect(page.locator('.local-chart svg').first()).toBeVisible();
  expect(errors).toEqual([]);
});

test('diagrammer oppdateres ved filbytte og popup har tastaturfokus', async ({ page }) => {
  await page.goto('./');
  await upload(page);
  await page.getByRole('button', { name: 'Overtid', exact: true }).click();
  const before = await page.locator('.local-chart').first().innerHTML();
  const expand = page.getByRole('button', { name: 'Åpne overtid per gruppe i større visning', exact: true });
  await expand.focus();
  await page.keyboard.press('Enter');
  await expect(page.getByRole('dialog')).toBeVisible();
  await expect(page.getByRole('button', { name: 'Lukk', exact: true })).toBeFocused();
  await page.keyboard.press('Escape');
  await expect(page.getByRole('dialog')).toHaveCount(0);
  await expect(expand).toBeFocused();
  await page.getByRole('button', { name: 'Datagrunnlag', exact: true }).click();
  const replacement = reportFiles()[1];
  replacement.rows[0].Beløp = 30000;
  await upload(page, [replacement]);
  await page.getByRole('button', { name: 'Overtid', exact: true }).click();
  await expect.poll(() => page.locator('.local-chart').first().innerHTML()).not.toBe(before);
});

test('arbeidsflatene passer på mobil uten vannrett sideskrolling', async ({ page }) => {
  await page.setViewportSize({ width: 390, height: 844 });
  await page.goto('./');
  await upload(page);
  for (const name of ['Datagrunnlag', 'Fastlønn', 'Overtid', 'Vakttillegg', 'Foreldrepermisjon', 'Årsrapport']) {
    await page.getByRole('button', { name, exact: true }).click();
    expect(await page.evaluate(() => document.documentElement.scrollWidth <= window.innerWidth + 1), name).toBe(true);
  }
});

test('kolonneoverskrifter og tall har samme høyrekant på bred skjerm', async ({ page }) => {
  await page.setViewportSize({ width: 3200, height: 1200 });
  await page.goto('./');
  await upload(page);
  for (const name of ['Fastlønn', 'Overtid', 'Vakttillegg', 'Foreldrepermisjon']) {
    await page.getByRole('button', { name, exact: true }).click();
    const cells = await page.locator('.report-table:visible').evaluateAll(tables => tables.flatMap(table => {
      const headers = [...table.querySelectorAll('thead th')];
      return [...table.querySelectorAll('tbody tr')].flatMap(row => headers.flatMap((head, index) => {
        if (!head.classList.contains('text-right')) return [];
        const cell = row.children[index], h = getComputedStyle(head), c = getComputedStyle(cell);
        return [{ header: h.textAlign, value: c.textAlign, delta: Math.abs((head.getBoundingClientRect().right - parseFloat(h.paddingRight)) - (cell.getBoundingClientRect().right - parseFloat(c.paddingRight))) }];
      }));
    }));
    expect(cells.length).toBeGreaterThan(0);
    expect(cells.every(cell => cell.header === 'right' && cell.value === 'right' && cell.delta < 1), name).toBe(true);
  }
  await page.getByRole('button', { name: 'Fastlønn', exact: true }).click();
  await expect(page.getByRole('columnheader', { name: 'Andel kvinner', exact: true }).first()).toBeVisible();
  await expect(page.getByRole('columnheader', { name: 'Andel menn', exact: true }).first()).toBeVisible();
  expect(await page.locator('.app-content').evaluate(element => element.getBoundingClientRect().width)).toBeLessThanOrEqual(1600);
  await expect(page.locator('main')).not.toContainText(/Største lønnsgap|er skjevest|Fokus på signaler/);
});

test('hele gruppenavn og grafverdier er lesbare på bred og smal skjerm', async ({ page }) => {
  await page.goto('./');
  const files = reportFiles();
  files[0].rows[0].Rapportgruppe = 'Øvrige saksbehandlere og andre';
  files[0].rows[1].Rapportgruppe = 'Direktørgruppen';
  files[0].rows[1]['107A - Individuell lønn årsbe'] = 1443791;
  await upload(page, files);
  for (const width of [1920, 390]) {
    await page.setViewportSize({ width, height: 1000 });
    for (const name of ['Fastlønn', 'Overtid', 'Vakttillegg', 'Foreldrepermisjon']) {
      await page.getByRole('button', { name, exact: true }).click();
      const labels = page.locator('.bar-category, .bar-value');
      expect(await labels.count()).toBeGreaterThan(0);
      const geometry = await labels.evaluateAll(nodes => nodes.map(node => {
        const box = node.getBoundingClientRect(), chart = node.closest('.local-chart').getBoundingClientRect();
        return { fits: node.scrollWidth <= node.clientWidth + 1 && box.left >= chart.left - 1 && box.right <= chart.right + 1, fontSize: parseFloat(getComputedStyle(node).fontSize), text: node.textContent };
      }));
      expect(geometry.every(item => item.fits && item.fontSize >= 12 && !item.text.includes('…')), `${name}, ${width}px`).toBe(true);
      expect(await page.evaluate(() => document.documentElement.scrollWidth <= innerWidth + 1)).toBe(true);
    }
    await page.getByRole('button', { name: 'Fastlønn', exact: true }).click();
    await expect(page.locator('.bar-category').filter({ hasText: /^Øvrige saksbehandlere og andre$/ }).first()).toBeVisible();
  }
});
