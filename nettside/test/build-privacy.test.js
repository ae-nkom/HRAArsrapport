import test from 'node:test';
import assert from 'node:assert/strict';
import { mkdtemp, mkdir, writeFile, symlink, rm } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import { verifyTree } from '../scripts/verify-build.mjs';

test('publiseringskontrollen stopper HR-filer, personnavn og symbolske lenker', async () => {
  const root = await mkdtemp(join(tmpdir(), 'hr-build-check-'));
  try {
    await writeFile(join(root, 'index.html'), '<html lang="nb">HR Årsrapport</html>');
    await verifyTree(root, ['kari eksempel']);
    await writeFile(join(root, 'app.js'), 'const name = "KARI EKSEMPEL";');
    await assert.rejects(verifyTree(root, ['kari eksempel']), /mulig personopplysning/);
    await rm(join(root, 'app.js'));
    await writeFile(join(root, 'salary.xlsx'), 'private');
    await assert.rejects(verifyTree(root), /filtype/);
    await rm(join(root, 'salary.xlsx'));
    await symlink('/tmp', join(root, 'linked'));
    await assert.rejects(verifyTree(root), /symbolsk lenke/);
    await rm(join(root, 'linked'));
    await mkdir(join(root, 'data'));
    await assert.rejects(verifyTree(root), /intern mappe/);
  } finally { await rm(root, { recursive: true, force: true }); }
});
