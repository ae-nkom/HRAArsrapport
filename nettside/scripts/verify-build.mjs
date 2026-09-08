import { readdir, readFile, lstat } from 'node:fs/promises';
import { resolve, relative, extname } from 'node:path';
import { pathToFileURL } from 'node:url';

const extensions = new Set(['.html', '.js', '.css', '.svg', '.png', '.ico', '.webp', '.woff', '.woff2', '.json']);
export async function verifyTree(root, privateTerms = []) {
  const issues = [];
  async function walk(dir) {
    for (const name of await readdir(dir)) {
      const path = resolve(dir, name), label = relative(root, path);
      const stat = await lstat(path);
      if (stat.isSymbolicLink()) { issues.push(`${label}: symbolsk lenke`); continue; }
      if (stat.isDirectory()) {
        if (/^(data|default_data|test|test-results|node_modules)$/i.test(name)) issues.push(`${label}: intern mappe`);
        await walk(path);
      } else {
        if (name !== '_headers' && !extensions.has(extname(name))) issues.push(`${label}: filtype som ikke skal publiseres`);
        const content = (await readFile(path, 'utf8')).normalize('NFKC').toLocaleLowerCase('nb');
        if (privateTerms.some((term) => content.includes(term))) issues.push(`${label}: mulig personopplysning fra lokalt grunnlag`);
      }
    }
  }
  await walk(root);
  if (issues.length) throw new Error(`Publiseringskontrollen stoppet:\n${issues.join('\n')}`);
}

async function localPrivateTerms() {
  const dir = resolve('../data');
  let files;
  try { files = await readdir(dir); } catch (error) { if (error.code === 'ENOENT') return []; throw error; }
  const XLSX = await import('xlsx');
  const terms = new Set();
  for (const file of files.filter((name) => /fastlonn.*Raadata\.xlsx$/i.test(name))) {
    const workbook = XLSX.read(await readFile(resolve(dir, file)), { type: 'buffer' });
    for (const row of XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { defval: '' })) {
      const first = String(row.Fornavn || '').trim(), last = String(row.Etternavn || '').trim();
      if (first && last) for (const text of [`${first} ${last}`, `${last} ${first}`, `${last}, ${first}`]) terms.add(text.normalize('NFKC').toLocaleLowerCase('nb'));
      for (const value of Object.values(row)) if (/^\d{11}$/.test(String(value))) terms.add(String(value));
    }
  }
  return [...terms];
}

if (import.meta.url === pathToFileURL(resolve(process.argv[1])).href) {
  const terms = await localPrivateTerms();
  await verifyTree(resolve('static'), terms);
  await verifyTree(resolve('build'), terms);
  // Also check intermediate output, but only the build directory is for publication.
  await verifyTree(resolve('.svelte-kit/output'), terms);
  console.log(`Publiseringskontroll OK: tillatte filtyper, ingen interne mapper eller lenker${terms.length ? ', kontrollert mot lokale personopplysninger' : '. Lokale personkilder mangler; navnekontroll hoppet over'}.`);
}
