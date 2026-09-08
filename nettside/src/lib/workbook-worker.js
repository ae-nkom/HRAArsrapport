import { parseExcelFiles } from './workbook-import.js';

self.onmessage = async ({ data }) => {
  try {
    self.postMessage({ files: await parseExcelFiles(data.files, data.role) });
  } catch (error) {
    self.postMessage({ error: error.message || 'Kunne ikke lese arbeidsboken.' });
  }
};
