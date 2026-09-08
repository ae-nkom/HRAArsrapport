// File contents stay on this device. A worker keeps the interface responsive and
// can be terminated if a malformed workbook takes too long to process.
export function readExcelFiles(files, role) {
  return new Promise((resolve, reject) => {
    const worker = new Worker(new URL('./workbook-worker.js', import.meta.url), { type: 'module' });
    const timeout = setTimeout(() => finish(new Error('Lesingen tok for lang tid. Prøv ett mindre uttrekk om gangen.')), 30_000);
    function finish(error, result) {
      clearTimeout(timeout);
      worker.terminate();
      if (error) reject(error); else resolve(result);
    }
    worker.onmessage = ({ data }) => finish(data.error ? new Error(data.error) : null, data.files);
    worker.onerror = () => finish(new Error('Kunne ikke starte Excel-lesingen. Last siden på nytt og prøv igjen.'));
    worker.postMessage({ files: Array.from(files), role });
  });
}
