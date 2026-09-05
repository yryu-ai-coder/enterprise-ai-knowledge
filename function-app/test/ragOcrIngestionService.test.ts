import assert from 'node:assert/strict';
import test from 'node:test';
import {
  ingestSharePointLibraryOcr,
  OcrLibraryIngestionDependencies
} from '../src/services/ragOcrIngestionService';

const environment = {
  RAG_SHAREPOINT_SITE_URL: 'https://contoso.sharepoint.com/sites/legal',
  RAG_SHAREPOINT_LIBRARY_NAME: 'Litigation Documents',
  DOCUMENT_INTELLIGENCE_ENDPOINT: 'https://contoso.cognitiveservices.azure.com',
  DOCUMENT_INTELLIGENCE_API_KEY: 'test-key',
  OCR_MAX_PDF_BYTES: '1024',
  OCR_MAX_PAGES: '3'
};

function createDependencies(): OcrLibraryIngestionDependencies & { ocrInputs: Buffer[]; indexedChunks: unknown[][] } {
  const ocrInputs: Buffer[] = [];
  const indexedChunks: unknown[][] = [];
  const graphJson = async <T>(url: string): Promise<T> => {
    if (url.endsWith('/drives')) return { value: [{ id: 'drive-1', name: 'Litigation Documents' }] } as T;
    if (url.includes('/sites/')) return { id: 'site-1' } as T;
    return {
      value: [
        { id: 'scanned', name: 'scanned.pdf', size: 50, webUrl: 'https://contoso/scanned.pdf', file: { mimeType: 'application/pdf' }, lastModifiedDateTime: '2026-09-03T12:00:00Z', parentReference: { path: '/drives/drive-1/root:/Scans' } },
        { id: 'digital', name: 'digital.pdf', size: 50, webUrl: 'https://contoso/digital.pdf', file: { mimeType: 'application/pdf' }, lastModifiedDateTime: '2026-09-03T12:01:00Z', parentReference: { path: '/drives/drive-1/root:' } },
        { id: 'word', name: 'notes.docx', size: 50, file: { mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' } }
      ]
    } as T;
  };

  return {
    environment,
    getGraphAccessToken: async () => 'graph-token',
    graphJson,
    graphFetch: async (url: string) => new Response(url.includes('/scanned/') ? Buffer.from('scanned-pdf') : Buffer.from('digital-pdf')),
    extractPdfText: async (bytes: Buffer) => bytes.toString() === 'scanned-pdf'
      ? { text: '', requiresOcr: true }
      : { text: 'This text-layer PDF must not be submitted for OCR.', requiresOcr: false },
    runOcr: async (bytes: Buffer) => {
      ocrInputs.push(bytes);
      return {
        text: 'First page text\n\nSecond page text',
        pages: [
          { pageNumber: 1, text: 'First page text' },
          { pageNumber: 2, text: 'Second page text' }
        ]
      };
    },
    indexLibraryChunks: async (chunks) => {
      indexedChunks.push(chunks);
      return chunks.length;
    },
    ocrInputs,
    indexedChunks
  };
}

test('ingestSharePointLibraryOcr submits only text-layer OCR-required PDFs and indexes page-aware citation chunks', async () => {
  const dependencies = createDependencies();

  const result = await ingestSharePointLibraryOcr(dependencies);

  assert.deepEqual(dependencies.ocrInputs.map(bytes => bytes.toString()), ['scanned-pdf']);
  assert.equal(result.discoveredFiles, 3);
  assert.equal(result.requiresOcr, 1);
  assert.equal(result.ocrProcessed, 1);
  assert.equal(result.indexedDocuments, 1);
  assert.equal(result.indexedChunks, 2);
  assert.equal(result.skippedUnsupported, 1);
  assert.equal(result.failedFiles.length, 0);

  const chunks = dependencies.indexedChunks.flat() as Array<Record<string, unknown>>;
  assert.deepEqual(chunks.map(chunk => chunk.pageNumber), [1, 2]);
  assert.deepEqual(chunks.map(chunk => chunk.content), ['First page text', 'Second page text']);
  assert.deepEqual(chunks.map(chunk => chunk.documentName), ['scanned.pdf', 'scanned.pdf']);
  assert.deepEqual(chunks.map(chunk => chunk.documentUrl), ['https://contoso/scanned.pdf', 'https://contoso/scanned.pdf']);
  assert.deepEqual(chunks.map(chunk => chunk.folderPath), ['/Litigation Documents/Scans', '/Litigation Documents/Scans']);
  assert.deepEqual(chunks.map(chunk => chunk.fileType), ['pdf', 'pdf']);
  assert.deepEqual(chunks.map(chunk => chunk.lastModified), ['2026-09-03T12:00:00Z', '2026-09-03T12:00:00Z']);
});
