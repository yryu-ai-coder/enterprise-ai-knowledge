import { formatOcrIngestionMessage, getOcrIngestEndpoint } from './ocrIngestion';

describe('OCR ingestion helpers', () => {
  it('targets the distinct OCR ingestion route from the configured ask endpoint', () => {
    expect(getOcrIngestEndpoint('https://api.example.test/api/ask')).toBe('https://api.example.test/api/rag/ingest-ocr');
  });

  it('reports OCR processed files, indexed files and chunks, and errors', () => {
    expect(formatOcrIngestionMessage({
      discoveredFiles: 5,
      ocrProcessed: 2,
      indexedDocuments: 2,
      indexedChunks: 18,
      skippedUnsupported: 1,
      skippedTooLarge: 1,
      failedFiles: [{ name: 'unreadable.pdf', reason: 'OCR failed' }]
    })).toBe('OCR processed: 2/5 file(s); indexed: 2 file(s), 18 chunk(s); unsupported: 1; over limit: 1; errors: 1 (unreadable.pdf: OCR failed).');
  });
});
