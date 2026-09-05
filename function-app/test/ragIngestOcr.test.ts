import assert from 'node:assert/strict';
import test from 'node:test';
import { createRagOcrIngestHandler } from '../src/functions/ragIngestOcr';

test('OCR ingestion route returns explicit OCR result totals', async () => {
  const handler = createRagOcrIngestHandler(async () => ({
    siteUrl: 'https://contoso.sharepoint.com/sites/legal',
    libraryName: 'Litigation Documents',
    discoveredFiles: 3,
    indexedDocuments: 1,
    indexedChunks: 2,
    skippedUnsupported: 1,
    skippedTooLarge: 0,
    requiresOcr: 1,
    ocrProcessed: 1,
    failedFiles: []
  }));

  const response = await handler({ json: async () => ({ libraryName: 'Litigation Documents' }) } as never, { error: () => undefined } as never);

  assert.equal(response.status, 200);
  assert.deepEqual(response.jsonBody, {
    ok: true,
    ingestion: {
      siteUrl: 'https://contoso.sharepoint.com/sites/legal',
      libraryName: 'Litigation Documents',
      discoveredFiles: 3,
      indexedDocuments: 1,
      indexedChunks: 2,
      skippedUnsupported: 1,
      skippedTooLarge: 0,
      requiresOcr: 1,
      ocrProcessed: 1,
      failedFiles: []
    }
  });
});
