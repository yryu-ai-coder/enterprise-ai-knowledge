export interface IOcrIngestionResult {
  discoveredFiles: number;
  ocrProcessed: number;
  indexedDocuments: number;
  indexedChunks: number;
  skippedUnsupported: number;
  skippedTooLarge: number;
  failedFiles: Array<{ name: string; reason: string }>;
}

export function getOcrIngestEndpoint(functionEndpoint: string): string {
  return functionEndpoint
    .replace('/api/chat', '/api/rag/ingest-ocr')
    .replace('/api/ask', '/api/rag/ingest-ocr');
}

export function formatOcrIngestionMessage(ingestion: IOcrIngestionResult): string {
  const errors = ingestion.failedFiles.length
    ? ` (${ingestion.failedFiles.map(file => `${file.name}: ${file.reason}`).join('; ')})`
    : '';

  return `OCR processed: ${ingestion.ocrProcessed}/${ingestion.discoveredFiles} file(s); indexed: ${ingestion.indexedDocuments} file(s), ${ingestion.indexedChunks} chunk(s); unsupported: ${ingestion.skippedUnsupported}; over limit: ${ingestion.skippedTooLarge}; errors: ${ingestion.failedFiles.length}${errors}.`;
}
