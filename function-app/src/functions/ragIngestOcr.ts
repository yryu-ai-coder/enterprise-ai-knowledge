import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { ingestSharePointLibraryOcr, OcrLibraryIngestionResult } from '../services/ragOcrIngestionService';

export type OcrIngestionOperation = (scope: { siteUrl?: string; libraryName?: string }) => Promise<OcrLibraryIngestionResult>;

export function createRagOcrIngestHandler(ingest: OcrIngestionOperation = (scope) => ingestSharePointLibraryOcr({}, scope)) {
  return async (request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    try {
      const scope = await request.json() as { siteUrl?: string; libraryName?: string };
      const result = await ingest(scope);
      return {
        status: 200,
        jsonBody: {
          ok: true,
          ingestion: result
        }
      };
    } catch (error) {
      const detail = (error as Error).message || 'OCR library ingestion failed.';
      context.error(detail);
      return {
        status: 500,
        jsonBody: {
          ok: false,
          error: 'OCR library ingestion failed.',
          detail
        }
      };
    }
  };
}

// App Service Authentication validates the Microsoft Entra bearer token first.
// OCR ingestion is explicit and only processes PDFs whose text-layer extraction requires it.
app.http('rag-ingest-ocr', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'rag/ingest-ocr',
  handler: createRagOcrIngestHandler()
});
