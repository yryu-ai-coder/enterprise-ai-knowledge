import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { ingestSharePointLibrary } from '../services/ragIngestionService';

// App Service Authentication validates the Microsoft Entra bearer token first.
// Ingestion is explicit: this route is never called by a normal Ask AI request.
app.http('rag-ingest', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'rag/ingest',
  handler: async (request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    try {
      const body = await request.json() as { siteUrl?: string; libraryName?: string };
      const result = await ingestSharePointLibrary(body);
      return {
        status: 200,
        jsonBody: {
          ok: true,
          ingestion: result
        }
      };
    } catch (error) {
      const detail = (error as Error).message || 'Library ingestion failed.';
      context.error(detail);
      return {
        status: 500,
        jsonBody: {
          ok: false,
          error: 'Library ingestion failed.',
          detail
        }
      };
    }
  }
});
