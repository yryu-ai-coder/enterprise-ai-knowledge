import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { getLibrarySearchStatus } from '../services/ragSearchService';

// App Service Authentication validates the Microsoft Entra bearer token first.
app.http('rag-health', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'rag/health',
  handler: async (_request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    try {
      const status = await getLibrarySearchStatus();
      return {
        status: 200,
        jsonBody: {
          ok: true,
          service: 'nextcore-ai-function',
          rag: status
        }
      };
    } catch (error) {
      const detail = (error as Error).message || 'Azure AI Search connectivity check failed.';
      context.error(detail);
      return {
        status: 500,
        jsonBody: {
          ok: false,
          error: 'Azure AI Search connectivity check failed.',
          detail
        }
      };
    }
  }
});
