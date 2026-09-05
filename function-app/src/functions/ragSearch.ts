import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { searchLibraryChunks } from '../services/ragSearchService';

interface RagSearchRequest {
  query?: string;
}

// App Service Authentication validates the Microsoft Entra bearer token first.
app.http('rag-search', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'rag/search',
  handler: async (request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    try {
      const body = await request.json() as RagSearchRequest;
      const query = body.query?.trim();

      if (!query) {
        return { status: 400, jsonBody: { error: 'Search query is required.' } };
      }
      if (query.length > 1000) {
        return { status: 400, jsonBody: { error: 'Search query is too long.' } };
      }

      const results = await searchLibraryChunks(query);
      return {
        status: 200,
        jsonBody: {
          results: results.map(result => ({
            documentName: result.documentName,
            documentUrl: result.documentUrl,
            folderPath: result.folderPath,
            fileType: result.fileType,
            lastModified: result.lastModified,
            pageNumber: result.pageNumber,
            excerpt: result.content.slice(0, 600),
            score: result.score
          }))
        }
      };
    } catch (error) {
      context.error(error);
      return { status: 500, jsonBody: { error: 'Library search failed.' } };
    }
  }
});
