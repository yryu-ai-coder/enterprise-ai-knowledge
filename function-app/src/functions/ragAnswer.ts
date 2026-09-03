import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { ChatRequest } from '../models';
import { answerQuestion } from '../services/aiProvider';
import { buildGroundedLibraryRequest } from '../services/ragAnswerService';
import { searchLibraryChunks } from '../services/ragSearchService';

// App Service Authentication validates the Microsoft Entra bearer token first.
app.http('rag-answer', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'rag/answer',
  handler: async (request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    const requestId = context.invocationId;
    try {
      const body = await request.json() as ChatRequest;
      const question = body.question?.trim();
      if (!question) {
        return { status: 400, jsonBody: { error: 'Question is required.' } };
      }
      if (question.length > 1_000) {
        return { status: 400, jsonBody: { error: 'Question is too long.' } };
      }

      const results = await searchLibraryChunks(question);
      if (results.length === 0) {
        return {
          status: 200,
          jsonBody: {
            answer: 'No matching indexed document excerpts were found in this library. Run library ingestion, or refine the question.',
            citations: [],
            provider: 'rag-search',
            requestId,
            status: 'success',
            metadata: { retrievalCount: 0, knowledgeScope: 'library-wide-rag' }
          }
        };
      }

      const groundedRequest = buildGroundedLibraryRequest({ ...body, question }, results);
      const response = await answerQuestion(groundedRequest, requestId);
      return {
        status: 200,
        jsonBody: {
          ...response,
          metadata: {
            ...(response.metadata || {}),
            retrievalCount: results.length,
            knowledgeScope: 'library-wide-rag'
          }
        }
      };
    } catch (error) {
      const detail = (error as Error).message || 'Library-wide RAG answer failed.';
      context.error(detail);
      return {
        status: 500,
        jsonBody: {
          status: 'error',
          error: 'Library-wide RAG answer failed.',
          detail,
          requestId
        }
      };
    }
  }
});
