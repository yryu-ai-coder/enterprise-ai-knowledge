import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { randomUUID } from 'crypto';
import { ChatRequest, ChatResponse } from '../models';
import { answerQuestion } from '../services/aiProvider';

function errorResponse(requestId: string, statusCode: number, message: string): HttpResponseInit {
  const body: ChatResponse = {
    answer: '',
    citations: [],
    provider: process.env.AI_PROVIDER || 'mock',
    requestId,
    status: 'error',
    error: message
  };

  return { status: statusCode, jsonBody: body };
}

app.http('chat', {
  methods: ['POST'],
  authLevel: 'function',
  route: 'chat',
  handler: async (request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    const requestId = randomUUID();

    try {
      const body = await request.json() as Partial<ChatRequest>;
      const question = body.question?.trim();

      if (!question) {
        return errorResponse(requestId, 400, 'Question is required.');
      }

      if (question.length > 4000) {
        return errorResponse(requestId, 400, 'Question is too long.');
      }

      const chatRequest: ChatRequest = {
        question,
        scenario: body.scenario || 'legal-document-library',
        mode: body.mode || 'ask',
        contextType: body.contextType || 'document-library',
        siteUrl: body.siteUrl,
        listId: body.listId,
        listTitle: body.listTitle,
        libraryName: body.libraryName,
        folderPath: body.folderPath,
        pageUrl: body.pageUrl,
        pageTitle: body.pageTitle,
        pageContext: body.pageContext,
        selectedFiles: body.selectedFiles || [],
        selectedItems: body.selectedItems || [],
        documentSnippets: body.documentSnippets || [],
        conversationId: body.conversationId,
        knowledgeScope: body.knowledgeScope || body.scenario || 'legal-document-library'
      };

      context.log(`chat request ${requestId} provider=${process.env.AI_PROVIDER || 'mock'}`);
      const response = await answerQuestion(chatRequest, requestId);
      return { status: 200, jsonBody: response };
    } catch (error) {
      context.error(error);
      return errorResponse(requestId, 500, (error as Error).message || 'Unexpected backend error.');
    }
  }
});

