import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { randomUUID } from 'crypto';
import { ChatRequest, ChatResponse } from '../models';
import { answerQuestion } from '../services/aiProvider';
import { extractPdfText } from '../services/pdfTextExtraction';

const MAX_SELECTED_PDF_BYTES = 4 * 1024 * 1024;
const MAX_DOCUMENT_TEXT_CHARS = 16000;

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

// App Service Authentication validates the Microsoft Entra bearer token before this
// handler runs. Keep the Functions host key out of SPFx browser code.
app.http('chat', {
  methods: ['POST'],
  authLevel: 'anonymous',
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

      const selectedDocument = body.selectedDocument;
      let extractedDocumentText = '';

      if (selectedDocument) {
        if (selectedDocument.fileType?.toLowerCase() !== 'pdf') {
          return errorResponse(requestId, 400, 'This proof of concept currently supports selected PDF files only.');
        }

        const documentBytes = Buffer.from(selectedDocument.contentBase64 || '', 'base64');
        if (documentBytes.length === 0) {
          return errorResponse(requestId, 400, 'Selected PDF content is empty.');
        }
        if (documentBytes.length > MAX_SELECTED_PDF_BYTES) {
          return errorResponse(requestId, 413, 'Selected PDF exceeds the 4 MB proof-of-concept limit.');
        }

        const extraction = await extractPdfText(documentBytes);
        if (extraction.requiresOcr) {
          return errorResponse(requestId, 422, 'No selectable text was found in the selected PDF. OCR is required before Ask AI can analyze it.');
        }

        extractedDocumentText = extraction.text.slice(0, MAX_DOCUMENT_TEXT_CHARS);
      }

      const selectedFiles = (body.selectedFiles || []).map((file) => {
        if (extractedDocumentText && selectedDocument && file.name === selectedDocument.name) {
          return { ...file, snippet: extractedDocumentText.slice(0, 800) };
        }
        return file;
      });

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
        selectedFiles,
        selectedItems: body.selectedItems || [],
        documentSnippets: extractedDocumentText
          ? [...(body.documentSnippets || []), `Extracted text from ${selectedDocument?.name || 'selected PDF'}:\n${extractedDocumentText}`]
          : body.documentSnippets || [],
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

