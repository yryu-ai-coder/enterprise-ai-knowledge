import { createServer, IncomingMessage, ServerResponse } from 'http';
import { existsSync, readFileSync } from 'fs';
import { join } from 'path';
import { randomUUID } from 'crypto';
import { answerQuestion } from './services/aiProvider';
import { extractPdfText } from './services/pdfTextExtraction';
import { ChatRequest, ChatResponse } from './models';

const MAX_SELECTED_PDF_BYTES = 4 * 1024 * 1024;
const MAX_DOCUMENT_TEXT_CHARS = 16000;

function loadLocalSettings(): void {
  const settingsPath = join(process.cwd(), 'local.settings.json');

  if (!existsSync(settingsPath)) {
    console.log('local.settings.json not found. Using existing process environment variables.');
    return;
  }

  try {
    const parsed = JSON.parse(readFileSync(settingsPath, 'utf8')) as { Values?: Record<string, string> };
    const values = parsed.Values || {};
    let loadedCount = 0;

    for (const [key, value] of Object.entries(values)) {
      if (typeof value === 'string' && value.length > 0 && !process.env[key]) {
        process.env[key] = value;
        loadedCount += 1;
      }
    }

    console.log(`Loaded ${loadedCount} setting(s) from local.settings.json.`);
    console.log(`AI provider: ${process.env.AI_PROVIDER || 'mock'}`);
  } catch (error) {
    console.warn(`Could not load local.settings.json: ${(error as Error).message}`);
  }
}

loadLocalSettings();

const port = Number(process.env.PORT || 7071);

function sendJson(response: ServerResponse, statusCode: number, body: unknown): void {
  response.writeHead(statusCode, {
    'Content-Type': 'application/json',
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type,Authorization'
  });
  response.end(JSON.stringify(body));
}

function readBody(request: IncomingMessage): Promise<string> {
  return new Promise((resolve, reject) => {
    let body = '';
    request.on('data', chunk => {
      body += chunk;
      if (body.length > 6 * 1024 * 1024) {
        reject(new Error('Request body too large.'));
      }
    });
    request.on('end', () => resolve(body));
    request.on('error', reject);
  });
}

function toErrorResponse(requestId: string, message: string): ChatResponse {
  return {
    answer: '',
    citations: [],
    provider: process.env.AI_PROVIDER || 'mock',
    requestId,
    status: 'error',
    error: message
  };
}

const server = createServer(async (request, response) => {
  if (request.method === 'OPTIONS') {
    sendJson(response, 204, {});
    return;
  }

  if (request.method === 'GET' && request.url === '/api/health') {
    sendJson(response, 200, {
      ok: true,
      service: 'nextcore-projects-function-app-dev-server',
      provider: process.env.AI_PROVIDER || 'mock'
    });
    return;
  }

  if (request.method === 'POST' && (request.url === '/api/chat' || request.url === '/api/ask')) {
    const requestId = randomUUID();

    try {
      const bodyText = await readBody(request);
      const body = JSON.parse(bodyText || '{}') as Partial<ChatRequest>;
      const question = body.question?.trim();

      if (!question) {
        sendJson(response, 400, toErrorResponse(requestId, 'Question is required.'));
        return;
      }

      const selectedDocument = body.selectedDocument;
      let extractedDocumentText = '';

      if (selectedDocument) {
        if (selectedDocument.fileType?.toLowerCase() !== 'pdf') {
          sendJson(response, 400, toErrorResponse(requestId, 'This proof of concept currently supports selected PDF files only.'));
          return;
        }

        const documentBytes = Buffer.from(selectedDocument.contentBase64 || '', 'base64');
        if (documentBytes.length === 0) {
          sendJson(response, 400, toErrorResponse(requestId, 'Selected PDF content is empty.'));
          return;
        }
        if (documentBytes.length > MAX_SELECTED_PDF_BYTES) {
          sendJson(response, 413, toErrorResponse(requestId, 'Selected PDF exceeds the 4 MB proof-of-concept limit.'));
          return;
        }

        const extraction = await extractPdfText(documentBytes);
        if (extraction.requiresOcr) {
          sendJson(response, 422, toErrorResponse(requestId, 'No selectable text was found in the selected PDF. OCR is required before Ask AI can analyze it.'));
          return;
        }
        extractedDocumentText = extraction.text.slice(0, MAX_DOCUMENT_TEXT_CHARS);
      }

      const selectedFiles = (body.selectedFiles || []).map((file) => {
        if (extractedDocumentText && selectedDocument && file.name === selectedDocument.name) {
          return { ...file, snippet: extractedDocumentText.slice(0, 800) };
        }
        return file;
      });

      const result = await answerQuestion({
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
      }, requestId);

      sendJson(response, 200, result);
    } catch (error) {
      sendJson(response, 500, toErrorResponse(requestId, (error as Error).message || 'Unexpected backend error.'));
    }
    return;
  }

  sendJson(response, 404, { error: 'Not found' });
});

server.listen(port, () => {
  console.log(`NextCore Function App dev server running at http://localhost:${port}`);
  console.log('Available endpoints: GET /api/health, POST /api/chat, POST /api/ask');
});
