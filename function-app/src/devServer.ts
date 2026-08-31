import { createServer, IncomingMessage, ServerResponse } from 'http';
import { existsSync, readFileSync } from 'fs';
import { join } from 'path';
import { randomUUID } from 'crypto';
import { answerQuestion } from './services/aiProvider';
import { ChatRequest, ChatResponse } from './models';

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
      if (body.length > 1024 * 1024) {
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
        selectedFiles: body.selectedFiles || [],
        selectedItems: body.selectedItems || [],
        documentSnippets: body.documentSnippets || [],
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
