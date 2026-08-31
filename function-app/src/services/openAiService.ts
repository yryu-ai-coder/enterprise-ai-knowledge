import { ChatRequest, ChatResponse, Citation } from '../models';

interface OpenAIChatResponse {
  choices?: Array<{
    message?: {
      content?: string;
    };
  }>;
  usage?: Record<string, unknown>;
}

function buildSystemPrompt(): string {
  return [
    'You are a Legal Document Library AI Assistant running behind a SharePoint SPFx web part.',
    'You help users summarize litigation documents, extract dates, identify parties/claims, and plan follow-up review work.',
    'Use only the supplied SharePoint metadata/snippets. If file content is not supplied, say the response is metadata-based only.',
    'Do not provide legal advice. Do not invent citations.'
  ].join('\n');
}

function buildUserPrompt(request: ChatRequest): string {
  return JSON.stringify({
    question: request.question,
    scenario: request.scenario || 'legal-document-library',
    mode: request.mode || 'legal-analysis',
    contextType: request.contextType || 'document-library',
    siteUrl: request.siteUrl || '',
    pageUrl: request.pageUrl || '',
    libraryName: request.libraryName || request.listTitle || 'Documents',
    folderPath: request.folderPath || '',
    selectedFiles: request.selectedFiles || [],
    selectedItems: request.selectedItems || [],
    documentSnippets: request.documentSnippets || [],
    expectedAnswerStyle: 'Concise enterprise legal-work-product style with bullets, risks, and next actions.'
  }, null, 2);
}

function buildCitations(request: ChatRequest): Citation[] {
  if (request.selectedFiles && request.selectedFiles.length > 0) {
    return request.selectedFiles.slice(0, 5).map(file => ({
      title: file.name,
      url: file.url || request.siteUrl || '',
      snippet: file.snippet || `Metadata context from ${file.libraryTitle || request.libraryName || 'SharePoint document library'}.`
    }));
  }

  return request.pageUrl ? [{
    title: request.pageTitle || request.libraryName || 'SharePoint context',
    url: request.pageUrl,
    snippet: 'Page/library context supplied by the SPFx web part.'
  }] : [];
}

export async function answerWithOpenAI(request: ChatRequest, requestId: string): Promise<ChatResponse> {
  const apiKey = process.env.OPENAI_API_KEY;
  const model = process.env.OPENAI_MODEL || 'gpt-4o-mini';
  const baseUrl = (process.env.OPENAI_BASE_URL || 'https://api.openai.com/v1').replace(/\/$/, '');
  const temperature = Number(process.env.AI_TEMPERATURE || 0.2);
  const maxTokens = Number(process.env.AI_MAX_TOKENS || 900);

  if (!apiKey) {
    throw new Error('OpenAI configuration is missing. Set AI_PROVIDER=openai and OPENAI_API_KEY on the Function App.');
  }

  const response = await fetch(`${baseUrl}/chat/completions`, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      Authorization: `Bearer ${apiKey}`
    },
    body: JSON.stringify({
      model,
      messages: [
        { role: 'system', content: buildSystemPrompt() },
        { role: 'user', content: buildUserPrompt(request) }
      ],
      temperature,
      max_tokens: maxTokens
    })
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`OpenAI returned ${response.status}: ${errorText}`);
  }

  const payload = await response.json() as OpenAIChatResponse;

  return {
    answer: payload.choices?.[0]?.message?.content || 'No answer returned from OpenAI.',
    citations: buildCitations(request),
    provider: 'openai',
    requestId,
    status: 'success',
    suggestedActions: [
      'Summarize selected litigation documents',
      'Extract key dates and deadlines',
      'Identify parties and claims',
      'Find missing evidence or open questions',
      'Draft a case timeline'
    ],
    metadata: {
      scenario: request.scenario || 'legal-document-library',
      mode: request.mode || 'legal-analysis',
      contextType: request.contextType || 'document-library',
      libraryName: request.libraryName || request.listTitle || '',
      folderPath: request.folderPath || '',
      selectedFileCount: request.selectedFiles?.length || 0,
      model,
      usage: payload.usage || {}
    }
  };
}
