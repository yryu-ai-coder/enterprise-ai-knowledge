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
    'Answer in plain, natural language for a compact chat panel. Do not use Markdown syntax, bold markers, headings, tables, or technical implementation commentary.',
    'Lead with a direct answer. Use at most three short bullets only when they materially help.',
    'End after the answer. Do not add an unsolicited offer to help further, a follow-up question, next steps, or a conclusion unless the user asks for one.',
    'Do not mention Azure AI Search, Graph, metadata, or snippets unless the user explicitly asks how the system works.',
    'If supplied evidence is insufficient, state what information is missing in user-facing language. Do not provide legal advice or invent citations.',
    'For spreadsheet or budget questions, use only explicitly supplied cell values. Align a value with its column header; a YEAR/total value is never a monthly value. A blank actual-month cell means that month has no entered actual and must not be reported as an overrun. State the planned amount, actual amount, and calculated difference before naming a largest variance. If the retrieved worksheet evidence cannot support the calculation, say so instead of estimating.'
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
    expectedAnswerStyle: 'Concise plain-language answer, direct answer first, with at most three useful bullets. Do not add risks, next actions, or an offer to help further unless requested.'
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
