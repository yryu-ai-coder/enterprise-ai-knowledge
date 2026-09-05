import { ChatRequest, ChatResponse, Citation } from '../models';

interface AzureOpenAIChatResponse {
  choices?: Array<{
    message?: {
      content?: string;
    };
  }>;
  usage?: Record<string, unknown>;
}

interface AzureOpenAIResponsesResponse {
  output_text?: string;
  output?: Array<{
    type?: string;
    content?: Array<{
      type?: string;
      text?: string;
    }>;
  }>;
  usage?: Record<string, unknown>;
}

function buildSystemPrompt(): string {
  return [
    'You are a Legal Document Library AI Assistant running behind a SharePoint SPFx web part.',
    'You help with litigation document organization, summaries, key dates, parties, claims, evidence gaps, and next-step planning.',
    'Important safety rules:',
    '- You are not a lawyer and do not provide legal advice.',
    '- Answer in plain, natural language suitable for a compact chat panel. Do not use Markdown syntax: no **, __, # headings, tables, or technical labels.',
    '- Lead with a direct answer. Add at most three short bullet points only when they materially help.',
    '- End after the answer. Do not add an unsolicited offer to help further, a follow-up question, next steps, or a conclusion unless the user asks for one.',
    '- Do not mention Azure AI Search, Graph, metadata, snippets, or implementation details unless the user explicitly asks how the system works.',
    '- Never invent citations. Cite only selected files or supplied context.',
    '- If the supplied evidence is insufficient, say what information is missing in user-facing language without describing the system architecture.',
    '- For spreadsheet or budget questions, use only explicitly supplied cell values. Align a value with its column header; a YEAR/total value is never a monthly value. A blank actual-month cell means that month has no entered actual and must not be reported as an overrun. State the planned amount, actual amount, and calculated difference before naming a largest variance. If the retrieved worksheet evidence cannot support the calculation, say so instead of estimating.'
  ].join('\n');
}

function buildUserPrompt(request: ChatRequest): string {
  const selectedFiles = (request.selectedFiles || []).map((file, index) => ({
    index: index + 1,
    name: file.name,
    url: file.url,
    fileType: file.fileType,
    libraryTitle: file.libraryTitle,
    lastModified: file.lastModified,
    snippet: file.snippet
  }));

  const selectedItems = request.selectedItems || [];

  return JSON.stringify({
    instruction: request.question,
    scenario: request.scenario || 'legal-document-library',
    mode: request.mode || 'legal-analysis',
    contextType: request.contextType || 'document-library',
    siteUrl: request.siteUrl || '',
    pageUrl: request.pageUrl || '',
    pageTitle: request.pageTitle || '',
    libraryName: request.libraryName || request.listTitle || 'Documents',
    folderPath: request.folderPath || '',
    knowledgeScope: request.knowledgeScope || 'legal-document-library',
    selectedFiles,
    selectedItems,
    documentSnippets: request.documentSnippets || [],
    responseFormat: {
      style: 'Plain natural-language chat response. Direct answer first; short paragraphs; at most three bullets when useful.',
      prohibited: 'Do not use Markdown, bold markers, headings, tables, technical implementation commentary, risk boilerplate, recommended-action boilerplate, or an unsolicited offer to help further.',
      citationGuidance: 'State the relevant document fact naturally; the interface presents sources separately.'
    }
  }, null, 2);
}

function buildCitations(request: ChatRequest): Citation[] {
  const selectedFiles = request.selectedFiles || [];

  if (selectedFiles.length > 0) {
    return selectedFiles.slice(0, 5).map(file => ({
      title: file.name,
      url: file.url || request.siteUrl || '',
      snippet: file.snippet || `Metadata context from ${file.libraryTitle || request.libraryName || 'SharePoint document library'}.`
    }));
  }

  if (request.pageUrl) {
    return [{
      title: request.pageTitle || request.libraryName || 'SharePoint context',
      url: request.pageUrl,
      snippet: 'Page/library context supplied by the SPFx web part.'
    }];
  }

  return [];
}

function endpointLooksLikeResponsesApi(endpoint: string): boolean {
  return /\/openai\/v\d+\/responses\/?$/i.test(endpoint.trim());
}

function normalizeEndpoint(endpoint: string): string {
  return endpoint.replace(/\/$/, '');
}

function getResponsesUrl(endpoint: string): string {
  const normalized = normalizeEndpoint(endpoint);

  if (endpointLooksLikeResponsesApi(normalized)) {
    return normalized;
  }

  // Supports Foundry resource endpoints such as:
  // https://<resource>.services.ai.azure.com
  // https://<resource>.openai.azure.com
  if (/\/openai\/v\d+$/i.test(normalized)) {
    return `${normalized}/responses`;
  }

  return `${normalized}/openai/v1/responses`;
}

function getChatCompletionsUrl(endpoint: string, deployment: string, apiVersion: string): string {
  const normalized = normalizeEndpoint(endpoint);

  if (endpointLooksLikeResponsesApi(normalized)) {
    throw new Error('The configured AZURE_OPENAI_ENDPOINT is a Responses API endpoint. Set AZURE_OPENAI_API_STYLE=responses or use the Azure OpenAI resource root endpoint for chat completions.');
  }

  return `${normalized}/openai/deployments/${encodeURIComponent(deployment)}/chat/completions?api-version=${encodeURIComponent(apiVersion)}`;
}

function extractResponseText(payload: AzureOpenAIResponsesResponse): string {
  if (payload.output_text) {
    return payload.output_text;
  }

  const textParts = (payload.output || [])
    .flatMap(item => item.content || [])
    .map(content => content.text)
    .filter((text): text is string => Boolean(text));

  return textParts.join('\n').trim() || 'No answer returned from Azure OpenAI Responses API.';
}

function buildSuccessResponse(
  answer: string,
  request: ChatRequest,
  requestId: string,
  deployment: string,
  apiVersion: string,
  apiStyle: string,
  usage: Record<string, unknown> | undefined
): ChatResponse {
  return {
    answer,
    citations: buildCitations(request),
    provider: 'azure-openai',
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
      modelDeployment: deployment,
      apiVersion,
      apiStyle,
      usage: usage || {}
    }
  };
}

async function answerWithResponsesApi(
  endpoint: string,
  apiKey: string,
  deployment: string,
  apiVersion: string,
  temperature: number,
  maxTokens: number,
  request: ChatRequest,
  requestId: string
): Promise<ChatResponse> {
  const url = getResponsesUrl(endpoint);

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'api-key': apiKey
    },
    body: JSON.stringify({
      model: deployment,
      instructions: buildSystemPrompt(),
      input: buildUserPrompt(request),
      temperature,
      max_output_tokens: maxTokens
    })
  });

  if (!response.ok) {
    const errorText = await response.text();
    throw new Error(`Azure OpenAI Responses API returned ${response.status}: ${errorText}`);
  }

  const payload = await response.json() as AzureOpenAIResponsesResponse;
  return buildSuccessResponse(extractResponseText(payload), request, requestId, deployment, apiVersion, 'responses', payload.usage);
}

async function answerWithChatCompletionsApi(
  endpoint: string,
  apiKey: string,
  deployment: string,
  apiVersion: string,
  temperature: number,
  maxTokens: number,
  request: ChatRequest,
  requestId: string
): Promise<ChatResponse> {
  const url = getChatCompletionsUrl(endpoint, deployment, apiVersion);

  const response = await fetch(url, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'api-key': apiKey
    },
    body: JSON.stringify({
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
    throw new Error(`Azure OpenAI Chat Completions API returned ${response.status}: ${errorText}`);
  }

  const payload = await response.json() as AzureOpenAIChatResponse;
  const answer = payload.choices?.[0]?.message?.content || 'No answer returned from Azure OpenAI Chat Completions API.';
  return buildSuccessResponse(answer, request, requestId, deployment, apiVersion, 'chat-completions', payload.usage);
}

export async function answerWithAzureOpenAI(request: ChatRequest, requestId: string): Promise<ChatResponse> {
  const endpoint = process.env.AZURE_OPENAI_ENDPOINT;
  const apiKey = process.env.AZURE_OPENAI_API_KEY;
  const deployment = process.env.AZURE_OPENAI_DEPLOYMENT;
  const apiVersion = process.env.AZURE_OPENAI_API_VERSION || '2024-06-01';
  const apiStyle = (process.env.AZURE_OPENAI_API_STYLE || '').toLowerCase();
  const temperature = Number(process.env.AI_TEMPERATURE || 0.2);
  const maxTokens = Number(process.env.AI_MAX_TOKENS || 900);

  if (!endpoint || !apiKey || !deployment) {
    throw new Error('Azure OpenAI configuration is missing. Set AI_PROVIDER=azure-openai, AZURE_OPENAI_ENDPOINT, AZURE_OPENAI_API_KEY, and AZURE_OPENAI_DEPLOYMENT on the Function App.');
  }

  if (apiStyle === 'responses' || endpointLooksLikeResponsesApi(endpoint)) {
    return answerWithResponsesApi(endpoint, apiKey, deployment, apiVersion, temperature, maxTokens, request, requestId);
  }

  return answerWithChatCompletionsApi(endpoint, apiKey, deployment, apiVersion, temperature, maxTokens, request, requestId);
}
