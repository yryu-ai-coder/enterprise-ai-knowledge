import { ChatRequest, ChatResponse, Citation } from '../models';

function buildCitations(request: ChatRequest): Citation[] {
  if (request.selectedFiles && request.selectedFiles.length > 0) {
    return request.selectedFiles.slice(0, 3).map(file => ({
      title: file.name,
      url: file.url || request.siteUrl || 'https://sharepoint.local/',
      snippet: file.snippet || `Mock citation from ${file.libraryTitle || request.libraryName || 'Documents'} metadata. Replace with Graph/Azure AI Search grounding later.`
    }));
  }

  if (request.contextType === 'sharepoint-list' && request.selectedItems && request.selectedItems.length > 0) {
    return request.selectedItems.slice(0, 3).map(item => ({
      title: item.title || `List item ${item.id}`,
      url: request.siteUrl || 'https://sharepoint.local/',
      snippet: `Mock citation from ${item.listTitle || request.listTitle || 'SharePoint List'} item ${item.id}.`
    }));
  }

  if (request.contextType === 'site-page') {
    return [{
      title: request.pageTitle || 'Current SharePoint Site Page',
      url: request.pageUrl || request.siteUrl || 'https://sharepoint.local/',
      snippet: 'Mock citation from current page context. Replace with Site Pages canvas content extraction later.'
    }];
  }

  return [{
    title: 'Legal Document Library',
    url: request.siteUrl || 'https://sharepoint.local/',
    snippet: 'Mock citation placeholder for legal document library grounding.'
  }];
}

function buildLegalAnswer(request: ChatRequest): string {
  const files = request.selectedFiles?.map(file => file.name).join(', ') || 'selected legal documents';
  const scope = request.knowledgeScope || request.scenario || 'legal-document-library';

  return [
    `Legal Document Library mock analysis for: "${request.question}"`,
    '',
    `Context type: ${request.contextType || 'document-library'}`,
    `Scenario: ${request.scenario || scope}`,
    `Library: ${request.libraryName || request.listTitle || 'Documents'}`,
    `Selected material: ${files}`,
    '',
    'What the real AI step should do next:',
    '1. Read selected litigation documents from SharePoint through Microsoft Graph or an Azure AI Search index.',
    '2. Produce a concise legal-work-product style summary: issues, parties, dates, obligations, evidence references, and open questions.',
    '3. Return source-linked citations so every statement can be traced back to a document.',
    '',
    'POC note: this is a mock response only. Do not use it as legal advice.'
  ].join('\n');
}

function buildGenericAnswer(request: ChatRequest): string {
  return `Mock response for: "${request.question}"

This confirms the reusable backend can serve ${request.contextType || 'SharePoint'} scenarios. Next step is switching AI_PROVIDER to azure-openai or foundry.`;
}

export async function answerWithMockProvider(request: ChatRequest, requestId: string): Promise<ChatResponse> {
  const isLegalScenario = (request.scenario || request.knowledgeScope || '').includes('legal');

  return {
    answer: isLegalScenario ? buildLegalAnswer(request) : buildGenericAnswer(request),
    citations: buildCitations(request),
    provider: 'mock',
    requestId,
    status: 'success',
    suggestedActions: isLegalScenario ? [
      'Summarize selected litigation documents',
      'Extract key dates and deadlines',
      'Identify parties and claims',
      'Find missing evidence or open questions',
      'Draft a case timeline'
    ] : [
      'Summarize selected content',
      'Extract key dates and owners',
      'Find risks or missing information',
      'Draft a follow-up message'
    ],
    metadata: {
      scenario: request.scenario || 'legal-document-library',
      mode: request.mode || 'ask',
      contextType: request.contextType || 'document-library',
      siteUrl: request.siteUrl || '',
      libraryName: request.libraryName || request.listTitle || '',
      selectedFileCount: request.selectedFiles?.length || 0,
      selectedItemCount: request.selectedItems?.length || 0
    }
  };
}
