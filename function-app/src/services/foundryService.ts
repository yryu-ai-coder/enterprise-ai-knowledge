import { ChatRequest, ChatResponse } from '../models';

export async function answerWithFoundry(request: ChatRequest, requestId: string): Promise<ChatResponse> {
  const endpoint = process.env.FOUNDRY_AGENT_ENDPOINT;
  const apiKey = process.env.FOUNDRY_AGENT_API_KEY;

  if (!endpoint || !apiKey) {
    throw new Error('Microsoft Foundry Agent configuration is missing.');
  }

  // Placeholder adapter. Replace payload shape with the selected Foundry Agent runtime endpoint contract.
  const response = await fetch(endpoint, {
    method: 'POST',
    headers: {
      'Content-Type': 'application/json',
      'api-key': apiKey
    },
    body: JSON.stringify({
      input: request.question,
      context: {
        siteUrl: request.siteUrl,
        pageContext: request.pageContext,
        knowledgeScope: request.knowledgeScope
      }
    })
  });

  if (!response.ok) {
    throw new Error(`Foundry endpoint returned ${response.status}`);
  }

  const payload = await response.json() as { answer?: string; output?: string; citations?: ChatResponse['citations'] };

  return {
    answer: payload.answer || payload.output || 'No answer returned from Foundry.',
    citations: payload.citations || [],
    provider: 'foundry',
    requestId,
    status: 'success'
  };
}
