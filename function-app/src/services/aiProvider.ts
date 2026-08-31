import { ChatRequest, ChatResponse } from '../models';
import { answerWithAzureOpenAI } from './azureOpenAiService';
import { answerWithFoundry } from './foundryService';
import { answerWithMockProvider } from './mockProvider';
import { answerWithOpenAI } from './openAiService';

export async function answerQuestion(request: ChatRequest, requestId: string): Promise<ChatResponse> {
  const provider = (process.env.AI_PROVIDER || 'mock').toLowerCase();

  if (provider === 'azure-openai' || provider === 'azure') {
    return answerWithAzureOpenAI(request, requestId);
  }

  if (provider === 'openai') {
    return answerWithOpenAI(request, requestId);
  }

  if (provider === 'foundry') {
    return answerWithFoundry(request, requestId);
  }

  return answerWithMockProvider(request, requestId);
}
