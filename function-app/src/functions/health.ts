import { app, HttpRequest, HttpResponseInit } from '@azure/functions';

/**
 * Anonymous readiness endpoint only. It deliberately exposes no secrets or
 * configuration values beyond the active provider name.
 */
app.http('health', {
  methods: ['GET'],
  authLevel: 'anonymous',
  route: 'health',
  handler: async (_request: HttpRequest): Promise<HttpResponseInit> => ({
    status: 200,
    jsonBody: {
      ok: true,
      service: 'nextcore-ai-function',
      provider: process.env.AI_PROVIDER || 'mock'
    }
  })
});
