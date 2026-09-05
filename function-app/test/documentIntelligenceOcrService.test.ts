import assert from 'node:assert/strict';
import test from 'node:test';
import {
  getDocumentIntelligenceOcrConfiguration,
  runDocumentIntelligenceOcr
} from '../src/services/documentIntelligenceOcrService';

test('getDocumentIntelligenceOcrConfiguration returns a normalized explicit OCR configuration', () => {
  const configuration = getDocumentIntelligenceOcrConfiguration({
    DOCUMENT_INTELLIGENCE_ENDPOINT: 'https://contoso.cognitiveservices.azure.com/',
    DOCUMENT_INTELLIGENCE_API_KEY: 'test-key',
    OCR_MAX_PDF_BYTES: '4194304',
    OCR_MAX_PAGES: '15'
  });

  assert.equal(configuration.endpoint, 'https://contoso.cognitiveservices.azure.com');
  assert.equal(configuration.apiKey, 'test-key');
  assert.equal(configuration.maxPdfBytes, 4 * 1024 * 1024);
  assert.equal(configuration.maxPages, 15);
});

test('getDocumentIntelligenceOcrConfiguration rejects a missing endpoint or API key', () => {
  assert.throws(
    () => getDocumentIntelligenceOcrConfiguration({ DOCUMENT_INTELLIGENCE_API_KEY: 'test-key' }),
    /DOCUMENT_INTELLIGENCE_ENDPOINT/i
  );
  assert.throws(
    () => getDocumentIntelligenceOcrConfiguration({ DOCUMENT_INTELLIGENCE_ENDPOINT: 'https://contoso.cognitiveservices.azure.com' }),
    /DOCUMENT_INTELLIGENCE_API_KEY/i
  );
});

test('runDocumentIntelligenceOcr posts PDF bytes then returns page text from a succeeded operation', async () => {
  const requests: Array<{ url: string; init?: RequestInit }> = [];
  const responses = [
    new Response(null, {
      status: 202,
      headers: { 'operation-location': 'https://contoso.cognitiveservices.azure.com/operations/123' }
    }),
    new Response(JSON.stringify({ status: 'running' }), { status: 200 }),
    new Response(JSON.stringify({
      status: 'succeeded',
      analyzeResult: {
        pages: [
          { pageNumber: 1, lines: [{ content: 'First line' }, { content: 'Second line' }] },
          { pageNumber: 2, lines: [{ content: 'Third line' }] }
        ]
      }
    }), { status: 200 })
  ];

  const result = await runDocumentIntelligenceOcr(
    Buffer.from('%PDF-1.7'),
    {
      endpoint: 'https://contoso.cognitiveservices.azure.com',
      apiKey: 'test-key',
      maxPdfBytes: 1024,
      maxPages: 2
    },
    {
      fetch: async (url: string | URL | Request, init?: RequestInit) => {
        requests.push({ url: String(url), init });
        const response = responses.shift();
        if (!response) throw new Error('Unexpected request');
        return response;
      },
      pollIntervalMs: 0
    }
  );

  assert.equal(requests.length, 3);
  assert.equal(requests[0].url, 'https://contoso.cognitiveservices.azure.com/documentintelligence/documentModels/prebuilt-read:analyze?api-version=2024-11-30');
  assert.equal(requests[0].init?.method, 'POST');
  assert.equal(new Headers(requests[0].init?.headers).get('Ocp-Apim-Subscription-Key'), 'test-key');
  assert.equal(new Headers(requests[0].init?.headers).get('Content-Type'), 'application/pdf');
  assert.equal(Buffer.from(requests[0].init?.body as ArrayBuffer).toString('utf8'), '%PDF-1.7');
  assert.deepEqual(result, {
    text: 'First line\nSecond line\n\nThird line',
    pages: [
      { pageNumber: 1, text: 'First line\nSecond line' },
      { pageNumber: 2, text: 'Third line' }
    ]
  });
});

test('runDocumentIntelligenceOcr stops polling after the configured bounded attempts', async () => {
  let requestCount = 0;
  const fetch = async (): Promise<Response> => {
    requestCount += 1;
    if (requestCount === 1) {
      return new Response(null, { status: 202, headers: { 'operation-location': 'https://contoso/operations/123' } });
    }
    return new Response(JSON.stringify({ status: 'running' }), { status: 200 });
  };

  await assert.rejects(
    runDocumentIntelligenceOcr(Buffer.from('%PDF'), {
      endpoint: 'https://contoso', apiKey: 'test-key', maxPdfBytes: 1024, maxPages: 2
    }, { fetch, maxPollAttempts: 2, pollIntervalMs: 0 }),
    /did not complete after 2 polling attempts/i
  );
  assert.equal(requestCount, 3);
});

test('runDocumentIntelligenceOcr returns a redacted error when the operation fails', async () => {
  const responses = [
    new Response(null, { status: 202, headers: { 'operation-location': 'https://contoso/operations/123' } }),
    new Response(JSON.stringify({ status: 'failed', error: { message: 'secret service diagnostic' } }), { status: 200 })
  ];
  const fetch = async (): Promise<Response> => {
    const response = responses.shift();
    if (!response) throw new Error('Unexpected request');
    return response;
  };

  await assert.rejects(
    runDocumentIntelligenceOcr(Buffer.from('%PDF'), {
      endpoint: 'https://contoso', apiKey: 'test-key', maxPdfBytes: 1024, maxPages: 2
    }, { fetch, pollIntervalMs: 0 }),
    (error: Error) => /operation failed/i.test(error.message)
      && !/test-key|secret service diagnostic/i.test(error.message)
  );
});

test('runDocumentIntelligenceOcr enforces PDF size before network submission', async () => {
  let wasCalled = false;

  await assert.rejects(
    runDocumentIntelligenceOcr(Buffer.from('12345'), {
      endpoint: 'https://contoso', apiKey: 'test-key', maxPdfBytes: 4, maxPages: 2
    }, {
      fetch: async () => {
        wasCalled = true;
        return new Response(null, { status: 202 });
      }
    }),
    /maximum size of 4 bytes/i
  );
  assert.equal(wasCalled, false);
});

test('runDocumentIntelligenceOcr rejects results that exceed the configured page maximum', async () => {
  const responses = [
    new Response(null, { status: 202, headers: { 'operation-location': 'https://contoso/operations/123' } }),
    new Response(JSON.stringify({
      status: 'succeeded',
      analyzeResult: { pages: [{ pageNumber: 1 }, { pageNumber: 2 }] }
    }), { status: 200 })
  ];
  const fetch = async (): Promise<Response> => {
    const response = responses.shift();
    if (!response) throw new Error('Unexpected request');
    return response;
  };

  await assert.rejects(
    runDocumentIntelligenceOcr(Buffer.from('%PDF'), {
      endpoint: 'https://contoso', apiKey: 'test-key', maxPdfBytes: 1024, maxPages: 1
    }, { fetch, pollIntervalMs: 0 }),
    /maximum of 1 pages/i
  );
});
