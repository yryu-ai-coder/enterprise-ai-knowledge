export interface DocumentIntelligenceOcrConfiguration {
  endpoint: string;
  apiKey: string;
  maxPdfBytes: number;
  maxPages: number;
}

export interface DocumentIntelligenceOcrPage {
  pageNumber: number;
  text: string;
}

export interface DocumentIntelligenceOcrResult {
  text: string;
  pages: DocumentIntelligenceOcrPage[];
}

export interface DocumentIntelligenceOcrDependencies {
  fetch?: typeof globalThis.fetch;
  maxPollAttempts?: number;
  pollIntervalMs?: number;
  sleep?: (milliseconds: number) => Promise<void>;
}

interface DocumentIntelligenceOperation {
  status?: string;
  analyzeResult?: {
    pages?: Array<{
      pageNumber?: number;
      lines?: Array<{ content?: string }>;
    }>;
  };
}

const DEFAULT_MAX_PDF_BYTES = 4 * 1024 * 1024;
const DEFAULT_MAX_PAGES = 15;
const DEFAULT_MAX_POLL_ATTEMPTS = 30;
const DEFAULT_POLL_INTERVAL_MS = 500;
const ANALYZE_PATH = '/documentintelligence/documentModels/prebuilt-read:analyze?api-version=2024-11-30';

function getPositiveInteger(value: string | undefined, fallback: number, name: string): number {
  if (!value?.trim()) {
    return fallback;
  }

  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 1) {
    throw new Error(`${name} must be a positive integer.`);
  }
  return parsed;
}

function getBoundedPositiveInteger(value: number | undefined, fallback: number, name: string): number {
  if (value === undefined) {
    return fallback;
  }
  if (!Number.isInteger(value) || value < 1) {
    throw new Error(`${name} must be a positive integer.`);
  }
  return value;
}

function defaultSleep(milliseconds: number): Promise<void> {
  return new Promise((resolve) => setTimeout(resolve, milliseconds));
}

function getAnalyzeUrl(endpoint: string): string {
  return `${endpoint}${ANALYZE_PATH}`;
}

function extractOcrResult(operation: DocumentIntelligenceOperation, maxPages: number): DocumentIntelligenceOcrResult {
  const sourcePages = operation.analyzeResult?.pages;
  if (!sourcePages) {
    throw new Error('OCR operation succeeded without an analysis result.');
  }
  if (sourcePages.length > maxPages) {
    throw new Error(`OCR result exceeds the configured maximum of ${maxPages} pages.`);
  }

  const pages = sourcePages.map((page, index) => ({
    pageNumber: page.pageNumber ?? index + 1,
    text: (page.lines ?? [])
      .map((line) => line.content?.trim())
      .filter((content): content is string => Boolean(content))
      .join('\n')
  }));

  return {
    text: pages.map((page) => page.text).filter(Boolean).join('\n\n'),
    pages
  };
}

async function readOperation(response: Response, stage: 'submission' | 'polling'): Promise<DocumentIntelligenceOperation> {
  if (!response.ok) {
    throw new Error(`Document Intelligence OCR ${stage} request failed with HTTP ${response.status}.`);
  }

  try {
    return await response.json() as DocumentIntelligenceOperation;
  } catch {
    throw new Error(`Document Intelligence OCR ${stage} response was invalid.`);
  }
}

export function getDocumentIntelligenceOcrConfiguration(
  environment: Record<string, string | undefined> = process.env
): DocumentIntelligenceOcrConfiguration {
  const endpoint = environment.DOCUMENT_INTELLIGENCE_ENDPOINT?.trim().replace(/\/$/, '');
  if (!endpoint) {
    throw new Error('DOCUMENT_INTELLIGENCE_ENDPOINT is required for OCR processing.');
  }

  const apiKey = environment.DOCUMENT_INTELLIGENCE_API_KEY?.trim();
  if (!apiKey) {
    throw new Error('DOCUMENT_INTELLIGENCE_API_KEY is required for OCR processing.');
  }

  return {
    endpoint,
    apiKey,
    maxPdfBytes: getPositiveInteger(environment.OCR_MAX_PDF_BYTES, DEFAULT_MAX_PDF_BYTES, 'OCR_MAX_PDF_BYTES'),
    maxPages: getPositiveInteger(environment.OCR_MAX_PAGES, DEFAULT_MAX_PAGES, 'OCR_MAX_PAGES')
  };
}

export async function runDocumentIntelligenceOcr(
  pdfBytes: Uint8Array,
  configuration: DocumentIntelligenceOcrConfiguration,
  dependencies: DocumentIntelligenceOcrDependencies = {}
): Promise<DocumentIntelligenceOcrResult> {
  if (pdfBytes.byteLength > configuration.maxPdfBytes) {
    throw new Error(`PDF exceeds the configured maximum size of ${configuration.maxPdfBytes} bytes.`);
  }

  const fetchImplementation = dependencies.fetch ?? globalThis.fetch;
  const maxPollAttempts = getBoundedPositiveInteger(
    dependencies.maxPollAttempts,
    DEFAULT_MAX_POLL_ATTEMPTS,
    'maxPollAttempts'
  );
  const pollIntervalMs = dependencies.pollIntervalMs ?? DEFAULT_POLL_INTERVAL_MS;
  if (!Number.isFinite(pollIntervalMs) || pollIntervalMs < 0) {
    throw new Error('pollIntervalMs must be a non-negative number.');
  }
  const sleep = dependencies.sleep ?? defaultSleep;

  const pdfBody = new ArrayBuffer(pdfBytes.byteLength);
  new Uint8Array(pdfBody).set(pdfBytes);
  const submission = await fetchImplementation(getAnalyzeUrl(configuration.endpoint), {
    method: 'POST',
    headers: {
      'Content-Type': 'application/pdf',
      'Ocp-Apim-Subscription-Key': configuration.apiKey
    },
    body: pdfBody
  });

  if (submission.status !== 202) {
    await readOperation(submission, 'submission');
    throw new Error('Document Intelligence OCR submission did not return an operation location.');
  }

  const operationLocation = submission.headers.get('operation-location');
  if (!operationLocation) {
    throw new Error('Document Intelligence OCR submission did not return an operation location.');
  }

  for (let attempt = 0; attempt < maxPollAttempts; attempt += 1) {
    if (attempt > 0 && pollIntervalMs > 0) {
      await sleep(pollIntervalMs);
    }

    const operation = await readOperation(
      await fetchImplementation(operationLocation, {
        method: 'GET',
        headers: { 'Ocp-Apim-Subscription-Key': configuration.apiKey }
      }),
      'polling'
    );

    if (operation.status === 'succeeded') {
      return extractOcrResult(operation, configuration.maxPages);
    }
    if (operation.status === 'failed') {
      throw new Error('Document Intelligence OCR operation failed.');
    }
  }

  throw new Error(`Document Intelligence OCR operation did not complete after ${maxPollAttempts} polling attempts.`);
}
