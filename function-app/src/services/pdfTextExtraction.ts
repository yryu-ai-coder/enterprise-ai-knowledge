interface PdfJsTextItem {
  str?: string;
}

interface PdfJsPage {
  getTextContent(): Promise<{ items: PdfJsTextItem[] }>;
}

interface PdfJsDocument {
  numPages: number;
  getPage(pageNumber: number): Promise<PdfJsPage>;
  destroy(): Promise<void>;
}

interface PdfJsModule {
  getDocument(source: { data: Uint8Array; useWorkerFetch: boolean; isEvalSupported: boolean }): { promise: Promise<PdfJsDocument> };
}

const dynamicImport = new Function('specifier', 'return import(specifier);') as (specifier: string) => Promise<PdfJsModule>;
const MIN_USABLE_TEXT_CHARS = 50;

async function getPdfJs(): Promise<PdfJsModule> {
  // Keep the ESM parser lazy so Function route discovery never loads PDF code.
  return dynamicImport('pdfjs-dist/legacy/build/pdf.mjs');
}

export interface PdfTextExtractionResult {
  text: string;
  requiresOcr: boolean;
}

export function doesPdfTextRequireOcr(text: string): boolean {
  return text.trim().length < MIN_USABLE_TEXT_CHARS;
}

export async function extractPdfText(fileBytes: Buffer): Promise<PdfTextExtractionResult> {
  let document: PdfJsDocument | undefined;

  try {
    const pdfjs = await getPdfJs();
    document = await pdfjs.getDocument({
      data: new Uint8Array(fileBytes),
      useWorkerFetch: false,
      isEvalSupported: false
    }).promise;

    const pages: string[] = [];
    for (let pageNumber = 1; pageNumber <= document.numPages; pageNumber += 1) {
      const page = await document.getPage(pageNumber);
      const content = await page.getTextContent();
      pages.push(content.items.map(item => item.str || '').join(' '));
    }

    const text = pages.join('\n').replace(/\s+\n/g, '\n').replace(/\r/g, '').trim();
    return { text, requiresOcr: doesPdfTextRequireOcr(text) };
  } catch (error) {
    throw new Error(`PDF text extraction failed: ${(error as Error).message || 'Unknown parser error.'}`);
  } finally {
    await document?.destroy();
  }
}
