import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { extractPdfText } from '../src/services/pdfTextExtraction';

test('extractPdfText returns text from a text-layer PDF', async () => {
  const fixturePath = path.resolve(process.cwd(), 'test/fixtures/text-layer-sample.pdf');
  const fileBytes = await readFile(fixturePath);

  const result = await extractPdfText(fileBytes);

  assert.match(result.text, /notice deadline is September 15, 2026/i);
  assert.equal(result.requiresOcr, false);
});

test('extractPdfText flags a PDF with no extractable text as requiring OCR', async () => {
  const fixturePath = path.resolve(process.cwd(), 'test/fixtures/blank-text-layer.pdf');
  const fileBytes = await readFile(fixturePath);

  const result = await extractPdfText(fileBytes);

  assert.equal(result.text, '');
  assert.equal(result.requiresOcr, true);
});
