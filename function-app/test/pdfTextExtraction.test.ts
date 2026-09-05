import assert from 'node:assert/strict';
import { readFile } from 'node:fs/promises';
import path from 'node:path';
import test from 'node:test';
import { doesPdfTextRequireOcr, extractPdfText } from '../src/services/pdfTextExtraction';

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

test('doesPdfTextRequireOcr classifies a sparse title-only text layer as OCR required', () => {
  assert.equal(doesPdfTextRequireOcr("Ava's Text Messages: March 10, 2026"), true);
  assert.equal(doesPdfTextRequireOcr('This document contains enough selectable text to be meaningfully grounded without OCR. '.repeat(4)), false);
});
