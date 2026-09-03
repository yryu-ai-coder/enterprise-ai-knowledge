import assert from 'node:assert/strict';
import test from 'node:test';
import { buildLibraryChunks, chunkDocumentText, getSharePointIngestionConfiguration } from '../src/services/ragIngestionService';

test('chunkDocumentText creates bounded overlapping chunks without losing source text', () => {
  const text = 'alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima';
  const chunks = chunkDocumentText(text, { maxChars: 24, overlapChars: 7 });

  assert.ok(chunks.length >= 3);
  assert.ok(chunks.every(chunk => chunk.length <= 24));
  for (const word of text.split(' ')) {
    assert.equal(chunks.some(chunk => chunk.split(' ').includes(word)), true);
  }
  assert.equal(chunks[1].startsWith('charlie'), true);
});

test('buildLibraryChunks assigns stable Search-safe ids and preserves file citations', () => {
  const chunks = buildLibraryChunks({
    driveItemId: '01ABC:example',
    name: 'consultation.pdf',
    documentUrl: 'https://youngryu.sharepoint.com/sites/enterprise-ai-knowledge/Litigation%20Documents/consultation.pdf',
    folderPath: '/Litigation Documents',
    fileType: 'pdf',
    lastModified: '2026-09-03T12:00:00Z',
    text: 'First part of the source document. Second part of the source document.',
    chunkOptions: { maxChars: 35, overlapChars: 0 }
  });

  assert.equal(chunks.length, 2);
  assert.match(chunks[0].id, /^[A-Za-z0-9_-]+$/);
  assert.notEqual(chunks[0].id, chunks[1].id);
  assert.equal(chunks[0].documentName, 'consultation.pdf');
  assert.equal(chunks[0].documentUrl.includes('consultation.pdf'), true);
  assert.equal(chunks[0].chunkOrdinal, 0);
  assert.equal(chunks[1].chunkOrdinal, 1);
});

test('getSharePointIngestionConfiguration uses a single explicit site and library', () => {
  const configuration = getSharePointIngestionConfiguration({
    RAG_SHAREPOINT_SITE_URL: 'https://youngryu.sharepoint.com/sites/enterprise-ai-knowledge/',
    RAG_SHAREPOINT_LIBRARY_NAME: 'Litigation Documents'
  });

  assert.equal(configuration.siteUrl, 'https://youngryu.sharepoint.com/sites/enterprise-ai-knowledge');
  assert.equal(configuration.libraryName, 'Litigation Documents');
});

test('getSharePointIngestionConfiguration rejects missing site configuration', () => {
  assert.throws(() => getSharePointIngestionConfiguration({}), /RAG_SHAREPOINT_SITE_URL/i);
});
