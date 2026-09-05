import assert from 'node:assert/strict';
import test from 'node:test';
import { buildLibraryChunks, chunkDocumentText, getSharePointIngestionConfiguration, validateRequestedLibraryScope } from '../src/services/ragIngestionService';

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
    siteUrl: 'https://youngryu.sharepoint.com/sites/enterprise-ai-knowledge',
    libraryId: 'drive-litigation',
    libraryName: 'Litigation Documents',
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

test('buildLibraryChunks never produces an Azure AI Search key beginning with an underscore', () => {
  const [chunk] = buildLibraryChunks({
    driveItemId: 'drive-78',
    siteUrl: 'https://contoso',
    libraryId: 'drive-78',
    libraryName: 'Litigation Documents',
    name: 'scanned.pdf',
    documentUrl: 'https://contoso/scanned.pdf',
    folderPath: '/Litigation Documents',
    fileType: 'pdf',
    lastModified: '2026-09-03T12:00:00Z',
    text: 'OCR text'
  });

  assert.match(chunk.id, /^[A-Za-z0-9][A-Za-z0-9_-]*$/);
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

test('validateRequestedLibraryScope accepts a named library only on the configured current site', () => {
  const scope = validateRequestedLibraryScope(
    { libraryName: 'Litigation Documents', siteUrl: 'https://contoso.sharepoint.com/sites/legal/' },
    { RAG_SHAREPOINT_SITE_URL: 'https://contoso.sharepoint.com/sites/legal' }
  );

  assert.deepEqual(scope, {
    siteUrl: 'https://contoso.sharepoint.com/sites/legal',
    libraryName: 'Litigation Documents'
  });
});

test('validateRequestedLibraryScope rejects blank library names and a different site', () => {
  const environment = { RAG_SHAREPOINT_SITE_URL: 'https://contoso.sharepoint.com/sites/legal' };
  assert.throws(() => validateRequestedLibraryScope({ libraryName: ' ' }, environment), /libraryName/i);
  assert.throws(
    () => validateRequestedLibraryScope({ libraryName: 'Documents', siteUrl: 'https://contoso.sharepoint.com/sites/other' }, environment),
    /configured SharePoint site/i
  );
});
