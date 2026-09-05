import assert from 'node:assert/strict';
import test from 'node:test';
import { createDocumentNameFilter, createDocumentUrlFilter, createFolderPathFilter, createLibraryIdFilter, createLibraryScopedFilter, doesLibraryChunkIndexRequireSchemaUpdate, getLibraryChunkIndexDefinition, getRagSearchConfiguration, getSearchTextForLibrarySearch, shouldFallbackToSelectedDocumentChunks } from '../src/services/ragSearchService';

test('getRagSearchConfiguration uses the configured endpoint and a stable default index name', () => {
  const config = getRagSearchConfiguration({
    AZURE_SEARCH_ENDPOINT: 'https://nextcore-ai-search.search.windows.net'
  });

  assert.equal(config.endpoint, 'https://nextcore-ai-search.search.windows.net');
  assert.equal(config.indexName, 'nextcore-library-chunks');
});

test('getRagSearchConfiguration rejects a missing endpoint before any Azure call', () => {
  assert.throws(
    () => getRagSearchConfiguration({}),
    /AZURE_SEARCH_ENDPOINT/i
  );
});

test('createDocumentUrlFilter scopes a search to one exact SharePoint document URL', () => {
  assert.equal(
    createDocumentUrlFilter("https://contoso.sharepoint.com/sites/legal/Doc's.pdf"),
    "documentUrl eq 'https://contoso.sharepoint.com/sites/legal/Doc''s.pdf'"
  );
});

test('createDocumentNameFilter provides an exact selected-file fallback without depending on URL representation', () => {
  assert.equal(
    createDocumentNameFilter("Exhibit G - Ava's text.pdf"),
    "documentName eq 'Exhibit G - Ava''s text.pdf'"
  );
});

test('createFolderPathFilter keeps current-folder retrieval inside one exact indexed folder', () => {
  assert.equal(
    createFolderPathFilter("/Litigation Documents/04 Office RAG Test/Client's files"),
    "folderPath eq '/Litigation Documents/04 Office RAG Test/Client''s files'"
  );
});

test('getSearchTextForLibrarySearch uses a wildcard fallback for a selected document', () => {
  assert.equal(getSearchTextForLibrarySearch('상대방이 집에 돌아간 시간은?', true), '*');
  assert.equal(getSearchTextForLibrarySearch('상대방이 집에 돌아간 시간은?', false), '상대방이 집에 돌아간 시간은?');
});

test('shouldFallbackToSelectedDocumentChunks only retries an empty selected-document search', () => {
  assert.equal(shouldFallbackToSelectedDocumentChunks('https://contoso.sharepoint.com/scan.pdf', 0), true);
  assert.equal(shouldFallbackToSelectedDocumentChunks('https://contoso.sharepoint.com/scan.pdf', 1), false);
  assert.equal(shouldFallbackToSelectedDocumentChunks(undefined, 0), false);
});

test('LibraryChunk Search schema makes optional pageNumber filterable and sortable', () => {
  const schema = getLibraryChunkIndexDefinition('test-index');
  const pageNumber = schema.fields.find(field => field.name === 'pageNumber');

  assert.deepEqual(pageNumber, {
    name: 'pageNumber',
    type: 'Edm.Int32',
    filterable: true,
    sortable: true
  });
});

test('doesLibraryChunkIndexRequireSchemaUpdate migrates existing indexes without the OCR page field', () => {
  assert.equal(doesLibraryChunkIndexRequireSchemaUpdate({
    name: 'nextcore-library-chunks',
    fields: [{ name: 'id', type: 'Edm.String' }]
  }), true);
  assert.equal(doesLibraryChunkIndexRequireSchemaUpdate(getLibraryChunkIndexDefinition('nextcore-library-chunks')), false);
});

test('library scope filter is mandatory and composes with selected document filters', () => {
  assert.equal(createLibraryIdFilter("drive'one"), "libraryId eq 'drive''one'");
  assert.equal(
    createLibraryScopedFilter('drive-one', createDocumentUrlFilter('https://contoso/doc.pdf')),
    "libraryId eq 'drive-one' and documentUrl eq 'https://contoso/doc.pdf'"
  );
  assert.throws(() => createLibraryIdFilter(' '), /libraryId/i);
});

test('LibraryChunk Search schema stores site and resolved drive scope fields', () => {
  const schema = getLibraryChunkIndexDefinition('test-index');
  assert.deepEqual(schema.fields.find(field => field.name === 'siteUrl'), { name: 'siteUrl', type: 'Edm.String', filterable: true });
  assert.deepEqual(schema.fields.find(field => field.name === 'libraryId'), { name: 'libraryId', type: 'Edm.String', filterable: true });
  assert.deepEqual(schema.fields.find(field => field.name === 'libraryName'), { name: 'libraryName', type: 'Edm.String', filterable: true, facetable: true });
});

test('same filename selections are isolated by their resolved drive IDs', () => {
  const selectedFileFilter = createDocumentNameFilter('status.pdf');
  assert.equal(createLibraryScopedFilter('drive-a', selectedFileFilter), "libraryId eq 'drive-a' and documentName eq 'status.pdf'");
  assert.equal(createLibraryScopedFilter('drive-b', selectedFileFilter), "libraryId eq 'drive-b' and documentName eq 'status.pdf'");
});
