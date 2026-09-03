import assert from 'node:assert/strict';
import test from 'node:test';
import { getRagSearchConfiguration } from '../src/services/ragSearchService';

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
