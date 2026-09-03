import { DefaultAzureCredential } from '@azure/identity';
import { SearchClient, SearchIndexClient, SearchIndex } from '@azure/search-documents';

export interface RagSearchConfiguration {
  endpoint: string;
  indexName: string;
}

export interface LibraryChunk {
  id: string;
  content: string;
  documentName: string;
  documentUrl: string;
  folderPath: string;
  fileType: string;
  lastModified: string;
  chunkOrdinal: number;
}

export interface LibrarySearchResult extends LibraryChunk {
  score?: number;
}

export function getRagSearchConfiguration(environment: Record<string, string | undefined> = process.env): RagSearchConfiguration {
  const endpoint = environment.AZURE_SEARCH_ENDPOINT?.trim().replace(/\/$/, '');

  if (!endpoint) {
    throw new Error('AZURE_SEARCH_ENDPOINT is required for library-wide RAG search.');
  }

  return {
    endpoint,
    indexName: environment.AZURE_SEARCH_INDEX_NAME?.trim() || 'nextcore-library-chunks'
  };
}

function getIndexClient(configuration: RagSearchConfiguration): SearchIndexClient {
  return new SearchIndexClient(configuration.endpoint, new DefaultAzureCredential());
}

function getChunkIndexDefinition(indexName: string): SearchIndex {
  return {
    name: indexName,
    fields: [
      { name: 'id', type: 'Edm.String', key: true, filterable: true, sortable: true },
      { name: 'content', type: 'Edm.String', searchable: true, analyzerName: 'en.microsoft' },
      { name: 'documentName', type: 'Edm.String', searchable: true, filterable: true, sortable: true, facetable: true },
      { name: 'documentUrl', type: 'Edm.String', filterable: true },
      { name: 'folderPath', type: 'Edm.String', searchable: true, filterable: true, facetable: true },
      { name: 'fileType', type: 'Edm.String', filterable: true, facetable: true },
      { name: 'lastModified', type: 'Edm.String', filterable: true, sortable: true },
      { name: 'chunkOrdinal', type: 'Edm.Int32', filterable: true, sortable: true }
    ]
  };
}

export async function ensureLibraryChunkIndex(): Promise<RagSearchConfiguration> {
  const configuration = getRagSearchConfiguration();
  const indexClient = getIndexClient(configuration);

  try {
    await indexClient.getIndex(configuration.indexName);
  } catch (error) {
    if ((error as { statusCode?: number }).statusCode !== 404) {
      throw error;
    }
    await indexClient.createIndex(getChunkIndexDefinition(configuration.indexName));
  }

  return configuration;
}

function getSearchClient(configuration: RagSearchConfiguration): SearchClient<LibraryChunk> {
  return new SearchClient<LibraryChunk>(configuration.endpoint, configuration.indexName, new DefaultAzureCredential());
}

export async function indexLibraryChunks(chunks: LibraryChunk[]): Promise<number> {
  if (chunks.length === 0) {
    return 0;
  }

  const configuration = await ensureLibraryChunkIndex();
  const client = getSearchClient(configuration);
  const result = await client.mergeOrUploadDocuments(chunks);
  const failed = result.results.filter(item => !item.succeeded);

  if (failed.length > 0) {
    throw new Error(`Azure AI Search rejected ${failed.length} chunk(s).`);
  }

  return result.results.length;
}

export async function searchLibraryChunks(query: string, top = 8): Promise<LibrarySearchResult[]> {
  const configuration = getRagSearchConfiguration();
  const client = getSearchClient(configuration);
  const response = await client.search(query, {
    top,
    select: ['id', 'content', 'documentName', 'documentUrl', 'folderPath', 'fileType', 'lastModified', 'chunkOrdinal']
  });

  const results: LibrarySearchResult[] = [];
  for await (const item of response.results) {
    results.push({ ...item.document, score: item.score });
  }

  return results;
}

export async function getLibrarySearchStatus(): Promise<{ configured: boolean; indexName?: string; indexExists: boolean }> {
  try {
    const configuration = getRagSearchConfiguration();
    const credential = new DefaultAzureCredential();
    const token = await credential.getToken('https://search.azure.com/.default');

    if (!token?.token) {
      throw new Error('Managed Identity did not return an Azure AI Search access token.');
    }

    const response = await fetch(
      `${configuration.endpoint}/indexes/${encodeURIComponent(configuration.indexName)}?api-version=2024-07-01`,
      { headers: { Authorization: `Bearer ${token.token}` } }
    );

    if (response.status === 404) {
      return { configured: true, indexName: configuration.indexName, indexExists: false };
    }
    if (!response.ok) {
      const detail = (await response.text()).slice(0, 500);
      throw new Error(`Azure AI Search returned ${response.status}: ${detail || response.statusText}`);
    }

    return { configured: true, indexName: configuration.indexName, indexExists: true };
  } catch (error) {
    if ((error as Error).message.includes('AZURE_SEARCH_ENDPOINT')) {
      return { configured: false, indexExists: false };
    }
    throw error;
  }
}
