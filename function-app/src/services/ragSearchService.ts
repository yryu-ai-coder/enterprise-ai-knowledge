import { DefaultAzureCredential } from '@azure/identity';
import { SearchClient, SearchIndexClient, SearchIndex } from '@azure/search-documents';

export interface RagSearchConfiguration {
  endpoint: string;
  indexName: string;
}

export interface LibraryChunk {
  id: string;
  content: string;
  siteUrl?: string;
  libraryId?: string;
  libraryName?: string;
  documentName: string;
  documentUrl: string;
  folderPath: string;
  fileType: string;
  lastModified: string;
  chunkOrdinal: number;
  pageNumber?: number;
  sourceLabel?: string;
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

export function getLibraryChunkIndexDefinition(indexName: string): SearchIndex {
  return {
    name: indexName,
    fields: [
      { name: 'id', type: 'Edm.String', key: true, filterable: true, sortable: true },
      { name: 'content', type: 'Edm.String', searchable: true, analyzerName: 'en.microsoft' },
      { name: 'siteUrl', type: 'Edm.String', filterable: true },
      { name: 'libraryId', type: 'Edm.String', filterable: true },
      { name: 'libraryName', type: 'Edm.String', filterable: true, facetable: true },
      { name: 'documentName', type: 'Edm.String', searchable: true, filterable: true, sortable: true, facetable: true },
      { name: 'documentUrl', type: 'Edm.String', filterable: true },
      { name: 'folderPath', type: 'Edm.String', searchable: true, filterable: true, facetable: true },
      { name: 'fileType', type: 'Edm.String', filterable: true, facetable: true },
      { name: 'lastModified', type: 'Edm.String', filterable: true, sortable: true },
      { name: 'chunkOrdinal', type: 'Edm.Int32', filterable: true, sortable: true },
      { name: 'pageNumber', type: 'Edm.Int32', filterable: true, sortable: true },
      { name: 'sourceLabel', type: 'Edm.String', filterable: true, sortable: true }
    ]
  };
}

export function doesLibraryChunkIndexRequireSchemaUpdate(existingIndex: SearchIndex): boolean {
  return !['siteUrl', 'libraryId', 'libraryName', 'pageNumber', 'sourceLabel'].every(requiredField => existingIndex.fields.some((field) => field.name === requiredField));
}

export async function ensureLibraryChunkIndex(): Promise<RagSearchConfiguration> {
  const configuration = getRagSearchConfiguration();
  const indexClient = getIndexClient(configuration);

  try {
    const existingIndex = await indexClient.getIndex(configuration.indexName);
    if (doesLibraryChunkIndexRequireSchemaUpdate(existingIndex)) {
      await indexClient.createOrUpdateIndex(getLibraryChunkIndexDefinition(configuration.indexName));
    }
  } catch (error) {
    if ((error as { statusCode?: number }).statusCode !== 404) {
      throw error;
    }
    await indexClient.createIndex(getLibraryChunkIndexDefinition(configuration.indexName));
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

export interface LibrarySearchOptions {
  libraryId: string;
  top?: number;
  documentUrl?: string;
  documentName?: string;
  documentNames?: string[];
  folderPath?: string;
  matchAll?: boolean;
}

function escapeODataString(value: string): string {
  return value.replace(/'/g, "''");
}

export function createLibraryIdFilter(libraryId: string): string {
  const normalizedLibraryId = libraryId.trim();
  if (!normalizedLibraryId) throw new Error('libraryId is required for library-scoped search.');
  return `libraryId eq '${escapeODataString(normalizedLibraryId)}'`;
}

export function createLibraryScopedFilter(libraryId: string, documentFilter?: string): string {
  const libraryFilter = createLibraryIdFilter(libraryId);
  return documentFilter ? `${libraryFilter} and ${documentFilter}` : libraryFilter;
}

export function createDocumentUrlFilter(documentUrl: string): string {
  return `documentUrl eq '${escapeODataString(documentUrl)}'`;
}

export function createDocumentNameFilter(documentName: string): string {
  return `documentName eq '${escapeODataString(documentName)}'`;
}

export function createDocumentNamesFilter(documentNames: string[]): string {
  const names = documentNames.map((name) => name.trim()).filter(Boolean);
  if (names.length === 0) throw new Error('At least one selected document name is required.');
  return `search.in(documentName, '${escapeODataString(names.join('|'))}', '|')`;
}

export function createFolderPathFilter(folderPath: string): string {
  const normalizedFolderPath = folderPath.trim().replace(/\/$/, '');
  if (!normalizedFolderPath) throw new Error('folderPath is required for current-folder search.');
  return `folderPath eq '${escapeODataString(normalizedFolderPath)}'`;
}

export function getSearchTextForLibrarySearch(query: string, matchAll = false): string {
  return matchAll ? '*' : query;
}

export function shouldFallbackToSelectedDocumentChunks(selectedDocumentUrl: string | undefined, resultCount: number): boolean {
  return Boolean(selectedDocumentUrl && resultCount === 0);
}

export async function searchLibraryChunks(query: string, options: LibrarySearchOptions = { libraryId: '' }): Promise<LibrarySearchResult[]> {
  const top = options.top ?? 8;
  const configuration = getRagSearchConfiguration();
  const client = getSearchClient(configuration);
  const documentFilter = options.documentUrl
    ? createDocumentUrlFilter(options.documentUrl)
    : options.documentNames && options.documentNames.length > 0
      ? createDocumentNamesFilter(options.documentNames)
      : options.documentName
        ? createDocumentNameFilter(options.documentName)
        : options.folderPath
          ? createFolderPathFilter(options.folderPath)
          : undefined;
  const response = await client.search(getSearchTextForLibrarySearch(query, options.matchAll), {
    top,
    filter: createLibraryScopedFilter(options.libraryId, documentFilter),
    select: ['id', 'content', 'siteUrl', 'libraryId', 'libraryName', 'documentName', 'documentUrl', 'folderPath', 'fileType', 'lastModified', 'chunkOrdinal', 'pageNumber', 'sourceLabel']
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
