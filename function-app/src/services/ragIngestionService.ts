import { createHash } from 'node:crypto';
import { DefaultAzureCredential } from '@azure/identity';
import { LibraryChunk } from './ragSearchService';

export interface ChunkOptions {
  maxChars?: number;
  overlapChars?: number;
}

export interface SourceDocumentForChunks {
  driveItemId: string;
  name: string;
  documentUrl: string;
  folderPath: string;
  fileType: string;
  lastModified: string;
  text: string;
  chunkOptions?: ChunkOptions;
}

export interface SharePointIngestionConfiguration {
  siteUrl: string;
  libraryName: string;
}

export function getSharePointIngestionConfiguration(
  environment: Record<string, string | undefined> = process.env
): SharePointIngestionConfiguration {
  const siteUrl = environment.RAG_SHAREPOINT_SITE_URL?.trim().replace(/\/$/, '');
  if (!siteUrl) {
    throw new Error('RAG_SHAREPOINT_SITE_URL is required for library ingestion.');
  }

  return {
    siteUrl,
    libraryName: environment.RAG_SHAREPOINT_LIBRARY_NAME?.trim() || 'Litigation Documents'
  };
}

const DEFAULT_MAX_CHARS = 1_500;
const DEFAULT_OVERLAP_CHARS = 180;

function normalizeText(text: string): string {
  return text.replace(/\s+/g, ' ').trim();
}

export function chunkDocumentText(text: string, options: ChunkOptions = {}): string[] {
  const maxChars = options.maxChars ?? DEFAULT_MAX_CHARS;
  const overlapChars = options.overlapChars ?? DEFAULT_OVERLAP_CHARS;
  const normalized = normalizeText(text);

  if (maxChars < 1 || overlapChars < 0 || overlapChars >= maxChars) {
    throw new Error('Chunk options require maxChars > 0 and 0 <= overlapChars < maxChars.');
  }
  if (!normalized) {
    return [];
  }

  const words = normalized.split(' ');
  const chunks: string[] = [];
  let currentWords: string[] = [];

  for (const word of words) {
    const candidate = [...currentWords, word].join(' ');
    if (candidate.length <= maxChars || currentWords.length === 0) {
      currentWords.push(word);
      continue;
    }

    chunks.push(currentWords.join(' '));
    const trailingWords: string[] = [];
    let trailingLength = 0;
    for (let index = currentWords.length - 1; index >= 0; index -= 1) {
      const trailingWord = currentWords[index];
      const candidateLength = trailingLength === 0 ? trailingWord.length : trailingWord.length + 1 + trailingLength;
      if (candidateLength > overlapChars) {
        break;
      }
      trailingWords.unshift(trailingWord);
      trailingLength = candidateLength;
    }
    currentWords = [...trailingWords, word];
  }

  if (currentWords.length > 0) {
    chunks.push(currentWords.join(' '));
  }

  return chunks;
}

function createChunkId(driveItemId: string, chunkOrdinal: number): string {
  return createHash('sha256')
    .update(`${driveItemId}:${chunkOrdinal}`)
    .digest('base64url');
}

export function buildLibraryChunks(source: SourceDocumentForChunks): LibraryChunk[] {
  return chunkDocumentText(source.text, source.chunkOptions).map((content, chunkOrdinal) => ({
    id: createChunkId(source.driveItemId, chunkOrdinal),
    content,
    documentName: source.name,
    documentUrl: source.documentUrl,
    folderPath: source.folderPath,
    fileType: source.fileType,
    lastModified: source.lastModified,
    chunkOrdinal
  }));
}

interface GraphDriveItem {
  id: string;
  name: string;
  size?: number;
  webUrl?: string;
  lastModifiedDateTime?: string;
  file?: { mimeType?: string };
  deleted?: Record<string, unknown>;
  parentReference?: { path?: string };
}

interface GraphDrive {
  id: string;
  name?: string;
}

interface GraphCollection<T> {
  value?: T[];
  '@odata.nextLink'?: string;
}

export interface LibraryIngestionResult {
  siteUrl: string;
  libraryName: string;
  discoveredFiles: number;
  indexedDocuments: number;
  indexedChunks: number;
  skippedUnsupported: number;
  skippedTooLarge: number;
  requiresOcr: number;
  failedFiles: Array<{ name: string; reason: string }>;
}

const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const MAX_PDF_BYTES = 4 * 1024 * 1024;
const MAX_EXTRACTED_TEXT_CHARS = 16_000;

async function graphFetch(url: string, token: string): Promise<Response> {
  const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!response.ok) {
    const detail = (await response.text()).slice(0, 500);
    throw new Error(`Microsoft Graph returned ${response.status}: ${detail || response.statusText}`);
  }
  return response;
}

async function graphJson<T>(url: string, token: string): Promise<T> {
  return await (await graphFetch(url, token)).json() as T;
}

function getFileType(name: string): string {
  const lastDot = name.lastIndexOf('.');
  return lastDot > -1 ? name.slice(lastDot + 1).toLowerCase() : '';
}

function getFolderPath(item: GraphDriveItem, libraryName: string): string {
  const graphPath = item.parentReference?.path || '';
  const rootMarker = '/root:';
  const relative = graphPath.includes(rootMarker) ? graphPath.slice(graphPath.indexOf(rootMarker) + rootMarker.length) : '';
  return `/${libraryName}${relative}`.replace(/\/$/, '') || `/${libraryName}`;
}

export async function ingestSharePointLibrary(): Promise<LibraryIngestionResult> {
  const configuration = getSharePointIngestionConfiguration();
  const siteUri = new URL(configuration.siteUrl);
  const credential = new DefaultAzureCredential();
  const accessToken = await credential.getToken(GRAPH_SCOPE);
  if (!accessToken?.token) {
    throw new Error('Managed Identity did not return a Microsoft Graph access token.');
  }

  const token = accessToken.token;
  const site = await graphJson<{ id: string }>(
    `https://graph.microsoft.com/v1.0/sites/${siteUri.host}:${siteUri.pathname}`,
    token
  );
  const drives = await graphJson<GraphCollection<GraphDrive>>(`https://graph.microsoft.com/v1.0/sites/${site.id}/drives`, token);
  const drive = drives.value?.find(candidate => candidate.name?.toLowerCase() === configuration.libraryName.toLowerCase());
  if (!drive) {
    throw new Error(`Document library '${configuration.libraryName}' was not found on the configured SharePoint site.`);
  }

  const files: GraphDriveItem[] = [];
  let nextUrl: string | undefined = `https://graph.microsoft.com/v1.0/drives/${drive.id}/root/delta?$select=id,name,size,webUrl,lastModifiedDateTime,file,deleted,parentReference`;
  while (nextUrl) {
    const page: GraphCollection<GraphDriveItem> = await graphJson<GraphCollection<GraphDriveItem>>(nextUrl, token);
    files.push(...(page.value || []).filter((item: GraphDriveItem) => Boolean(item.file) && !item.deleted));
    nextUrl = page['@odata.nextLink'];
  }

  const result: LibraryIngestionResult = {
    siteUrl: configuration.siteUrl,
    libraryName: configuration.libraryName,
    discoveredFiles: files.length,
    indexedDocuments: 0,
    indexedChunks: 0,
    skippedUnsupported: 0,
    skippedTooLarge: 0,
    requiresOcr: 0,
    failedFiles: []
  };
  const allChunks: LibraryChunk[] = [];
  const { extractPdfText } = await import('./pdfTextExtraction');

  for (const file of files) {
    const fileType = getFileType(file.name);
    if (fileType !== 'pdf') {
      result.skippedUnsupported += 1;
      continue;
    }
    if (!file.size || file.size > MAX_PDF_BYTES) {
      result.skippedTooLarge += 1;
      continue;
    }

    try {
      const contentResponse = await graphFetch(`https://graph.microsoft.com/v1.0/drives/${drive.id}/items/${file.id}/content`, token);
      const extraction = await extractPdfText(Buffer.from(await contentResponse.arrayBuffer()));
      if (extraction.requiresOcr) {
        result.requiresOcr += 1;
        continue;
      }

      const chunks = buildLibraryChunks({
        driveItemId: file.id,
        name: file.name,
        documentUrl: file.webUrl || `${configuration.siteUrl}/${file.name}`,
        folderPath: getFolderPath(file, configuration.libraryName),
        fileType,
        lastModified: file.lastModifiedDateTime || '',
        text: extraction.text.slice(0, MAX_EXTRACTED_TEXT_CHARS)
      });
      if (chunks.length > 0) {
        allChunks.push(...chunks);
        result.indexedDocuments += 1;
      }
    } catch (error) {
      result.failedFiles.push({ name: file.name, reason: (error as Error).message.slice(0, 300) });
    }
  }

  if (allChunks.length > 0) {
    const { indexLibraryChunks } = await import('./ragSearchService');
    result.indexedChunks = await indexLibraryChunks(allChunks);
  }

  return result;
}
