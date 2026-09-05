import { createHash } from 'node:crypto';
import { DefaultAzureCredential } from '@azure/identity';
import { LibraryChunk } from './ragSearchService';

export interface ChunkOptions {
  maxChars?: number;
  overlapChars?: number;
}

export interface SourceDocumentForChunks {
  driveItemId: string;
  siteUrl: string;
  libraryId: string;
  libraryName: string;
  name: string;
  documentUrl: string;
  folderPath: string;
  fileType: string;
  lastModified: string;
  text: string;
  pageNumber?: number;
  sourceLabel?: string;
  chunkOrdinalOffset?: number;
  chunkOptions?: ChunkOptions;
}

export interface SharePointIngestionConfiguration {
  siteUrl: string;
  libraryName: string;
}

export interface RequestedLibraryScope {
  siteUrl?: string;
  libraryName?: string;
}

export interface ValidatedLibraryScope {
  siteUrl: string;
  libraryName: string;
}

export function validateRequestedLibraryScope(
  request: RequestedLibraryScope,
  environment: Record<string, string | undefined> = process.env
): ValidatedLibraryScope {
  const configuration = getSharePointIngestionConfiguration(environment);
  const libraryName = request.libraryName?.trim();
  if (!libraryName) {
    throw new Error('libraryName is required for library ingestion.');
  }

  const requestedSiteUrl = request.siteUrl?.trim().replace(/\/$/, '');
  if (requestedSiteUrl && requestedSiteUrl !== configuration.siteUrl) {
    throw new Error('The requested siteUrl must match the configured SharePoint site.');
  }

  return { siteUrl: configuration.siteUrl, libraryName };
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
  const hash = createHash('sha256')
    .update(`${driveItemId}:${chunkOrdinal}`)
    .digest('base64url');

  // Azure AI Search document keys may contain underscores but cannot begin with one.
  return hash.startsWith('_') ? `k${hash}` : hash;
}

export function buildLibraryChunks(source: SourceDocumentForChunks): LibraryChunk[] {
  const chunkOrdinalOffset = source.chunkOrdinalOffset ?? 0;
  return chunkDocumentText(source.text, source.chunkOptions).map((content, chunkIndex) => {
    const chunkOrdinal = chunkOrdinalOffset + chunkIndex;
    return {
      id: createChunkId(source.driveItemId, chunkOrdinal),
      content,
      siteUrl: source.siteUrl,
      libraryId: source.libraryId,
      libraryName: source.libraryName,
      documentName: source.name,
      documentUrl: source.documentUrl,
      folderPath: source.folderPath,
      fileType: source.fileType,
      lastModified: source.lastModified,
      chunkOrdinal,
      ...(source.pageNumber === undefined ? {} : { pageNumber: source.pageNumber }),
      ...(source.sourceLabel === undefined ? {} : { sourceLabel: source.sourceLabel })
    };
  });
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
const MAX_OFFICE_BYTES = 10 * 1024 * 1024;
const SUPPORTED_FILE_TYPES = new Set(['pdf', 'docx', 'pptx', 'xlsx']);
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

export interface ResolvedLibraryScope extends ValidatedLibraryScope {
  siteId: string;
  libraryId: string;
}

export async function resolveLibraryDrive(
  scope: ValidatedLibraryScope,
  token: string,
  requestGraphJson: <T>(url: string, token: string) => Promise<T>
): Promise<ResolvedLibraryScope> {
  const siteUri = new URL(scope.siteUrl);
  const site = await requestGraphJson<{ id: string }>(
    `https://graph.microsoft.com/v1.0/sites/${siteUri.host}:${siteUri.pathname}`,
    token
  );
  const drives = await requestGraphJson<GraphCollection<GraphDrive>>(`https://graph.microsoft.com/v1.0/sites/${site.id}/drives`, token);
  const drive = drives.value?.find(candidate => candidate.name?.toLowerCase() === scope.libraryName.toLowerCase());
  if (!drive?.id) {
    throw new Error(`Document library '${scope.libraryName}' was not found on the configured SharePoint site.`);
  }
  return { ...scope, siteId: site.id, libraryId: drive.id };
}

export async function resolveRequestedLibraryDrive(request: RequestedLibraryScope): Promise<ResolvedLibraryScope> {
  const scope = validateRequestedLibraryScope(request);
  const accessToken = await new DefaultAzureCredential().getToken(GRAPH_SCOPE);
  if (!accessToken?.token) {
    throw new Error('Managed Identity did not return a Microsoft Graph access token.');
  }
  return await resolveLibraryDrive(scope, accessToken.token, graphJson);
}

export async function ingestSharePointLibrary(request: RequestedLibraryScope): Promise<LibraryIngestionResult> {
  const scope = validateRequestedLibraryScope(request);
  const credential = new DefaultAzureCredential();
  const accessToken = await credential.getToken(GRAPH_SCOPE);
  if (!accessToken?.token) {
    throw new Error('Managed Identity did not return a Microsoft Graph access token.');
  }

  const token = accessToken.token;
  const resolvedScope = await resolveLibraryDrive(scope, token, graphJson);

  const files: GraphDriveItem[] = [];
  let nextUrl: string | undefined = `https://graph.microsoft.com/v1.0/drives/${resolvedScope.libraryId}/root/delta?$select=id,name,size,webUrl,lastModifiedDateTime,file,deleted,parentReference`;
  while (nextUrl) {
    const page: GraphCollection<GraphDriveItem> = await graphJson<GraphCollection<GraphDriveItem>>(nextUrl, token);
    files.push(...(page.value || []).filter((item: GraphDriveItem) => Boolean(item.file) && !item.deleted));
    nextUrl = page['@odata.nextLink'];
  }

  const result: LibraryIngestionResult = {
    siteUrl: resolvedScope.siteUrl,
    libraryName: resolvedScope.libraryName,
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
  const { extractOfficeSections } = await import('./officeDocumentExtraction');

  for (const file of files) {
    const fileType = getFileType(file.name);
    if (!SUPPORTED_FILE_TYPES.has(fileType)) {
      result.skippedUnsupported += 1;
      continue;
    }
    if (!file.size || file.size > (fileType === 'pdf' ? MAX_PDF_BYTES : MAX_OFFICE_BYTES)) {
      result.skippedTooLarge += 1;
      continue;
    }

    try {
      const contentResponse = await graphFetch(`https://graph.microsoft.com/v1.0/drives/${resolvedScope.libraryId}/items/${file.id}/content`, token);
      const bytes = Buffer.from(await contentResponse.arrayBuffer());
      const source = {
        driveItemId: file.id,
        siteUrl: resolvedScope.siteUrl,
        libraryId: resolvedScope.libraryId,
        libraryName: resolvedScope.libraryName,
        name: file.name,
        documentUrl: file.webUrl || `${resolvedScope.siteUrl}/${file.name}`,
        folderPath: getFolderPath(file, resolvedScope.libraryName),
        fileType,
        lastModified: file.lastModifiedDateTime || ''
      };
      let chunks: LibraryChunk[];
      if (fileType === 'pdf') {
        const extraction = await extractPdfText(bytes);
        if (extraction.requiresOcr) {
          result.requiresOcr += 1;
          continue;
        }
        chunks = buildLibraryChunks({ ...source, text: extraction.text.slice(0, MAX_EXTRACTED_TEXT_CHARS) });
      } else {
        const sections = await extractOfficeSections(fileType, bytes);
        chunks = sections.flatMap((section, index) => buildLibraryChunks({
          ...source,
          text: section.text.slice(0, MAX_EXTRACTED_TEXT_CHARS),
          sourceLabel: section.citationLabel,
          chunkOrdinalOffset: index * 1_000
        }));
      }
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
