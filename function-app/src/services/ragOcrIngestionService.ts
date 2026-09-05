import { DefaultAzureCredential } from '@azure/identity';
import { getDocumentIntelligenceOcrConfiguration, runDocumentIntelligenceOcr, DocumentIntelligenceOcrResult } from './documentIntelligenceOcrService';
import { extractPdfText, PdfTextExtractionResult } from './pdfTextExtraction';
import {
  buildLibraryChunks,
  getSharePointIngestionConfiguration,
  resolveLibraryDrive,
  validateRequestedLibraryScope,
  RequestedLibraryScope,
  LibraryIngestionResult
} from './ragIngestionService';
import { indexLibraryChunks, LibraryChunk } from './ragSearchService';

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

export interface OcrLibraryIngestionResult extends LibraryIngestionResult {
  ocrProcessed: number;
}

export interface OcrLibraryIngestionDependencies {
  environment?: Record<string, string | undefined>;
  getGraphAccessToken?: () => Promise<string>;
  graphFetch?: (url: string, token: string) => Promise<Response>;
  graphJson?: <T>(url: string, token: string) => Promise<T>;
  extractPdfText?: (fileBytes: Buffer) => Promise<PdfTextExtractionResult>;
  runOcr?: (pdfBytes: Buffer) => Promise<DocumentIntelligenceOcrResult>;
  indexLibraryChunks?: (chunks: LibraryChunk[]) => Promise<number>;
}

const GRAPH_SCOPE = 'https://graph.microsoft.com/.default';
const MAX_EXTRACTED_TEXT_CHARS = 16_000;

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

async function defaultGraphFetch(url: string, token: string): Promise<Response> {
  const response = await fetch(url, { headers: { Authorization: `Bearer ${token}` } });
  if (!response.ok) {
    const detail = (await response.text()).slice(0, 500);
    throw new Error(`Microsoft Graph returned ${response.status}: ${detail || response.statusText}`);
  }
  return response;
}

async function defaultGraphJson<T>(url: string, token: string): Promise<T> {
  return await (await defaultGraphFetch(url, token)).json() as T;
}

async function defaultGraphAccessToken(): Promise<string> {
  const accessToken = await new DefaultAzureCredential().getToken(GRAPH_SCOPE);
  if (!accessToken?.token) {
    throw new Error('Managed Identity did not return a Microsoft Graph access token.');
  }
  return accessToken.token;
}

export async function ingestSharePointLibraryOcr(
  dependencies: OcrLibraryIngestionDependencies = {},
  request?: RequestedLibraryScope
): Promise<OcrLibraryIngestionResult> {
  const environment = dependencies.environment ?? process.env;
  const configured = getSharePointIngestionConfiguration(environment);
  const scope = validateRequestedLibraryScope(request ?? { libraryName: configured.libraryName }, environment);
  const ocrConfiguration = getDocumentIntelligenceOcrConfiguration(environment);
  const getGraphAccessToken = dependencies.getGraphAccessToken ?? defaultGraphAccessToken;
  const graphFetch = dependencies.graphFetch ?? defaultGraphFetch;
  const graphJson = dependencies.graphJson ?? defaultGraphJson;
  const token = await getGraphAccessToken();
  const resolvedScope = await resolveLibraryDrive(scope, token, graphJson);

  const files: GraphDriveItem[] = [];
  let nextUrl: string | undefined = `https://graph.microsoft.com/v1.0/drives/${resolvedScope.libraryId}/root/delta?$select=id,name,size,webUrl,lastModifiedDateTime,file,deleted,parentReference`;
  while (nextUrl) {
    const graphPage: GraphCollection<GraphDriveItem> = await graphJson<GraphCollection<GraphDriveItem>>(nextUrl, token);
    files.push(...(graphPage.value || []).filter((item: GraphDriveItem) => Boolean(item.file) && !item.deleted));
    nextUrl = graphPage['@odata.nextLink'];
  }

  const result: OcrLibraryIngestionResult = {
    siteUrl: resolvedScope.siteUrl,
    libraryName: resolvedScope.libraryName,
    discoveredFiles: files.length,
    indexedDocuments: 0,
    indexedChunks: 0,
    skippedUnsupported: 0,
    skippedTooLarge: 0,
    requiresOcr: 0,
    ocrProcessed: 0,
    failedFiles: []
  };
  const allChunks: LibraryChunk[] = [];
  const textExtractor = dependencies.extractPdfText ?? extractPdfText;
  const ocr = dependencies.runOcr ?? ((pdfBytes: Buffer) => runDocumentIntelligenceOcr(pdfBytes, ocrConfiguration));
  const indexChunks = dependencies.indexLibraryChunks ?? indexLibraryChunks;

  for (const file of files) {
    const fileType = getFileType(file.name);
    if (fileType !== 'pdf') {
      result.skippedUnsupported += 1;
      continue;
    }
    if (!file.size || file.size > ocrConfiguration.maxPdfBytes) {
      result.skippedTooLarge += 1;
      continue;
    }

    try {
      const contentResponse = await graphFetch(`https://graph.microsoft.com/v1.0/drives/${resolvedScope.libraryId}/items/${file.id}/content`, token);
      const fileBytes = Buffer.from(await contentResponse.arrayBuffer());
      const extraction = await textExtractor(fileBytes);
      if (!extraction.requiresOcr) {
        continue;
      }

      result.requiresOcr += 1;
      const ocrResult = await ocr(fileBytes);
      const pageChunks = ocrResult.pages.flatMap((page, pageIndex) => buildLibraryChunks({
        driveItemId: file.id,
        siteUrl: resolvedScope.siteUrl,
        libraryId: resolvedScope.libraryId,
        libraryName: resolvedScope.libraryName,
        name: file.name,
        documentUrl: file.webUrl || `${resolvedScope.siteUrl}/${file.name}`,
        folderPath: getFolderPath(file, resolvedScope.libraryName),
        fileType,
        lastModified: file.lastModifiedDateTime || '',
        text: page.text.slice(0, MAX_EXTRACTED_TEXT_CHARS),
        pageNumber: page.pageNumber,
        chunkOrdinalOffset: pageIndex * 1_000_000
      }));
      if (pageChunks.length > 0) {
        allChunks.push(...pageChunks);
        result.indexedDocuments += 1;
      }
      result.ocrProcessed += 1;
    } catch (error) {
      result.failedFiles.push({ name: file.name, reason: (error as Error).message.slice(0, 300) });
    }
  }

  if (allChunks.length > 0) {
    result.indexedChunks = await indexChunks(allChunks);
  }
  return result;
}
