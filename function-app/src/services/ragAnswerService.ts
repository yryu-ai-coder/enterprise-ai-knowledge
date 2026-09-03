import { ChatRequest } from '../models';
import { LibrarySearchResult } from './ragSearchService';

const MAX_RETRIEVED_SOURCES = 5;
const MAX_CITATION_EXCERPT_CHARS = 800;
const MAX_PROMPT_EXCERPT_CHARS = 1_200;

export function buildGroundedLibraryRequest(request: ChatRequest, results: LibrarySearchResult[]): ChatRequest {
  const uniqueResults: LibrarySearchResult[] = [];
  const seenDocuments = new Set<string>();
  for (const result of results) {
    if (seenDocuments.has(result.documentUrl)) {
      continue;
    }
    seenDocuments.add(result.documentUrl);
    uniqueResults.push(result);
    if (uniqueResults.length >= MAX_RETRIEVED_SOURCES) {
      break;
    }
  }

  return {
    ...request,
    mode: 'library-rag-search',
    knowledgeScope: 'library-wide-rag',
    selectedFiles: uniqueResults.map(result => ({
      name: result.documentName,
      url: result.documentUrl,
      fileType: result.fileType,
      libraryTitle: request.libraryName || 'Documents',
      lastModified: result.lastModified,
      snippet: result.content.slice(0, MAX_CITATION_EXCERPT_CHARS)
    })),
    documentSnippets: results.slice(0, MAX_RETRIEVED_SOURCES).map(result =>
      `Retrieved excerpt from ${result.documentName} (${result.folderPath}):\n${result.content.slice(0, MAX_PROMPT_EXCERPT_CHARS)}`
    )
  };
}
