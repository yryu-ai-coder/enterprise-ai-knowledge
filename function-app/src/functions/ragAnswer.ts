import { app, HttpRequest, HttpResponseInit, InvocationContext } from '@azure/functions';
import { ChatRequest } from '../models';
import { answerQuestion } from '../services/aiProvider';
import { buildDeterministicBudgetForecastAnswer, buildGroundedLibraryRequest, hasDeterministicSpreadsheetAnalysis, prioritizeDeterministicSpreadsheetAnalysis, requiresDeterministicBudgetAnalysis } from '../services/ragAnswerService';
import { resolveRequestedLibraryDrive } from '../services/ragIngestionService';
import { searchLibraryChunks, shouldFallbackToSelectedDocumentChunks } from '../services/ragSearchService';

const MAX_SELECTED_SPREADSHEET_CHUNKS = 100;

function normalizeCurrentFolderPath(folderPath: string | undefined, libraryName: string): string | undefined {
  const requestedPath = folderPath?.trim().replace(/\/$/, '');
  if (!requestedPath) return undefined;
  const marker = `/${libraryName}`.toLowerCase();
  const markerIndex = requestedPath.toLowerCase().indexOf(marker);
  if (markerIndex < 0) return undefined;
  return requestedPath.slice(markerIndex).replace(/\/$/, '');
}

// App Service Authentication validates the Microsoft Entra bearer token first.
app.http('rag-answer', {
  methods: ['POST'],
  authLevel: 'anonymous',
  route: 'rag/answer',
  handler: async (request: HttpRequest, context: InvocationContext): Promise<HttpResponseInit> => {
    const requestId = context.invocationId;
    try {
      const body = await request.json() as ChatRequest;
      const question = body.question?.trim();
      if (!question) {
        return { status: 400, jsonBody: { error: 'Question is required.' } };
      }
      if (question.length > 1_000) {
        return { status: 400, jsonBody: { error: 'Question is too long.' } };
      }

      const resolvedLibrary = await resolveRequestedLibraryDrive({ siteUrl: body.siteUrl, libraryName: body.libraryName });
      const selectedDocumentUrl = (body as ChatRequest & { selectedDocumentUrl?: string }).selectedDocumentUrl?.trim();
      const selectedDocumentNames = (body as ChatRequest & { selectedDocumentNames?: string[] }).selectedDocumentNames?.map((name) => name.trim()).filter(Boolean) || [];
      const selectedDocumentName = body.selectedFiles?.[0]?.name?.trim();
      const selectedFileType = body.selectedFiles?.[0]?.fileType?.trim().toLowerCase();
      const retrievalTop = selectedFileType === 'xlsx' ? 20 : undefined;
      const requestedFolderScope = body.knowledgeScope === 'current-folder-rag'
        ? normalizeCurrentFolderPath(body.folderPath, resolvedLibrary.libraryName)
        : undefined;
      if (body.knowledgeScope === 'current-folder-rag' && !requestedFolderScope) {
        throw new Error('The requested folder must belong to the selected document library.');
      }
      const isCurrentFolderScope = Boolean(requestedFolderScope && requestedFolderScope.toLowerCase() !== `/${resolvedLibrary.libraryName}`.toLowerCase());
      let results = await searchLibraryChunks(question, selectedDocumentUrl
        ? { libraryId: resolvedLibrary.libraryId, documentUrl: selectedDocumentUrl, top: retrievalTop }
        : selectedDocumentNames.length > 1
          ? { libraryId: resolvedLibrary.libraryId, documentNames: selectedDocumentNames, top: retrievalTop }
          : isCurrentFolderScope
            ? { libraryId: resolvedLibrary.libraryId, folderPath: requestedFolderScope!, top: retrievalTop }
            : { libraryId: resolvedLibrary.libraryId, top: retrievalTop });
      if (selectedDocumentNames.length > 1 && results.length === 0) {
        results = await searchLibraryChunks(question, { libraryId: resolvedLibrary.libraryId, documentNames: selectedDocumentNames, matchAll: true, top: retrievalTop });
      }
      if (shouldFallbackToSelectedDocumentChunks(selectedDocumentUrl, results.length)) {
        results = await searchLibraryChunks(question, { libraryId: resolvedLibrary.libraryId, documentUrl: selectedDocumentUrl, matchAll: true, top: retrievalTop });
      }
      if (selectedDocumentNames.length <= 1 && shouldFallbackToSelectedDocumentChunks(selectedDocumentName, results.length)) {
        results = await searchLibraryChunks(question, { libraryId: resolvedLibrary.libraryId, documentName: selectedDocumentName, matchAll: true, top: retrievalTop });
      }
      if (isCurrentFolderScope && results.length === 0) {
        results = await searchLibraryChunks('*', {
          libraryId: resolvedLibrary.libraryId,
          folderPath: requestedFolderScope!,
          matchAll: true,
          top: retrievalTop
        });
      }
      const selectedDocumentIsSpreadsheet = selectedFileType === 'xlsx'
        || Boolean(selectedDocumentName?.toLowerCase().endsWith('.xlsx'))
        || selectedDocumentNames.some(name => name.toLowerCase().endsWith('.xlsx'))
        || results.some(result => result.fileType.toLowerCase() === 'xlsx');
      if (selectedDocumentIsSpreadsheet && selectedDocumentUrl) {
        // A document can contain many raw-sheet chunks before the final analysis
        // section. Inspect a bounded full selected-document set, never the library.
        let spreadsheetSections = await searchLibraryChunks('*', {
          libraryId: resolvedLibrary.libraryId,
          documentUrl: selectedDocumentUrl,
          top: MAX_SELECTED_SPREADSHEET_CHUNKS
        });
        // The SharePoint UI URL can differ from Graph's indexed webUrl. Keep the
        // existing exact-name fallback inside the same library for that case.
        if (!hasDeterministicSpreadsheetAnalysis(spreadsheetSections) && selectedDocumentName) {
          spreadsheetSections = await searchLibraryChunks('*', {
            libraryId: resolvedLibrary.libraryId,
            documentName: selectedDocumentName,
            top: MAX_SELECTED_SPREADSHEET_CHUNKS
          });
        }
        results = prioritizeDeterministicSpreadsheetAnalysis([...results, ...spreadsheetSections]);
      }
      if (selectedDocumentIsSpreadsheet && requiresDeterministicBudgetAnalysis(question) && !hasDeterministicSpreadsheetAnalysis(results)) {
        return {
          status: 200,
          jsonBody: {
            answer: 'This spreadsheet does not yet have the verified budget calculation result. Run Index supported documents now for this Excel file, then ask the question again.',
            citations: [],
            provider: 'rag-search',
            requestId,
            status: 'success',
            metadata: { retrievalCount: results.length, knowledgeScope: 'library-wide-rag', deterministicAnalysisRequired: true }
          }
        };
      }
      if (results.length === 0) {
        return {
          status: 200,
          jsonBody: {
            answer: 'No matching indexed document excerpts were found in this library. Run library ingestion, or refine the question.',
            citations: [],
            provider: 'rag-search',
            requestId,
            status: 'success',
            metadata: { retrievalCount: 0, knowledgeScope: 'library-wide-rag' }
          }
        };
      }

      const groundedResults = selectedDocumentIsSpreadsheet
        ? prioritizeDeterministicSpreadsheetAnalysis(results)
        : results;
      const deterministicForecastAnswer = selectedDocumentIsSpreadsheet
        ? buildDeterministicBudgetForecastAnswer(question, groundedResults)
        : undefined;
      if (deterministicForecastAnswer) {
        const analysis = groundedResults.find(result => result.sourceLabel?.startsWith('Budget variance analysis:'))!;
        return {
          status: 200,
          jsonBody: {
            answer: deterministicForecastAnswer,
            citations: [{
              title: `${analysis.documentName} · ${analysis.sourceLabel}`,
              url: analysis.documentUrl,
              snippet: analysis.content.slice(0, 800)
            }],
            provider: 'rag-search',
            requestId,
            status: 'success',
            metadata: { retrievalCount: groundedResults.length, knowledgeScope: 'library-wide-rag', deterministicBudgetForecast: true }
          }
        };
      }
      const groundedRequest = buildGroundedLibraryRequest({ ...body, question }, groundedResults);
      const response = await answerQuestion(groundedRequest, requestId);
      return {
        status: 200,
        jsonBody: {
          ...response,
          metadata: {
            ...(response.metadata || {}),
            retrievalCount: results.length,
            knowledgeScope: 'library-wide-rag'
          }
        }
      };
    } catch (error) {
      const detail = (error as Error).message || 'Library-wide RAG answer failed.';
      context.error(detail);
      return {
        status: 500,
        jsonBody: {
          status: 'error',
          error: 'Library-wide RAG answer failed.',
          detail,
          requestId
        }
      };
    }
  }
});
