import { ChatRequest } from '../models';
import { LibrarySearchResult } from './ragSearchService';

const MAX_RETRIEVED_SOURCES = 5;
const MAX_DOCUMENT_INVENTORY_SOURCES = 20;
const MAX_SPREADSHEET_RETRIEVED_SOURCES = 20;
const MAX_CITATION_EXCERPT_CHARS = 800;
const MAX_PROMPT_EXCERPT_CHARS = 1_200;
const BUDGET_VARIANCE_ANALYSIS_LABEL = 'Budget variance analysis:';

/**
 * A matching deterministic analysis is the authoritative evidence for an XLSX
 * budget question. Raw sheets remain indexed for other questions, but must not
 * compete with the computed result when the analysis section is available.
 */
export function prioritizeDeterministicSpreadsheetAnalysis(results: LibrarySearchResult[]): LibrarySearchResult[] {
  const deterministicAnalysis = results.filter(result =>
    result.fileType.toLowerCase() === 'xlsx'
    && result.sourceLabel?.startsWith(BUDGET_VARIANCE_ANALYSIS_LABEL)
  );
  return deterministicAnalysis.length ? deterministicAnalysis : results;
}

export function hasDeterministicSpreadsheetAnalysis(results: LibrarySearchResult[]): boolean {
  return results.some(result =>
    result.fileType.toLowerCase() === 'xlsx'
    && result.sourceLabel?.startsWith(BUDGET_VARIANCE_ANALYSIS_LABEL)
  );
}

export function requiresDeterministicBudgetAnalysis(question: string): boolean {
  return /계획|실제|초과|차이|분산|예산|variance|budget|planned|actual|overrun/i.test(question);
}

interface BudgetForecast {
  category: string;
  annualPlan: number;
  annualizedActual: number;
  forecastVariance: number;
}

function isAnnualBudgetForecastQuestion(question: string): boolean {
  return /연간.*(초과|가능)|연환산|남은 기간.*유지|annual.*(forecast|overrun)|run.?rate/i.test(question);
}

function formatBudgetAmount(value: number): string {
  return value.toLocaleString('en-US', { minimumFractionDigits: 0, maximumFractionDigits: 0 });
}

/** Returns a server-composed answer when the indexed analysis contains annual forecasts. */
export function buildDeterministicBudgetForecastAnswer(question: string, results: LibrarySearchResult[]): string | undefined {
  if (!isAnnualBudgetForecastQuestion(question)) return undefined;
  const analysis = results.find(result => result.sourceLabel?.startsWith(BUDGET_VARIANCE_ANALYSIS_LABEL));
  if (!analysis) return undefined;
  const forecastMarker = 'Annualized category forecast using the populated-month run rate:';
  const forecastText = analysis.content.includes(forecastMarker)
    ? analysis.content.slice(analysis.content.indexOf(forecastMarker) + forecastMarker.length)
    : analysis.content;
  const forecasts: BudgetForecast[] = [...forecastText.matchAll(/([^:.]+): planned for populated months [\d.-]+; actual for populated months [\d.-]+; populated-month variance [\d.-]+; annual plan ([\d.-]+); annualized actual ([\d.-]+); forecast variance ([\d.-]+)\./g)]
    .map(match => ({ category: match[1].trim(), annualPlan: Number(match[2]), annualizedActual: Number(match[3]), forecastVariance: Number(match[4]) }))
    .filter(forecast => Number.isFinite(forecast.annualPlan) && Number.isFinite(forecast.annualizedActual) && Number.isFinite(forecast.forecastVariance));
  if (!forecasts.length) return undefined;

  const highestRisk = forecasts.reduce((highest, forecast) => forecast.forecastVariance > highest.forecastVariance ? forecast : highest);
  if (highestRisk.forecastVariance <= 0) {
    return `6월까지 입력된 실제 지출을 같은 추세로 단순 연환산하면 연간 예산 초과가 예상되는 범주는 없습니다. 가장 예산 여유가 작은 범주는 ${highestRisk.category}입니다. 연간 예산 ${formatBudgetAmount(highestRisk.annualPlan)}, 연환산 실제 ${formatBudgetAmount(highestRisk.annualizedActual)}, 차이 ${formatBudgetAmount(highestRisk.forecastVariance)}입니다.`;
  }
  return `6월까지 입력된 실제 지출을 같은 추세로 단순 연환산하면 ${highestRisk.category}의 연간 예산 초과 가능성이 가장 높습니다. 연간 예산 ${formatBudgetAmount(highestRisk.annualPlan)}, 연환산 실제 ${formatBudgetAmount(highestRisk.annualizedActual)}, 차이 +${formatBudgetAmount(highestRisk.forecastVariance)}입니다.`;
}

export function isExplicitSpreadsheetSelection(
  selectedFileType: string | undefined,
  selectedDocumentName: string | undefined,
  selectedDocumentNames: string[]
): boolean {
  return selectedFileType === 'xlsx'
    || Boolean(selectedDocumentName?.toLowerCase().endsWith('.xlsx'))
    || selectedDocumentNames.some(name => name.toLowerCase().endsWith('.xlsx'));
}

export function isExplicitHba1cQuestion(question: string): boolean {
  return /(?:당화혈색소|hba1c|a1c|glycated\s+hemoglobin|hemoglobin\s+a1c)/i.test(question);
}

export function isAbnormalLabResultsQuestion(question: string): boolean {
  return /(?:검사|혈액|피검사|소변|lab|laboratory).*(?:비정상|정상\s*범위|높|낮|abnormal|high|low|out\s+of\s+range)/i.test(question)
    || /(?:비정상|정상\s*범위|높|낮|abnormal|high|low|out\s+of\s+range).*(?:검사|혈액|피검사|소변|lab|laboratory)/i.test(question);
}

export function expandRetrievalQuery(question: string): string {
  const normalizedQuestion = question.trim();
  if (isExplicitHba1cQuestion(normalizedQuestion)) {
    return `${normalizedQuestion} HbA1c Hemoglobin A1c glycated hemoglobin`;
  }
  if (isAbnormalLabResultsQuestion(normalizedQuestion)) {
    return `${normalizedQuestion} High Low Abnormal Positive Hemoglobin A1c HbA1c`;
  }
  return normalizedQuestion;
}

export function ensureInventoryEvidencePrefix(answer: string, indexedDocumentCount: number): string {
  const prefix = `확인된 ${indexedDocumentCount}개 인덱싱 문서 기준:`;
  const normalizedAnswer = answer.trim();
  return normalizedAnswer.startsWith(prefix) ? normalizedAnswer : `${prefix} ${normalizedAnswer}`;
}

export interface IDeterministicHba1cAnswer {
  answer: string;
  source: LibrarySearchResult;
}

export function buildDeterministicHba1cAnswer(_question: string, results: LibrarySearchResult[]): IDeterministicHba1cAnswer | undefined {
  for (const source of results) {
    const headingIndex = source.content.search(/(?:hemoglobin\s+a1c|hba1c)/i);
    if (headingIndex < 0) continue;
    const section = source.content.slice(headingIndex, headingIndex + 1_500);
    const range = section.match(/normal\s+range:\s*([0-9.]+)\s*-\s*([0-9.]+)\s*%?/i);
    const flaggedValues = [...section.matchAll(/\b([0-9]+(?:\.[0-9]+)?)\s+(High|Low)\b/gi)];
    const result = flaggedValues.at(-1);
    if (!range || !result) continue;

    return {
      answer: `PDF에 표시된 Hemoglobin A1c 결과는 ${result[1]}%이며 ${result[2]}로 표시되어 있습니다. 참고 범위는 ${range[1]}–${range[2]}%입니다.`,
      source
    };
  }

  return undefined;
}

export function isDocumentInventoryQuestion(question: string): boolean {
  const normalizedQuestion = question.trim();
  return /(?:라이브러리|라이브러이|library|폴더|folder).*(?:문서|파일|file).*(?:종류|어떤|무엇|있)/i.test(normalizedQuestion)
    || /(?:라이브러리|라이브러이|library|폴더|folder).*(?:종류|어떤|무엇).*(?:문서|파일|file)/i.test(normalizedQuestion)
    || /(?:문서|파일|file).*(?:종류|어떤|무엇|있).*(?:라이브러리|라이브러이|library|폴더|folder)/i.test(normalizedQuestion)
    || /(?:문서|파일|document|file).*(?:유형|종류|목록|리스트|types?|list).*(?:확인|보여|알려|있|what|show|list)/i.test(normalizedQuestion)
    || /(?:인덱싱|indexed).*(?:문서|파일|document|file).*(?:목록|리스트|list|보여|show)/i.test(normalizedQuestion);
}

function distinctDocuments(results: LibrarySearchResult[], limit: number): LibrarySearchResult[] {
  const documents: LibrarySearchResult[] = [];
  const seenDocumentUrls = new Set<string>();
  for (const result of results) {
    if (seenDocumentUrls.has(result.documentUrl)) continue;
    seenDocumentUrls.add(result.documentUrl);
    documents.push(result);
    if (documents.length >= limit) break;
  }
  return documents;
}

function hasCitationPageNumber(pageNumber: unknown): pageNumber is number {
  return typeof pageNumber === 'number' && Number.isInteger(pageNumber) && pageNumber > 0;
}

export function buildGroundedLibraryRequest(request: ChatRequest, results: LibrarySearchResult[]): ChatRequest {
  // Budget questions often need several chunks from the same worksheet to compare
  // categories/months. Keep a bounded but wider evidence window for XLSX only.
  const isInventoryOverview = isDocumentInventoryQuestion(request.question);
  const sourceResults = isInventoryOverview
    ? distinctDocuments(results, MAX_DOCUMENT_INVENTORY_SOURCES)
    : results;
  const maxPromptSources = isInventoryOverview
    ? MAX_DOCUMENT_INVENTORY_SOURCES
    : results.some(result => result.fileType.toLowerCase() === 'xlsx')
      ? MAX_SPREADSHEET_RETRIEVED_SOURCES
      : MAX_RETRIEVED_SOURCES;
  const uniqueResults: LibrarySearchResult[] = [];
  const seenSourceLocations = new Set<string>();
  for (const result of sourceResults) {
    // Keep distinct Office sections and OCR pages visible as separate source cards.
    // Plain document chunks remain collapsed to one card per document.
    const pageNumber = hasCitationPageNumber(result.pageNumber) ? result.pageNumber : undefined;
    const sourceLocation = result.sourceLabel
      ? `${result.documentUrl}:section:${result.sourceLabel}`
      : pageNumber === undefined
        ? result.documentUrl
        : `${result.documentUrl}:page:${pageNumber}`;
    if (seenSourceLocations.has(sourceLocation)) {
      continue;
    }
    seenSourceLocations.add(sourceLocation);
    uniqueResults.push(result);
    if (uniqueResults.length >= maxPromptSources) {
      break;
    }
  }

  return {
    ...request,
    mode: isInventoryOverview ? 'library-inventory-overview' : 'library-rag-search',
    knowledgeScope: 'library-wide-rag',
    selectedFiles: uniqueResults.map(result => {
      const pageNumber = hasCitationPageNumber(result.pageNumber) ? result.pageNumber : undefined;
      const locationLabel = result.sourceLabel || (pageNumber === undefined ? '' : `p. ${pageNumber}`);
      const locationPrefix = result.sourceLabel || (pageNumber === undefined ? '' : `Page ${pageNumber} — `);
      return {
        name: `${result.documentName}${locationLabel ? ` · ${locationLabel}` : ''}`,
        url: result.documentUrl,
        fileType: result.fileType,
        libraryTitle: request.libraryName || 'Documents',
        lastModified: result.lastModified,
        snippet: `${locationPrefix}${locationPrefix ? ' — ' : ''}${result.content.slice(0, MAX_CITATION_EXCERPT_CHARS)}`
      };
    }),
    documentSnippets: (isInventoryOverview ? uniqueResults : results.slice(0, maxPromptSources)).map(result => {
      const pageNumber = hasCitationPageNumber(result.pageNumber) ? result.pageNumber : undefined;
      const locationContext = result.sourceLabel || (pageNumber === undefined ? '' : `page ${pageNumber}; `);
      const overviewPreamble = isInventoryOverview
        ? `Inventory overview: one of ${uniqueResults.length} distinct indexed documents in the requested scope. `
        : '';
      return `${overviewPreamble}Retrieved excerpt from ${result.documentName} (${locationContext ? `${locationContext}; ` : ''}${result.folderPath}):\n${result.content.slice(0, MAX_PROMPT_EXCERPT_CHARS)}`;
    })
  };
}
