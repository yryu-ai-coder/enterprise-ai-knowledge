import { ChatRequest } from '../models';
import { LibrarySearchResult } from './ragSearchService';

const MAX_RETRIEVED_SOURCES = 5;
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

export function buildGroundedLibraryRequest(request: ChatRequest, results: LibrarySearchResult[]): ChatRequest {
  // Budget questions often need several chunks from the same worksheet to compare
  // categories/months. Keep a bounded but wider evidence window for XLSX only.
  const maxPromptSources = results.some(result => result.fileType.toLowerCase() === 'xlsx')
    ? MAX_SPREADSHEET_RETRIEVED_SOURCES
    : MAX_RETRIEVED_SOURCES;
  const uniqueResults: LibrarySearchResult[] = [];
  const seenSourceLocations = new Set<string>();
  for (const result of results) {
    // Keep distinct Office sections and OCR pages visible as separate source cards.
    // Plain document chunks remain collapsed to one card per document.
    const sourceLocation = result.sourceLabel
      ? `${result.documentUrl}:section:${result.sourceLabel}`
      : result.pageNumber === undefined
        ? result.documentUrl
        : `${result.documentUrl}:page:${result.pageNumber}`;
    if (seenSourceLocations.has(sourceLocation)) {
      continue;
    }
    seenSourceLocations.add(sourceLocation);
    uniqueResults.push(result);
    if (uniqueResults.length >= MAX_RETRIEVED_SOURCES) {
      break;
    }
  }

  return {
    ...request,
    mode: 'library-rag-search',
    knowledgeScope: 'library-wide-rag',
    selectedFiles: uniqueResults.map(result => {
      const locationLabel = result.sourceLabel || (result.pageNumber === undefined ? '' : `p. ${result.pageNumber}`);
      const locationPrefix = result.sourceLabel || (result.pageNumber === undefined ? '' : `Page ${result.pageNumber} — `);
      return {
        name: `${result.documentName}${locationLabel ? ` · ${locationLabel}` : ''}`,
        url: result.documentUrl,
        fileType: result.fileType,
        libraryTitle: request.libraryName || 'Documents',
        lastModified: result.lastModified,
        snippet: `${locationPrefix}${locationPrefix ? ' — ' : ''}${result.content.slice(0, MAX_CITATION_EXCERPT_CHARS)}`
      };
    }),
    documentSnippets: results.slice(0, maxPromptSources).map(result => {
      const locationContext = result.sourceLabel || (result.pageNumber === undefined ? '' : `page ${result.pageNumber}; `);
      return `Retrieved excerpt from ${result.documentName} (${locationContext ? `${locationContext}; ` : ''}${result.folderPath}):\n${result.content.slice(0, MAX_PROMPT_EXCERPT_CHARS)}`;
    })
  };
}
