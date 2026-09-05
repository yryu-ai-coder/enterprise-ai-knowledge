import assert from 'node:assert/strict';
import test from 'node:test';
import { buildDeterministicBudgetForecastAnswer, buildGroundedLibraryRequest, prioritizeDeterministicSpreadsheetAnalysis, requiresDeterministicBudgetAnalysis } from '../src/services/ragAnswerService';

test('buildGroundedLibraryRequest converts retrieved chunks into bounded source-linked chat context', () => {
  const request = buildGroundedLibraryRequest(
    {
      question: 'Find Ava academic plan communications.',
      siteUrl: 'https://youngryu.sharepoint.com/sites/enterprise-ai-knowledge',
      libraryName: 'Litigation Documents'
    },
    [
      {
        id: 'chunk-one',
        content: 'Ava discussed the independent study plan and a September deadline.',
        documentName: 'consultation.pdf',
        documentUrl: 'https://youngryu.sharepoint.com/consultation.pdf',
        folderPath: '/Litigation Documents',
        fileType: 'pdf',
        lastModified: '2026-09-03T12:00:00Z',
        chunkOrdinal: 0
      }
    ]
  );

  assert.equal(request.knowledgeScope, 'library-wide-rag');
  assert.equal(request.selectedFiles?.[0].name, 'consultation.pdf');
  assert.match(request.selectedFiles?.[0].snippet || '', /September deadline/);
  assert.match(request.documentSnippets?.[0] || '', /Retrieved excerpt from consultation\.pdf/);
});

test('buildGroundedLibraryRequest preserves an OCR source page in its citation-facing source label', () => {
  const request = buildGroundedLibraryRequest(
    { question: 'What is on the scanned page?', libraryName: 'Litigation Documents' },
    [{
      id: 'ocr-page-3', content: 'Scanned printed text.', documentName: 'scan.pdf',
      documentUrl: 'https://youngryu.sharepoint.com/scan.pdf', folderPath: '/Litigation Documents',
      fileType: 'pdf', lastModified: '2026-09-03T12:00:00Z', chunkOrdinal: 2, pageNumber: 3
    }]
  );

  assert.equal(request.selectedFiles?.[0].name, 'scan.pdf · p. 3');
  assert.match(request.selectedFiles?.[0].snippet || '', /^Page 3 —/);
  assert.match(request.documentSnippets?.[0] || '', /scan\.pdf \(page 3;/);
});

test('buildGroundedLibraryRequest retains section, slide, and worksheet citations', () => {
  const request = buildGroundedLibraryRequest(
    { question: 'What is the conclusion?', libraryName: 'Documents' },
    [{
      id: 'deck-slide-2', content: 'The approved budget is $1200.', documentName: 'case-update.pptx',
      documentUrl: 'https://youngryu.sharepoint.com/case-update.pptx', folderPath: '/Documents',
      fileType: 'pptx', lastModified: '2026-09-04T12:00:00Z', chunkOrdinal: 1000, sourceLabel: 'Slide 2'
    }]
  );

  assert.equal(request.selectedFiles?.[0].name, 'case-update.pptx · Slide 2');
  assert.match(request.selectedFiles?.[0].snippet || '', /^Slide 2 —/);
  assert.match(request.documentSnippets?.[0] || '', /Slide 2;/);
});

test('buildGroundedLibraryRequest retains distinct section citations from the same Office document', () => {
  const request = buildGroundedLibraryRequest(
    { question: 'Summarize this newsletter.', libraryName: 'Documents' },
    [
      { id: 'word-section-one', content: 'Opening material.', documentName: 'review.docx', documentUrl: 'https://contoso/review.docx', folderPath: '/Documents', fileType: 'docx', lastModified: '2026-09-04T12:00:00Z', chunkOrdinal: 0, sourceLabel: 'The Review' },
      { id: 'word-section-two', content: 'Newsletter findings.', documentName: 'review.docx', documentUrl: 'https://contoso/review.docx', folderPath: '/Documents', fileType: 'docx', lastModified: '2026-09-04T12:00:00Z', chunkOrdinal: 1000, sourceLabel: 'New finds this week' }
    ]
  );

  assert.deepEqual(request.selectedFiles?.map(file => file.name), [
    'review.docx · The Review',
    'review.docx · New finds this week'
  ]);
});

test('buildGroundedLibraryRequest retains broader worksheet evidence for a spreadsheet calculation', () => {
  const chunks = Array.from({ length: 8 }, (_, index) => ({
    id: `budget-${index}`,
    content: `Budget evidence chunk ${index + 1}`,
    documentName: 'budget.xlsx',
    documentUrl: 'https://contoso/budget.xlsx',
    folderPath: '/Documents',
    fileType: 'xlsx',
    lastModified: '2026-09-04T12:00:00Z',
    chunkOrdinal: index,
    sourceLabel: 'Sheet: Budget · Rows 1–39'
  }));
  const request = buildGroundedLibraryRequest({ question: 'Which month has the largest variance?', libraryName: 'Documents' }, chunks);

  assert.equal(request.documentSnippets?.length, 8);
  assert.equal(request.selectedFiles?.length, 1);
});

test('prioritizeDeterministicSpreadsheetAnalysis excludes raw worksheet chunks when a budget-analysis section is retrieved', () => {
  const rawWorksheet = {
    id: 'raw-worksheet', content: 'Apr collateral preparation: planned 5000; actual 5500.',
    documentName: 'budget.xlsx', documentUrl: 'https://contoso/budget.xlsx', folderPath: '/Documents',
    fileType: 'xlsx', lastModified: '2026-09-04T12:00:00Z', chunkOrdinal: 0, sourceLabel: 'Sheet: Actual expenses · Rows 1–42'
  };
  const deterministicAnalysis = {
    id: 'budget-analysis', content: 'Largest item variance: Training-related travel costs in Jun; planned 2000.00; actual 3500.00; variance 1500.00.',
    documentName: 'budget.xlsx', documentUrl: 'https://contoso/budget.xlsx', folderPath: '/Documents',
    fileType: 'xlsx', lastModified: '2026-09-04T12:00:00Z', chunkOrdinal: 5000, sourceLabel: 'Budget variance analysis: Planned expenses vs Actual expenses'
  };

  assert.deepEqual(
    prioritizeDeterministicSpreadsheetAnalysis([rawWorksheet, deterministicAnalysis]).map(result => result.id),
    ['budget-analysis']
  );
});

test('requiresDeterministicBudgetAnalysis recognizes Korean planned-versus-actual questions', () => {
  assert.equal(requiresDeterministicBudgetAnalysis('계획 대비 실제 지출이 가장 크게 초과된 항목은 무엇인가요?'), true);
  assert.equal(requiresDeterministicBudgetAnalysis('이 파일의 작성자는 누구인가요?'), false);
});

test('buildDeterministicBudgetForecastAnswer uses category labels paired with their own forecast values', () => {
  const answer = buildDeterministicBudgetForecastAnswer('6월까지의 실제 추세가 유지되면 연간 예산을 초과할 범주는 무엇인가요?', [{
    id: 'analysis', documentName: 'budget.xlsx', documentUrl: 'https://contoso/budget.xlsx', folderPath: '/Documents', fileType: 'xlsx', lastModified: '', chunkOrdinal: 1,
    sourceLabel: 'Budget variance analysis: Planned expenses vs Actual expenses',
    content: 'Annualized category forecast using the populated-month run rate: Employee costs: planned for populated months 657225.00; actual for populated months 659130.00; populated-month variance 1905.00; annual plan 1355090.00; annualized actual 1318260.00; forecast variance -36830.00. Office costs: planned for populated months 69620.00; actual for populated months 69350.00; populated-month variance -270.00; annual plan 139040.00; annualized actual 138700.00; forecast variance -340.00. Marketing costs: planned for populated months 32400.00; actual for populated months 33159.00; populated-month variance 759.00; annual plan 67800.00; annualized actual 66318.00; forecast variance -1482.00. Training/travel: planned for populated months 24000.00; actual for populated months 21300.00; populated-month variance -2700.00; annual plan 48000.00; annualized actual 42600.00; forecast variance -5400.00.'
  }]);

  assert.match(answer || '', /연간 예산 초과가 예상되는 범주는 없습니다/);
  assert.match(answer || '', /Office costs/);
  assert.match(answer || '', /139,040/);
  assert.match(answer || '', /138,700/);
  assert.match(answer || '', /-340/);
  assert.doesNotMatch(answer || '', /Marketing costs.*48,000/);
});
