import assert from 'node:assert/strict';
import test from 'node:test';
import { buildDeterministicBudgetForecastAnswer, buildDeterministicHba1cAnswer, buildGroundedLibraryRequest, ensureInventoryEvidencePrefix, expandRetrievalQuery, isAbnormalLabResultsQuestion, isDocumentInventoryQuestion, isExplicitHba1cQuestion, isExplicitSpreadsheetSelection, prioritizeDeterministicSpreadsheetAnalysis, requiresDeterministicBudgetAnalysis } from '../src/services/ragAnswerService';

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

test('expands a Korean HbA1c question with the English text-layer PDF terms', () => {
  assert.equal(
    expandRetrievalQuery('이 폴더 문서안에서 당화혈색소 수치가 있으면 알려줘'),
    '이 폴더 문서안에서 당화혈색소 수치가 있으면 알려줘 HbA1c Hemoglobin A1c glycated hemoglobin'
  );
  assert.equal(expandRetrievalQuery('이 폴더의 검사 결과를 요약해줘'), '이 폴더의 검사 결과를 요약해줘');
  assert.equal(
    expandRetrievalQuery('선택된 문서에서 피검사 결과 비정상적인 범위에 있는 수치 알려줘'),
    '선택된 문서에서 피검사 결과 비정상적인 범위에 있는 수치 알려줘 High Low Abnormal Positive Hemoglobin A1c HbA1c'
  );
  assert.equal(isAbnormalLabResultsQuestion('선택된 문서에서 피검사 결과 비정상적인 범위에 있는 수치 알려줘'), true);
  assert.equal(isExplicitHba1cQuestion('당화혈색소 수치를 알려줘'), true);
  assert.equal(isExplicitHba1cQuestion('비정상적인 검사 결과를 알려줘'), false);
});

test('uses the flagged Hemoglobin A1c result instead of the first normal-range boundary', () => {
  const result = buildDeterministicHba1cAnswer('당화혈색소 수치를 알려줘', [{
    id: 'a1c',
    content: 'HEMOGLOBIN A1C Results Hemoglobin A1c Normal range: 4.8 - 5.6 % Prediabetes: 5.7 - 6.4 4.8 4.8 5.6 5.6 6.1 High',
    documentName: 'test-details.pdf', documentUrl: 'https://contoso/test-details.pdf', folderPath: '/Patient Documents', fileType: 'pdf', lastModified: '', chunkOrdinal: 0
  }]);

  assert.equal(result?.answer, 'PDF에 표시된 Hemoglobin A1c 결과는 6.1%이며 High로 표시되어 있습니다. 참고 범위는 4.8–5.6%입니다.');
  assert.equal(result?.source.documentName, 'test-details.pdf');
});

test('labels inventory answers with the confirmed indexed-document count', () => {
  assert.equal(
    ensureInventoryEvidencePrefix('혈액검사 관련 PDF가 확인됩니다.', 5),
    '확인된 5개 인덱싱 문서 기준: 혈액검사 관련 PDF가 확인됩니다.'
  );
  assert.equal(
    ensureInventoryEvidencePrefix('확인된 5개 인덱싱 문서 기준: 혈액검사 관련 PDF가 확인됩니다.', 5),
    '확인된 5개 인덱싱 문서 기준: 혈액검사 관련 PDF가 확인됩니다.'
  );
});

test('library and folder inventory questions use the overview retrieval path', () => {
  assert.equal(isDocumentInventoryQuestion('이 라이브러이에는 주로 어떤 종류의 문서들이 있어?'), true);
  assert.equal(isDocumentInventoryQuestion('이 폴더에는 어떤 파일이 있나요?'), true);
  assert.equal(isDocumentInventoryQuestion('문서 유형별로 어떤 파일들이 확인되나요?'), true);
  assert.equal(isDocumentInventoryQuestion('인덱싱된 파일 목록을 보여줘.'), true);
  assert.equal(isDocumentInventoryQuestion('전자송달 증명서는 언제 제출되었나요?'), false);
});

test('only an explicitly selected workbook enables deterministic spreadsheet handling', () => {
  assert.equal(isExplicitSpreadsheetSelection(undefined, undefined, []), false);
  assert.equal(isExplicitSpreadsheetSelection('', undefined, []), false);
  assert.equal(isExplicitSpreadsheetSelection('xlsx', undefined, []), true);
  assert.equal(isExplicitSpreadsheetSelection(undefined, 'Business expense budget-Excel.xlsx', []), true);
  assert.equal(isExplicitSpreadsheetSelection(undefined, undefined, ['first.pdf', 'budget.xlsx']), true);
});

test('buildGroundedLibraryRequest preserves distinct documents for an inventory overview', () => {
  const results = Array.from({ length: 7 }, (_, index) => ({
    id: `overview-${index}`,
    content: `Document ${index + 1} evidence`,
    documentName: `document-${index + 1}.pdf`,
    documentUrl: `https://contoso/document-${index + 1}.pdf`,
    folderPath: '/Litigation Documents/01_Legal_Official',
    fileType: 'pdf',
    lastModified: '2026-09-04T12:00:00Z',
    chunkOrdinal: index
  }));
  const request = buildGroundedLibraryRequest({ question: '이 폴더에는 어떤 문서가 있나요?', libraryName: 'Litigation Documents' }, results);

  assert.equal(request.mode, 'library-inventory-overview');
  assert.equal(request.selectedFiles?.length, 7);
  assert.equal(request.documentSnippets?.length, 7);
  assert.match(request.documentSnippets?.[0] || '', /7 distinct indexed documents/);
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

test('buildGroundedLibraryRequest omits a null page label from a text-layer PDF citation', () => {
  const request = buildGroundedLibraryRequest(
    { question: 'What does this PDF say?', libraryName: 'Litigation Documents' },
    [{
      id: 'text-pdf', content: 'Text-layer PDF evidence.', documentName: 'text.pdf',
      documentUrl: 'https://youngryu.sharepoint.com/text.pdf', folderPath: '/Litigation Documents',
      fileType: 'pdf', lastModified: '2026-09-03T12:00:00Z', chunkOrdinal: 0,
      pageNumber: null as unknown as number
    }]
  );

  assert.equal(request.selectedFiles?.[0].name, 'text.pdf');
  assert.doesNotMatch(request.selectedFiles?.[0].snippet || '', /Page null/);
  assert.doesNotMatch(request.documentSnippets?.[0] || '', /page null/);
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
