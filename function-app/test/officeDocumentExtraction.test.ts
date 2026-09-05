import assert from 'node:assert/strict';
import test from 'node:test';
import ExcelJS from 'exceljs';
import JSZip from 'jszip';
import { extractDocxSections, extractOfficeSections, extractPptxSections, extractXlsxSections } from '../src/services/officeDocumentExtraction';

async function zipBytes(files: Record<string, string>): Promise<Buffer> {
  const zip = new JSZip();
  Object.entries(files).forEach(([name, value]) => zip.file(name, value));
  return Buffer.from(await zip.generateAsync({ type: 'nodebuffer' }));
}

test('extractDocxSections keeps Word heading boundaries as citation labels', async () => {
  const bytes = await zipBytes({
    'word/document.xml': `<?xml version="1.0"?><w:document xmlns:w="w"><w:body>
      <w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Case Background</w:t></w:r></w:p>
      <w:p><w:r><w:t>The agreement was signed on Monday.</w:t></w:r></w:p>
      <w:p><w:pPr><w:pStyle w:val="Heading2"/></w:pPr><w:r><w:t>Key Terms</w:t></w:r></w:p>
      <w:p><w:r><w:t>The term is twelve months.</w:t></w:r></w:p>
    </w:body></w:document>`
  });
  const sections = await extractDocxSections(bytes);
  assert.deepEqual(sections, [
    { citationLabel: 'Case Background', text: 'The agreement was signed on Monday.' },
    { citationLabel: 'Key Terms', text: 'The term is twelve months.' }
  ]);
});

test('extractPptxSections returns one citation-aware section per slide', async () => {
  const bytes = await zipBytes({
    'ppt/slides/slide2.xml': '<p:sld xmlns:p="p"><a:t xmlns:a="a">Second slide conclusion</a:t></p:sld>',
    'ppt/slides/slide1.xml': '<p:sld xmlns:p="p"><a:t xmlns:a="a">Executive summary</a:t></p:sld>'
  });
  assert.deepEqual(await extractPptxSections(bytes), [
    { citationLabel: 'Slide 1', text: 'Executive summary' },
    { citationLabel: 'Slide 2', text: 'Second slide conclusion' }
  ]);
});

test('extractXlsxSections uses visible sheet and row range citations', async () => {
  const workbook = new ExcelJS.Workbook();
  const budget = workbook.addWorksheet('Budget');
  budget.addRow(['Category', 'Amount']);
  budget.addRow(['Discovery', 1200]);
  workbook.addWorksheet('Hidden', { state: 'hidden' }).addRow(['Do not index']);
  const bytes = Buffer.from(await workbook.xlsx.writeBuffer());
  const sections = await extractXlsxSections(bytes);
  assert.equal(sections.length, 1);
  assert.equal(sections[0].citationLabel, 'Sheet: Budget · Rows 1–2');
  assert.match(sections[0].text, /A2: Discovery/);
  assert.match(sections[0].text, /B2: 1200/);
});

test('extractXlsxSections adds deterministic planned-versus-actual variance for populated months', async () => {
  const workbook = new ExcelJS.Workbook();
  const planned = workbook.addWorksheet('Planned expenses');
  planned.addRow(['Expense', 'Jan', 'Feb', 'Mar', 'YEAR']);
  planned.addRow(['Wages', 100, 100, 100, 300]);
  planned.addRow(['Rent', 50, 50, 50, 150]);

  const actual = workbook.addWorksheet('Actual expenses');
  actual.addRow(['Expense', 'Jan', 'Feb', 'Mar', 'YEAR']);
  actual.addRow(['Wages', 120, undefined, 80, 999]);
  actual.addRow(['Rent', 110, undefined, 110, 999]);

  const sections = await extractXlsxSections(Buffer.from(await workbook.xlsx.writeBuffer()));
  const analysis = sections.find(section => section.citationLabel === 'Budget variance analysis: Planned expenses vs Actual expenses');

  assert.ok(analysis);
  assert.match(analysis.text, /Populated actual months: Jan, Mar\./);
  assert.match(analysis.text, /Jan: planned 150\.00; actual 230\.00; variance 80\.00\./);
  assert.match(analysis.text, /Mar: planned 150\.00; actual 190\.00; variance 40\.00\./);
  assert.match(analysis.text, /Total for populated months: planned 300\.00; actual 420\.00; variance 120\.00\./);
  assert.match(analysis.text, /Largest item variance: Rent in Jan; planned 50\.00; actual 110\.00; variance 60\.00\./);
  assert.match(analysis.text, /Annualized category forecast using the populated-month run rate: Expense: planned for populated months 300\.00; actual for populated months 420\.00; populated-month variance 120\.00; annual plan 450\.00; annualized actual 630\.00; forecast variance 180\.00\./);
  assert.doesNotMatch(analysis.text, /YEAR|999|Feb/);
});

test('extractXlsxSections excludes subtotal rows and formula-only future months from budget variance analysis', async () => {
  const workbook = new ExcelJS.Workbook();
  const planned = workbook.addWorksheet('Planned expenses');
  planned.addRow(['Expense', 'Jan', 'Feb', 'YEAR']);
  planned.addRow(['Wages', 100, 100, 200]);
  planned.addRow(['Benefits', 20, 20, 40]);
  planned.addRow(['Subtotal', 120, 120, 240]);

  const actual = workbook.addWorksheet('Actual expenses');
  actual.addRow(['Expense', 'Jan', 'Feb', 'YEAR']);
  actual.addRow(['Wages', 110, undefined, 110]);
  actual.addRow(['Benefits', { formula: 'B2*0.2', result: 22 }, { formula: 'C2*0.2', result: 0 }, { formula: 'SUM(B3:C3)', result: 22 }]);
  actual.addRow(['Subtotal', { formula: 'SUM(B2:B3)', result: 132 }, { formula: 'SUM(C2:C3)', result: 0 }, { formula: 'SUM(D2:D3)', result: 132 }]);

  const sections = await extractXlsxSections(Buffer.from(await workbook.xlsx.writeBuffer()));
  const analysis = sections.find(section => section.citationLabel === 'Budget variance analysis: Planned expenses vs Actual expenses');

  assert.ok(analysis);
  assert.match(analysis.text, /Populated actual months: Jan\./);
  assert.match(analysis.text, /Jan: planned 120\.00; actual 132\.00; variance 12\.00\./);
  assert.doesNotMatch(analysis.text, /Feb|Subtotal/);
});

test('extractOfficeSections rejects an unsupported file type', async () => {
  await assert.rejects(extractOfficeSections('doc', Buffer.from('not office')), /Unsupported Office/i);
});

test('extractXlsxSections handles a SharePoint-style workbook relationship with an absolute xl target', async () => {
  const bytes = await zipBytes({
    'xl/workbook.xml': '<workbook><sheets><sheet name="Budget" sheetId="1" r:id="rId1" xmlns:r="r"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Target="/xl/worksheets/sheet1.xml"/></Relationships>',
    'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Item</t></is></c><c r="B1" t="inlineStr"><is><t>Amount</t></is></c></row><row r="2"><c r="A2" t="inlineStr"><is><t>Discovery</t></is></c><c r="B2"><v>12500</v></c></row></sheetData></worksheet>'
  });
  const sections = await extractXlsxSections(bytes);
  assert.deepEqual(sections, [{
    citationLabel: 'Sheet: Budget · Rows 1–2',
    text: 'A1: Item | B1: Amount\nA2: Discovery | B2: 12500'
  }]);
});

test('extractXlsxSections does not attach a year total to the first blank future-month cell', async () => {
  const bytes = await zipBytes({
    'xl/workbook.xml': '<workbook><sheets><sheet name="Actual expenses" sheetId="1" r:id="rId1" xmlns:r="r"/></sheets></workbook>',
    'xl/_rels/workbook.xml.rels': '<Relationships><Relationship Id="rId1" Target="worksheets/sheet1.xml"/></Relationships>',
    'xl/worksheets/sheet1.xml': '<worksheet><sheetData><row r="6"><c r="C6" t="inlineStr"><is><t>Jan</t></is></c><c r="I6" t="inlineStr"><is><t>Jul</t></is></c><c r="O6" t="inlineStr"><is><t>YEAR</t></is></c></row><row r="7"><c r="B7" t="inlineStr"><is><t>Wages</t></is></c><c r="C7"><v>85000</v></c><c r="D7"><v>85000</v></c><c r="I7"/><c r="J7"/><c r="K7"/><c r="L7"/><c r="M7"/><c r="N7"/><c r="O7"><f>SUM(C7:N7)</f><v>519000</v></c></row></sheetData></worksheet>'
  });
  const [section] = await extractXlsxSections(bytes);
  assert.match(section.text, /O7: 519000/);
  assert.doesNotMatch(section.text, /I7: 519000/);
});
