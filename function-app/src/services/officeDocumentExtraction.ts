import JSZip from 'jszip';

export interface ExtractedOfficeSection {
  text: string;
  citationLabel: string;
}

const XML_ENTITY_MAP: Record<string, string> = {
  amp: '&', lt: '<', gt: '>', quot: '"', apos: "'"
};

function decodeXml(value: string): string {
  return value.replace(/&#(x[0-9a-fA-F]+|\d+);|&([a-zA-Z]+);/g, (_match, numeric: string | undefined, named: string | undefined) => {
    if (numeric) {
      const codePoint = numeric.startsWith('x') ? Number.parseInt(numeric.slice(1), 16) : Number.parseInt(numeric, 10);
      return Number.isFinite(codePoint) ? String.fromCodePoint(codePoint) : '';
    }
    return named ? XML_ENTITY_MAP[named] || '' : '';
  });
}

function textFromXml(xml: string): string {
  return decodeXml(xml
    .replace(/<w:tab\b[^>]*\/>/g, ' ')
    .replace(/<a:br\b[^>]*\/>/g, ' ')
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim());
}

function paragraphsFromWordXml(xml: string): Array<{ text: string; heading?: string }> {
  const paragraphs: Array<{ text: string; heading?: string }> = [];
  for (const match of xml.matchAll(/<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g)) {
    const paragraphXml = match[1];
    const text = textFromXml(paragraphXml);
    if (!text) continue;
    const headingMatch = paragraphXml.match(/<w:pStyle\b[^>]*w:val="([^"]+)"/);
    const style = headingMatch?.[1]?.toLowerCase() || '';
    paragraphs.push({ text, heading: /heading|title/.test(style) ? text : undefined });
  }
  return paragraphs;
}

export async function extractDocxSections(bytes: Buffer): Promise<ExtractedOfficeSection[]> {
  const zip = await JSZip.loadAsync(bytes);
  const document = zip.file('word/document.xml');
  if (!document) throw new Error('DOCX is missing word/document.xml.');
  const paragraphs = paragraphsFromWordXml(await document.async('string'));
  const sections: ExtractedOfficeSection[] = [];
  let heading = 'Document content';
  let content: string[] = [];
  const flush = (): void => {
    const text = content.join(' ').trim();
    if (text) sections.push({ citationLabel: heading, text });
    content = [];
  };
  for (const paragraph of paragraphs) {
    if (paragraph.heading) {
      flush();
      heading = paragraph.heading;
    } else {
      content.push(paragraph.text);
    }
  }
  flush();
  return sections;
}

export async function extractPptxSections(bytes: Buffer): Promise<ExtractedOfficeSection[]> {
  const zip = await JSZip.loadAsync(bytes);
  const slideNames = Object.keys(zip.files)
    .filter(name => /^ppt\/slides\/slide\d+\.xml$/.test(name))
    .sort((a, b) => Number(a.match(/\d+/)?.[0]) - Number(b.match(/\d+/)?.[0]));
  const sections: ExtractedOfficeSection[] = [];
  for (const [index, slideName] of slideNames.entries()) {
    const xml = await zip.file(slideName)?.async('string');
    const text = xml ? textFromXml(xml) : '';
    if (text) sections.push({ citationLabel: `Slide ${index + 1}`, text });
  }
  return sections;
}

function xmlAttribute(xml: string, attribute: string): string | undefined {
  return xml.match(new RegExp(`\\b${attribute}="([^"]*)"`))?.[1];
}

function xlsxTextValue(xml: string): string {
  return decodeXml(Array.from(xml.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)).map(match => match[1]).join(''));
}

function xlsxCellValue(cellXml: string, sharedStrings: string[]): string {
  const type = xmlAttribute(cellXml, 't');
  const formula = cellXml.match(/<f\b[^>]*>([\s\S]*?)<\/f>/)?.[1];
  const value = cellXml.match(/<v\b[^>]*>([\s\S]*?)<\/v>/)?.[1];
  if (type === 'inlineStr') return xlsxTextValue(cellXml);
  if (type === 's' && value !== undefined) return sharedStrings[Number(value)] || '';
  if (value !== undefined) return decodeXml(value).trim();
  return formula ? `Formula: ${decodeXml(formula).trim()}` : '';
}

interface XlsxSheetData {
  name: string;
  rows: Array<Map<number, string>>;
  manualCellColumns: Array<Set<number>>;
  lastRow: number;
}

const MONTH_NAMES = new Map([
  ['jan', 'Jan'], ['january', 'Jan'], ['feb', 'Feb'], ['february', 'Feb'], ['mar', 'Mar'], ['march', 'Mar'],
  ['apr', 'Apr'], ['april', 'Apr'], ['may', 'May'], ['jun', 'Jun'], ['june', 'Jun'], ['jul', 'Jul'], ['july', 'Jul'],
  ['aug', 'Aug'], ['august', 'Aug'], ['sep', 'Sep'], ['sept', 'Sep'], ['september', 'Sep'], ['oct', 'Oct'], ['october', 'Oct'],
  ['nov', 'Nov'], ['november', 'Nov'], ['dec', 'Dec'], ['december', 'Dec']
]);

function columnNumber(cellReference: string): number {
  const letters = cellReference.match(/^[A-Z]+/i)?.[0]?.toUpperCase() || '';
  return [...letters].reduce((result, letter) => result * 26 + letter.charCodeAt(0) - 64, 0);
}

function numericCellValue(value: string | undefined): number | undefined {
  if (!value) return undefined;
  const normalized = value.trim().replace(/[$,\s]/g, '');
  const negative = /^\(.*\)$/.test(normalized);
  const number = Number(negative ? normalized.slice(1, -1) : normalized);
  return Number.isFinite(number) ? (negative ? -number : number) : undefined;
}

function monthColumns(sheet: XlsxSheetData): Map<string, { column: number; row: number }> {
  const months = new Map<string, { column: number; row: number }>();
  sheet.rows.forEach((row, rowIndex) => {
    for (const [column, value] of row) {
      const month = MONTH_NAMES.get(value.trim().toLowerCase());
      if (month && !months.has(month)) months.set(month, { column, row: rowIndex });
    }
  });
  return months;
}

function itemValues(sheet: XlsxSheetData, headerRow: number, firstMonthColumn: number): Map<string, Map<number, number>> {
  const items = new Map<string, Map<number, number>>();
  for (let rowIndex = headerRow + 1; rowIndex < sheet.rows.length; rowIndex += 1) {
    const row = sheet.rows[rowIndex];
    const label = [...row.entries()].find(([column, value]) => column < firstMonthColumn && value.trim())?.[1]?.trim();
    if (!label || /^(subtotal|totals?)$/i.test(label) || /^monthly\b/i.test(label) || /^total\b/i.test(label)) continue;
    const values = new Map<number, number>();
    for (const [column, value] of row) {
      const number = numericCellValue(value);
      if (number !== undefined) values.set(column, number);
    }
    if (values.size) items.set(label.toLowerCase(), values);
  }
  return items;
}

function hasEnteredNumericValue(sheet: XlsxSheetData, headerRow: number, column: number): boolean {
  for (let rowIndex = headerRow + 1; rowIndex < sheet.rows.length; rowIndex += 1) {
    if (!sheet.manualCellColumns[rowIndex]?.has(column)) continue;
    if (numericCellValue(sheet.rows[rowIndex].get(column)) !== undefined) return true;
  }
  return false;
}

interface XlsxExpenseCategory {
  name: string;
  items: Map<string, Map<number, number>>;
}

function hasMonthHeader(row: Map<number, string>): boolean {
  return [...row.values()].some(value => MONTH_NAMES.has(value.trim().toLowerCase()));
}

function expenseCategories(sheet: XlsxSheetData, headerRow: number, firstMonthColumn: number): Map<string, XlsxExpenseCategory> {
  const categories = new Map<string, XlsxExpenseCategory>();
  let current: XlsxExpenseCategory | undefined;
  for (let rowIndex = headerRow; rowIndex < sheet.rows.length; rowIndex += 1) {
    const row = sheet.rows[rowIndex];
    const label = [...row.entries()].find(([column, value]) => column < firstMonthColumn && value.trim())?.[1]?.trim();
    if (!label) continue;
    if (hasMonthHeader(row)) {
      current = /^totals?$/i.test(label) ? undefined : { name: label, items: new Map() };
      if (current) categories.set(label.toLowerCase(), current);
      continue;
    }
    if (!current || /^(subtotal|totals?)$/i.test(label) || /^monthly\b/i.test(label) || /^total\b/i.test(label)) continue;
    const values = new Map<number, number>();
    for (const [column, value] of row) {
      const number = numericCellValue(value);
      if (number !== undefined) values.set(column, number);
    }
    if (values.size) current.items.set(label.toLowerCase(), values);
  }
  return categories;
}

function categoryForecastLines(
  planned: XlsxSheetData,
  actual: XlsxSheetData,
  plannedHeaderRow: number,
  actualHeaderRow: number,
  plannedMonths: Map<string, { column: number; row: number }>,
  sharedMonths: Array<[string, { column: number; row: number }]>
): string[] {
  const plannedCategories = expenseCategories(planned, plannedHeaderRow, Math.min(...[...plannedMonths.values()].map(value => value.column)));
  const actualCategories = expenseCategories(actual, actualHeaderRow, Math.min(...sharedMonths.map(([, value]) => value.column)));
  const lines: string[] = [];
  for (const [key, actualCategory] of actualCategories) {
    const plannedCategory = plannedCategories.get(key);
    if (!plannedCategory) continue;
    let annualPlan = 0;
    let plannedToDate = 0;
    let actualToDate = 0;
    for (const [item, actualValues] of actualCategory.items) {
      const plannedValues = plannedCategory.items.get(item);
      if (!plannedValues) continue;
      for (const { column } of plannedMonths.values()) annualPlan += plannedValues.get(column) || 0;
      for (const [month, actualColumn] of sharedMonths) {
        plannedToDate += plannedValues.get(plannedMonths.get(month)!.column) || 0;
        actualToDate += actualValues.get(actualColumn.column) || 0;
      }
    }
    if (!annualPlan && !actualToDate) continue;
    const annualizedActual = actualToDate / sharedMonths.length * plannedMonths.size;
    lines.push(`${actualCategory.name}: planned for populated months ${formatAmount(plannedToDate)}; actual for populated months ${formatAmount(actualToDate)}; populated-month variance ${formatAmount(actualToDate - plannedToDate)}; annual plan ${formatAmount(annualPlan)}; annualized actual ${formatAmount(annualizedActual)}; forecast variance ${formatAmount(annualizedActual - annualPlan)}.`);
  }
  return lines;
}

function formatAmount(value: number): string {
  return value.toFixed(2);
}

function budgetVarianceSection(sheets: XlsxSheetData[]): ExtractedOfficeSection | undefined {
  const planned = sheets.find(sheet => /planned.*expenses?/i.test(sheet.name));
  const actual = sheets.find(sheet => /actual.*expenses?/i.test(sheet.name));
  if (!planned || !actual) return undefined;

  const plannedMonths = monthColumns(planned);
  const actualMonths = monthColumns(actual);
  const candidateMonths = [...actualMonths.entries()].filter(([month]) => plannedMonths.has(month));
  if (!candidateMonths.length) return undefined;

  const plannedHeaderRow = Math.min(...[...plannedMonths.values()].map(value => value.row));
  const actualHeaderRow = Math.min(...[...actualMonths.values()].map(value => value.row));
  const sharedMonths = candidateMonths.filter(([, actualColumn]) => hasEnteredNumericValue(actual, actualHeaderRow, actualColumn.column));
  if (!sharedMonths.length) return undefined;
  const plannedItems = itemValues(planned, plannedHeaderRow, Math.min(...[...plannedMonths.values()].map(value => value.column)));
  const actualItems = itemValues(actual, actualHeaderRow, Math.min(...[...actualMonths.values()].map(value => value.column)));
  const lines: string[] = [];
  let totalPlanned = 0;
  let totalActual = 0;
  let largest: { item: string; month: string; planned: number; actual: number; variance: number } | undefined;

  for (const [month, actualColumn] of sharedMonths) {
    const plannedColumn = plannedMonths.get(month)!;
    let monthlyPlanned = 0;
    let monthlyActual = 0;
    let populated = false;
    for (const [item, actualValues] of actualItems) {
      const actualValue = actualValues.get(actualColumn.column);
      const plannedValue = plannedItems.get(item)?.get(plannedColumn.column);
      if (actualValue === undefined || plannedValue === undefined) continue;
      populated = true;
      monthlyPlanned += plannedValue;
      monthlyActual += actualValue;
      const variance = actualValue - plannedValue;
      if (!largest || Math.abs(variance) > Math.abs(largest.variance)) {
        largest = { item, month, planned: plannedValue, actual: actualValue, variance };
      }
    }
    if (!populated) continue;
    totalPlanned += monthlyPlanned;
    totalActual += monthlyActual;
    lines.push(`${month}: planned ${formatAmount(monthlyPlanned)}; actual ${formatAmount(monthlyActual)}; variance ${formatAmount(monthlyActual - monthlyPlanned)}.`);
  }
  if (!lines.length || !largest) return undefined;

  const categoryForecasts = categoryForecastLines(planned, actual, plannedHeaderRow, actualHeaderRow, plannedMonths, sharedMonths);
  const displayItem = [...actualItems.keys()].find(item => item === largest!.item) || largest.item;
  return {
    citationLabel: `Budget variance analysis: ${planned.name} vs ${actual.name}`,
    text: [
      `Populated actual months: ${lines.map(line => line.slice(0, 3)).join(', ')}.`,
      ...lines,
      `Total for populated months: planned ${formatAmount(totalPlanned)}; actual ${formatAmount(totalActual)}; variance ${formatAmount(totalActual - totalPlanned)}.`,
      `Largest item variance: ${displayItem.replace(/\b\w/g, letter => letter.toUpperCase())} in ${largest.month}; planned ${formatAmount(largest.planned)}; actual ${formatAmount(largest.actual)}; variance ${formatAmount(largest.variance)}.`,
      ...(categoryForecasts.length ? [`Annualized category forecast using the populated-month run rate: ${categoryForecasts.join(' ')}`] : [])
    ].join(' ')
  };
}

export async function extractXlsxSections(bytes: Buffer): Promise<ExtractedOfficeSection[]> {
  const zip = await JSZip.loadAsync(bytes);
  const workbook = zip.file('xl/workbook.xml');
  const relationships = zip.file('xl/_rels/workbook.xml.rels');
  if (!workbook || !relationships) throw new Error('XLSX is missing workbook metadata.');
  const sharedStringsXml = await zip.file('xl/sharedStrings.xml')?.async('string');
  const sharedStrings = sharedStringsXml
    ? Array.from(sharedStringsXml.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)).map(match => xlsxTextValue(match[1]))
    : [];
  const relationshipTargets = new Map<string, string>();
  for (const match of (await relationships.async('string')).matchAll(/<Relationship\b[^>]*>/g)) {
    const id = xmlAttribute(match[0], 'Id');
    const target = xmlAttribute(match[0], 'Target');
    if (id && target) {
      const normalizedTarget = target.replace(/^\//, '');
      relationshipTargets.set(id, normalizedTarget.startsWith('xl/') ? normalizedTarget : `xl/${normalizedTarget}`);
    }
  }
  const sections: ExtractedOfficeSection[] = [];
  const sheets: XlsxSheetData[] = [];
  for (const match of (await workbook.async('string')).matchAll(/<sheet\b[^>]*\/>/g)) {
    const sheetXml = match[0];
    if (xmlAttribute(sheetXml, 'state') === 'hidden' || xmlAttribute(sheetXml, 'state') === 'veryHidden') continue;
    const name = decodeXml(xmlAttribute(sheetXml, 'name') || 'Worksheet');
    const sheetPath = relationshipTargets.get(xmlAttribute(sheetXml, 'r:id') || '');
    const worksheet = sheetPath ? zip.file(sheetPath) : undefined;
    if (!worksheet) continue;
    const rows: string[] = [];
    const sheetRows: Array<Map<number, string>> = [];
    const manualCellColumns: Array<Set<number>> = [];
    let lastRow = 0;
    for (const rowMatch of (await worksheet.async('string')).matchAll(/<row\b[^>]*>([\s\S]*?)<\/row>/g)) {
      const rowNumber = Number(xmlAttribute(rowMatch[0], 'r')) || lastRow + 1;
      const cells: string[] = [];
      const sheetRow = new Map<number, string>();
      const manualColumns = new Set<number>();
      // Match self-closing empty cells separately. A paired-tag-only regex would
      // start at an empty July cell and consume through a later YEAR total cell.
      for (const cellMatch of rowMatch[1].matchAll(/<c\b[^>]*\/>|<c\b[^>]*>[\s\S]*?<\/c>/g)) {
        const cellXml = cellMatch[0];
        const cellRef = xmlAttribute(cellXml, 'r') || '';
        const value = xlsxCellValue(cellXml, sharedStrings);
        if (value) {
          cells.push(`${cellRef}: ${value}`);
          const column = columnNumber(cellRef);
          sheetRow.set(column, value);
          if (!/<f\b[^>]*>/.test(cellXml)) manualColumns.add(column);
        }
      }
      if (cells.length) rows.push(cells.join(' | '));
      sheetRows.push(sheetRow);
      manualCellColumns.push(manualColumns);
      lastRow = rowNumber;
    }
    if (rows.length) {
      sections.push({ citationLabel: `Sheet: ${name} · Rows 1–${lastRow}`, text: rows.join('\n') });
      sheets.push({ name, rows: sheetRows, manualCellColumns, lastRow });
    }
  }
  const analysis = budgetVarianceSection(sheets);
  if (analysis) sections.push(analysis);
  return sections;
}

export async function extractOfficeSections(fileType: string, bytes: Buffer): Promise<ExtractedOfficeSection[]> {
  switch (fileType.toLowerCase()) {
    case 'docx': return extractDocxSections(bytes);
    case 'pptx': return extractPptxSections(bytes);
    case 'xlsx': return extractXlsxSections(bytes);
    default: throw new Error(`Unsupported Office file type: ${fileType}.`);
  }
}
