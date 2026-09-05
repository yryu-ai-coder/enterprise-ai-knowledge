import { getFileIconKind } from './fileIcon';

describe('file icon presentation', () => {
  it('uses the dedicated PDF icon for case-insensitive PDF file types', () => {
    expect(getFileIconKind('PDF')).toBe('pdf');
  });

  it.each([
    ['DOCX', 'word'],
    ['xlsx', 'excel'],
    ['PPTX', 'powerpoint']
  ])('uses the current Microsoft-style %s application icon', (fileType, expectedKind) => {
    expect(getFileIconKind(fileType)).toBe(expectedKind);
  });

  it('keeps the generic file icon for an unknown file type', () => {
    expect(getFileIconKind('zip')).toBe('generic');
  });
});
