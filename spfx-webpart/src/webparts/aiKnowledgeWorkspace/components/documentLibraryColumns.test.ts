import {
  buildDisplayColumns,
  buildListItemSelect,
  getOrderedColumnSlots,
  buildSelectedDisplayColumns,
  parseDisplayFieldOrder
} from './documentLibraryColumns';

describe('document library column configuration', () => {
  const availableFields = [
    { internalName: 'FileLeafRef', title: 'Name' },
    { internalName: 'Modified', title: 'Modified' },
    { internalName: 'Editor', title: 'Modified By' },
    { internalName: 'CaseNumber', title: 'Case Number' },
    { internalName: 'ClientName', title: 'Client Name' }
  ];

  it('keeps Name first and preserves the editor requested order for visible fields', () => {
    expect(buildDisplayColumns(availableFields, 'ClientName, CaseNumber, Modified')).toEqual([
      { internalName: 'FileLeafRef', title: 'Name' },
      { internalName: 'ClientName', title: 'Client Name' },
      { internalName: 'CaseNumber', title: 'Case Number' },
      { internalName: 'Modified', title: 'Modified' }
    ]);
  });

  it('ignores unavailable and duplicate field names in persisted configuration', () => {
    expect(parseDisplayFieldOrder('CaseNumber,Missing,CaseNumber, FileLeafRef')).toEqual([
      'CaseNumber',
      'Missing',
      'FileLeafRef'
    ]);
    expect(buildDisplayColumns(availableFields, 'CaseNumber,Missing,CaseNumber,FileLeafRef')).toEqual([
      { internalName: 'FileLeafRef', title: 'Name' },
      { internalName: 'CaseNumber', title: 'Case Number' }
    ]);
  });

  it('queries selected internal names alongside fields needed to preserve file behavior', () => {
    expect(buildListItemSelect(['ClientName', 'Modified', 'ClientName'])).toEqual([
      'Id',
      'FileLeafRef',
      'FileRef',
      'FileDirRef',
      'File_x0020_Type',
      'Modified',
      'FSObjType',
      'Editor/Title',
      'ClientName'
    ]);
  });

  it('uses five editor-friendly dropdown slots in their displayed order and ignores duplicates', () => {
    expect(getOrderedColumnSlots(['ClientName', 'CaseNumber', '', 'CaseNumber', 'Modified'])).toEqual([
      'ClientName',
      'CaseNumber',
      'Modified'
    ]);
  });

  it('renders persisted built-in dropdown selections before asynchronous field discovery completes', () => {
    expect(buildSelectedDisplayColumns(
      ['Editor', 'Modified', 'Created', 'File_x0020_Type'],
      []
    )).toEqual([
      { internalName: 'Editor', title: 'Modified By' },
      { internalName: 'Modified', title: 'Modified' },
      { internalName: 'Created', title: 'Created' },
      { internalName: 'File_x0020_Type', title: 'Type' }
    ]);
  });
});
