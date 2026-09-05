import { buildRenameEndpoint, validateRenameName } from './renameItem';

describe('rename item helpers', () => {
  it('accepts a trimmed valid name that differs from the current name', () => {
    expect(validateRenameName('  Final contract.pdf  ', 'Draft contract.pdf')).toEqual({ valid: true, value: 'Final contract.pdf' });
  });

  it('rejects blank, unchanged, invalid, and overlong names', () => {
    expect(validateRenameName('   ', 'Report.docx')).toEqual({ valid: false, error: 'Enter a name.' });
    expect(validateRenameName('Report.docx', 'Report.docx')).toEqual({ valid: false, error: 'Enter a different name.' });
    expect(validateRenameName('Bad/Name.docx', 'Report.docx')).toEqual({ valid: false, error: 'Names cannot contain \\ / : * ? " < > | # % { } ~ &.' });
    expect(validateRenameName('a'.repeat(129), 'Report.docx')).toEqual({ valid: false, error: 'Names must be 128 characters or fewer.' });
  });

  it('builds a non-overwriting file MoveTo request and escapes OData apostrophes', () => {
    expect(buildRenameEndpoint('https://contoso.sharepoint.com/sites/legal', 'file', "/sites/legal/Documents/O'Brien.docx", "Final O'Brien.docx")).toBe(
      "https://contoso.sharepoint.com/sites/legal/_api/web/GetFileByServerRelativeUrl('/sites/legal/Documents/O''Brien.docx')/MoveTo(newurl='/sites/legal/Documents/Final O''Brien.docx',flags=0)"
    );
  });

  it('builds a non-overwriting folder MoveTo request', () => {
    expect(buildRenameEndpoint('https://contoso.sharepoint.com/sites/legal/', 'folder', '/sites/legal/Documents/Discovery', 'Evidence')).toBe(
      "https://contoso.sharepoint.com/sites/legal/_api/web/GetFolderByServerRelativeUrl('/sites/legal/Documents/Discovery')/MoveTo(newurl='/sites/legal/Documents/Evidence',flags=0)"
    );
  });
});
