import { getSelectionCommandState, toggleSelection } from './librarySelection';

describe('library selection helpers', () => {
  it('adds and removes files independently without clearing other selections', () => {
    const firstSelection = toggleSelection(['https://contoso/sites/legal/one.pdf'], 'https://contoso/sites/legal/two.pdf');
    expect(firstSelection).toEqual(['https://contoso/sites/legal/one.pdf', 'https://contoso/sites/legal/two.pdf']);

    expect(toggleSelection(firstSelection, 'https://contoso/sites/legal/one.pdf')).toEqual(['https://contoso/sites/legal/two.pdf']);
  });

  it('enables preview and rename only for one selected file', () => {
    expect(getSelectionCommandState(['https://contoso/one.pdf'], [])).toEqual({
      total: 1,
      selectedFileUrl: 'https://contoso/one.pdf',
      selectedFolderPath: '',
      canPreview: true,
      canRename: true,
      canDelete: true,
      previewTitle: 'Preview selected file',
      renameTitle: 'Rename selected item',
      deleteTitle: 'Delete selected item'
    });
  });

  it('allows a single folder to rename and delete but not preview', () => {
    expect(getSelectionCommandState([], ['/sites/legal/Documents/Evidence'])).toEqual({
      total: 1,
      selectedFileUrl: '',
      selectedFolderPath: '/sites/legal/Documents/Evidence',
      canPreview: false,
      canRename: true,
      canDelete: true,
      previewTitle: 'Select exactly one file to preview',
      renameTitle: 'Rename selected item',
      deleteTitle: 'Delete selected item'
    });
  });

  it('disables preview and rename while permitting batch delete for multiple files or folders', () => {
    expect(getSelectionCommandState(['https://contoso/one.pdf'], ['/sites/legal/Documents/Evidence'])).toEqual({
      total: 2,
      selectedFileUrl: '',
      selectedFolderPath: '',
      canPreview: false,
      canRename: false,
      canDelete: true,
      previewTitle: 'Select exactly one file to preview',
      renameTitle: 'Select exactly one file or folder to rename',
      deleteTitle: 'Move 2 selected items to the SharePoint recycle bin'
    });
  });
});
