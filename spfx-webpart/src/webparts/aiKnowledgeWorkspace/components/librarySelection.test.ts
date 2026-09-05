import { getFolderOpenSelection, getSelectionCommandState, toggleSelection } from './librarySelection';

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
      canDownload: true,
      canRename: true,
      canDelete: true,
      previewTitle: 'Preview selected file',
      downloadTitle: 'Download selected file',
      renameTitle: 'Rename selected item',
      deleteTitle: 'Delete selected item'
    });
  });

  it('does not select a file when opening a folder', () => {
    expect(getFolderOpenSelection()).toEqual({ selectedFileUrl: '', selectedFileUrls: [], selectedFolderPaths: [] });
  });

  it('allows a single folder to rename and delete but not preview', () => {
    expect(getSelectionCommandState([], ['/sites/legal/Documents/Evidence'])).toEqual({
      total: 1,
      selectedFileUrl: '',
      selectedFolderPath: '/sites/legal/Documents/Evidence',
      canPreview: false,
      canDownload: false,
      canRename: true,
      canDelete: true,
      previewTitle: 'Select exactly one file to preview',
      downloadTitle: 'Select one or more files to download',
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
      canDownload: true,
      canRename: false,
      canDelete: true,
      previewTitle: 'Select exactly one file to preview',
      downloadTitle: 'Download selected file',
      renameTitle: 'Select exactly one file or folder to rename',
      deleteTitle: 'Move 2 selected items to the SharePoint recycle bin'
    });
  });
});
