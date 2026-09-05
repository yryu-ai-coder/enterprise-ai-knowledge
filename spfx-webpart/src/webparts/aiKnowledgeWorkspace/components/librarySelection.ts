export interface ISelectionCommandState {
  total: number;
  selectedFileUrl: string;
  selectedFolderPath: string;
  canPreview: boolean;
  canDownload: boolean;
  canRename: boolean;
  canDelete: boolean;
  previewTitle: string;
  downloadTitle: string;
  renameTitle: string;
  deleteTitle: string;
}

export interface IFolderOpenSelection {
  selectedFileUrl: string;
  selectedFileUrls: string[];
  selectedFolderPaths: string[];
}

export function getFolderOpenSelection(): IFolderOpenSelection {
  return { selectedFileUrl: '', selectedFileUrls: [], selectedFolderPaths: [] };
}

export function toggleSelection(selection: string[], item: string): string[] {
  return selection.indexOf(item) >= 0
    ? selection.filter((selectedItem) => selectedItem !== item)
    : [...selection, item];
}

export function getSelectionCommandState(selectedFileUrls: string[], selectedFolderPaths: string[]): ISelectionCommandState {
  const total = selectedFileUrls.length + selectedFolderPaths.length;
  const hasExactlyOneSelection = total === 1;
  const selectedFileUrl = hasExactlyOneSelection && selectedFileUrls.length === 1 ? selectedFileUrls[0] : '';
  const selectedFolderPath = hasExactlyOneSelection && selectedFolderPaths.length === 1 ? selectedFolderPaths[0] : '';
  const canPreview = !!selectedFileUrl;
  const canDownload = selectedFileUrls.length > 0;
  const canRename = hasExactlyOneSelection;
  const canDelete = total > 0;

  return {
    total,
    selectedFileUrl,
    selectedFolderPath,
    canPreview,
    canDownload,
    canRename,
    canDelete,
    previewTitle: canPreview ? 'Preview selected file' : 'Select exactly one file to preview',
    downloadTitle: !canDownload ? 'Select one or more files to download' : selectedFileUrls.length === 1 ? 'Download selected file' : `Download ${selectedFileUrls.length} selected files as a ZIP`,
    renameTitle: canRename ? 'Rename selected item' : 'Select exactly one file or folder to rename',
    deleteTitle: total > 1 ? `Move ${total} selected items to the SharePoint recycle bin` : canDelete ? 'Delete selected item' : 'Select a file or folder before deleting'
  };
}
