import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './AiKnowledgeWorkspace.module.scss';
import type { IAiKnowledgeWorkspaceProps } from './IAiKnowledgeWorkspaceProps';
import { escape } from '@microsoft/sp-lodash-subset';

interface ICitation {
  title: string;
  url: string;
  snippet?: string;
}

interface IAskResponse {
  answer: string;
  citations?: ICitation[];
  provider?: string;
  requestId?: string;
  status?: string;
  suggestedActions?: string[];
  error?: string;
}

interface ILegalDocument {
  id: number;
  name: string;
  type: string;
  libraryTitle: string;
  url: string;
  serverRelativeUrl: string;
  folderPath: string;
  displayFolderPath: string;
  snippet: string;
  lastModified: string;
  modifiedBy: string;
}

interface ISharePointFileItem {
  Id: number;
  FileLeafRef: string;
  FileRef: string;
  FileDirRef: string;
  File_x0020_Type?: string;
  Modified: string;
  FSObjType: number;
  Editor?: {
    Title?: string;
  };
}

interface ISharePointListResponse {
  value: ISharePointFileItem[];
}

type LibraryStatus = 'loading' | 'loaded' | 'empty' | 'error';
type ViewMode = 'parents' | 'children' | 'files';
type SurfaceMode = 'document-library' | 'sharepoint-list' | 'site-pages';
type NavigationState = Pick<IAiKnowledgeWorkspaceState, 'selectedFolderPath' | 'selectedParentFolderName' | 'selectedFileUrl' | 'viewMode'>;

interface IAiKnowledgeWorkspaceState {
  question: string;
  answer: string;
  citations: ICitation[];
  suggestedActions: string[];
  provider: string;
  requestId: string;
  isLoading: boolean;
  error: string;
  commandMessage: string;
  isSettingsOpen: boolean;
  isPreviewOpen: boolean;
  surfaceMode: SurfaceMode;
  documents: ILegalDocument[];
  folderPaths: string[];
  selectedFolderPath: string;
  selectedParentFolderName: string;
  selectedFileUrl: string;
  viewMode: ViewMode;
  isAiPanelOpen: boolean;
  libraryStatus: LibraryStatus;
  libraryMessage: string;
  backendStatus: 'checking' | 'online' | 'offline';
  backendMessage: string;
}

const legalActions: string[] = [
  'Summarize selected litigation document',
  'Extract key dates and deadlines',
  'Identify parties and claims',
  'Find missing evidence or open questions',
  'Draft a case timeline'
];

export default class AiKnowledgeWorkspace extends React.Component<IAiKnowledgeWorkspaceProps, IAiKnowledgeWorkspaceState> {
  private readonly _fileInputRef: React.RefObject<HTMLInputElement> = React.createRef<HTMLInputElement>();

  public constructor(props: IAiKnowledgeWorkspaceProps) {
    super(props);

    this.state = {
      question: 'Summarize selected litigation document',
      answer: '',
      citations: [],
      suggestedActions: [],
      provider: 'mock',
      requestId: '',
      isLoading: false,
      error: '',
      commandMessage: '',
      isSettingsOpen: false,
      isPreviewOpen: false,
      surfaceMode: 'document-library',
      documents: [],
      folderPaths: [],
      selectedFolderPath: '',
      selectedParentFolderName: '',
      selectedFileUrl: '',
      viewMode: 'parents',
      isAiPanelOpen: false,
      libraryStatus: 'loading',
      libraryMessage: `Loading ${props.documentLibraryName || 'Litigation Documents'}...`,
      backendStatus: 'checking',
      backendMessage: 'Checking backend health...'
    };
  }

  public componentDidMount(): void {
    this._checkBackendHealth().catch(() => undefined);
    this._loadLibraryFiles().catch(() => undefined);
  }

  public componentDidUpdate(prevProps: IAiKnowledgeWorkspaceProps): void {
    if (prevProps.documentLibraryName !== this.props.documentLibraryName || prevProps.siteUrl !== this.props.siteUrl) {
      this._loadLibraryFiles().catch(() => undefined);
    }
  }

  public render(): React.ReactElement<IAiKnowledgeWorkspaceProps> {
    const selectedFile = this._getSelectedFile();

    return (
      <section className={styles.aiKnowledgeWorkspace}>
        <div className={styles.commandBar}>
          <button className={styles.primaryCommand} type="button" onClick={() => this._openUploadPicker()}>↑ Upload</button>
          <button className={styles.commandButton} type="button" onClick={() => this._createFolder().catch(() => undefined)}>＋ New</button>
          <button className={styles.commandButton} type="button" onClick={() => this.setState({ isSettingsOpen: true })}>⚙ Settings</button>
          <button
            className={styles.commandButton}
            type="button"
            disabled={!selectedFile}
            title={selectedFile ? `Preview ${selectedFile.name}` : 'Select a file to preview'}
            onClick={() => this.setState({ isPreviewOpen: true })}
          >
            👁 Preview
          </button>
          <button
            className={styles.dangerCommand}
            type="button"
            disabled={!selectedFile}
            title={selectedFile ? `Delete ${selectedFile.name}` : 'Select a file before deleting'}
            onClick={() => this._deleteSelectedFile().catch(() => undefined)}
          >
            🗑 Delete
          </button>
          <button className={styles.askCommand} type="button" onClick={() => this.setState({ isAiPanelOpen: true })}>
            ✨ Ask AI
          </button>
          <button className={styles.refreshCommand} type="button" onClick={() => this._loadLibraryFiles().catch(() => undefined)}>
            Refresh
          </button>
        </div>

        <input
          ref={this._fileInputRef}
          className={styles.hiddenInput}
          type="file"
          multiple
          onChange={(event) => this._uploadFiles(event).catch(() => undefined)}
        />

        <div className={styles.surfaceTabs} aria-label="SharePoint AI surface filters">
          <button
            className={this.state.surfaceMode === 'document-library' ? styles.surfaceTabActive : styles.surfaceTab}
            type="button"
            onClick={() => this._selectSurfaceMode('document-library')}
          >
            Document Library
          </button>
          <button
            className={this.state.surfaceMode === 'sharepoint-list' ? styles.surfaceTabActive : styles.surfaceTab}
            type="button"
            onClick={() => this._selectSurfaceMode('sharepoint-list')}
          >
            SharePoint List
          </button>
          <button
            className={this.state.surfaceMode === 'site-pages' ? styles.surfaceTabActive : styles.surfaceTab}
            type="button"
            onClick={() => this._selectSurfaceMode('site-pages')}
          >
            Site Pages
          </button>
        </div>

        <div className={styles.pageLayout}>
          <main className={styles.libraryCard}>
            <div className={styles.libraryHeader}>
              <div>
                <div className={styles.backendPill}>
                  <span className={this._getStatusDotClass()} />
                  <strong>{escape(this.state.backendMessage)}</strong>
                </div>
                <h1>{escape(this.props.documentLibraryName)}</h1>
              </div>
            </div>
            {this.state.commandMessage && <div className={styles.commandMessage}>{escape(this.state.commandMessage)}</div>}

            {this.state.surfaceMode === 'document-library' ? this._renderLibraryContent(selectedFile) : this._renderFutureSurface()} 
          </main>

          {this.state.isAiPanelOpen && this._renderAiPopup(selectedFile)}
          {this.state.isSettingsOpen && this._renderSettingsPopup()}
          {this.state.isPreviewOpen && selectedFile && this._renderPreviewPopup(selectedFile)}
        </div>
      </section>
    );
  }


  private _renderAiPopup(selectedFile: ILegalDocument | undefined): React.ReactElement {
    return (
      <div className={styles.aiOverlay} role="dialog" aria-modal="true" aria-label="Ask Legal AI">
        <div className={styles.aiDialog}>
          <button className={styles.closeButton} type="button" aria-label="Close Ask AI panel" onClick={() => this.setState({ isAiPanelOpen: false })}>×</button>
  <aside className={styles.aiPanel}>
        <div className={styles.aiPanelHeader}>
          <div>
            <p>ASK LEGAL AI</p>
            <h2>Document Assistant</h2>
          </div>
          <span className={this._getProviderBadgeClass()}>{escape(this.state.provider || 'mock')}</span>
        </div>

        <div className={styles.selectedBox}>
          <strong>Selected context</strong>
          <span>{escape(this._getSelectedContextLabel())}</span>
        </div>

        <div className={styles.actionChips}>
          {legalActions.slice(0, 3).map((action) => (
            <button key={action} type="button" onClick={() => this._runSuggestedPrompt(action)}>
              {action}
            </button>
          ))}
        </div>

        <label className={styles.questionLabel} htmlFor="ai-question">Question</label>
        <textarea
          id="ai-question"
          className={styles.questionInput}
          value={this.state.question}
          onChange={(event) => this.setState({ question: event.currentTarget.value })}
          placeholder="Ask about selected legal documents..."
        />

        <button
          className={styles.askButton}
          type="button"
          disabled={this.state.isLoading || !this.state.question.trim()}
          onClick={() => this._ask(this.state.question).catch(() => undefined)}
        >
          {this.state.isLoading ? 'Analyzing...' : 'Ask AI'}
        </button>

        {this.state.error && <div className={styles.errorBox}>{escape(this.state.error)}</div>}

        <div className={styles.answerBox}>
          <h3>Response</h3>
          {this.state.answer ? <p>{this.state.answer}</p> : <p className={styles.placeholder}>Select a file for file-level analysis, or ask from the current library/folder context.</p>}
          {this.state.requestId && <span className={styles.requestId}>Request ID: {escape(this.state.requestId)}</span>}
        </div>

        <div className={styles.citationBox}>
          <h3>Citations</h3>
          {this.state.citations.length > 0 ? this.state.citations.map((citation) => (
            <a key={`${citation.title}-${citation.url}`} href={citation.url} target="_blank" rel="noreferrer">
              <strong>{citation.title}</strong>
              <span>{citation.snippet || citation.url}</span>
            </a>
          )) : <p className={styles.placeholder}>Source-linked citations will appear here.</p>}
        </div>
  </aside>
        </div>
      </div>
    );
  }



  private _renderPreviewPopup(selectedFile: ILegalDocument): React.ReactElement {
    return (
      <div className={styles.previewOverlay} role="dialog" aria-modal="true" aria-label="Document preview">
        <div className={styles.previewDialog}>
          <div className={styles.previewHeader}>
            <div>
              <span>DOCUMENT PREVIEW</span>
              <h2>{escape(selectedFile.name)}</h2>
            </div>
            <div className={styles.previewActions}>
              <a href={selectedFile.url} target="_blank" rel="noreferrer">Open in SharePoint</a>
              <button className={styles.closeButton} type="button" aria-label="Close document preview" onClick={() => this.setState({ isPreviewOpen: false })}>×</button>
            </div>
          </div>
          <iframe className={styles.previewFrame} src={this._getPreviewUrl(selectedFile)} title={`Preview ${selectedFile.name}`} />
        </div>
      </div>
    );
  }


  private _renderSettingsPopup(): React.ReactElement {
    return (
      <div className={styles.aiOverlay} role="dialog" aria-modal="true" aria-label="Workspace settings">
        <div className={styles.settingsDialog}>
          <button className={styles.closeButton} type="button" aria-label="Close settings" onClick={() => this.setState({ isSettingsOpen: false })}>×</button>
          <div className={styles.settingsHeader}>
            <p>WORKSPACE SETTINGS</p>
            <h2>Library connection</h2>
          </div>
          <div className={styles.settingsGrid}>
            <strong>SharePoint site</strong>
            <span>{escape(this.props.siteUrl)}</span>
            <strong>Document library</strong>
            <span>{escape(this.props.documentLibraryName)}</span>
            <strong>Backend API</strong>
            <span>{escape(this.props.functionEndpoint)}</span>
            <strong>Current upload target</strong>
            <span>{escape(this._getCurrentFolderServerRelativeUrl())}</span>
            <strong>Loaded metadata</strong>
            <span>{this.state.documents.length} file(s), {this.state.folderPaths.length} folder path(s)</span>
          </div>
          <button className={styles.askButton} type="button" onClick={() => this._checkBackendHealth().catch(() => undefined)}>
            Recheck backend health
          </button>
        </div>
      </div>
    );
  }

  private _openUploadPicker(): void {
    this.setState({ commandMessage: '' });
    this._fileInputRef.current?.click();
  }

  private async _uploadFiles(event: React.ChangeEvent<HTMLInputElement>): Promise<void> {
    const files = Array.from(event.currentTarget.files || []);
    event.currentTarget.value = '';

    if (files.length === 0) {
      return;
    }

    const targetFolder = this._getCurrentFolderServerRelativeUrl();
    this.setState({ commandMessage: `Uploading ${files.length} file(s) to ${targetFolder}...` });

    try {
      for (const file of files) {
        const endpoint = `${this.props.siteUrl}/_api/web/GetFolderByServerRelativeUrl('${this._escapeODataString(targetFolder)}')/Files/add(url='${this._escapeODataString(file.name)}',overwrite=true)`;
        const body = await file.arrayBuffer();
        const response = await this.props.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
          headers: {
            Accept: 'application/json;odata=nometadata'
          },
          body
        });

        if (!response.ok) {
          throw new Error(`Upload failed for ${file.name} (${response.status} ${response.statusText}).`);
        }
      }

      this.setState({ commandMessage: `Uploaded ${files.length} file(s). Refreshing library metadata...` });
      await this._loadLibraryFiles();
      this.setState({ commandMessage: `Uploaded ${files.length} file(s) to ${targetFolder}.` });
    } catch (error) {
      this.setState({ commandMessage: (error as Error).message || 'Upload failed.' });
    }
  }


  private async _deleteSelectedFile(): Promise<void> {
    const selectedFile = this._getSelectedFile();

    if (!selectedFile) {
      this.setState({ commandMessage: 'Select a file before deleting.' });
      return;
    }

    const confirmed = window.confirm(`Move this file to the SharePoint recycle bin?\n\n${selectedFile.name}`);

    if (!confirmed) {
      return;
    }

    const previousParentFolderName = this.state.selectedParentFolderName;
    const previousFolderPath = this.state.selectedFolderPath;
    const previousViewMode = this.state.viewMode;

    this.setState({ commandMessage: `Deleting ${selectedFile.name}...` });

    try {
      const endpoint = `${this.props.siteUrl}/_api/web/GetFileByServerRelativeUrl('${this._escapeODataString(selectedFile.serverRelativeUrl)}')/recycle()`;
      const response = await this.props.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      });

      if (!response.ok) {
        throw new Error(`Delete failed for ${selectedFile.name} (${response.status} ${response.statusText}).`);
      }

      await this._loadLibraryFiles({
        selectedParentFolderName: previousParentFolderName,
        selectedFolderPath: previousFolderPath,
        selectedFileUrl: '',
        viewMode: previousViewMode
      });
      this.setState({
        isPreviewOpen: false,
        answer: '',
        citations: [],
        suggestedActions: [],
        requestId: '',
        commandMessage: `Moved to recycle bin: ${selectedFile.name}`
      });
    } catch (error) {
      this.setState({ commandMessage: (error as Error).message || 'Delete failed.' });
    }
  }

  private async _createFolder(): Promise<void> {
    const parentFolder = this._getCurrentFolderServerRelativeUrl();
    const folderName = window.prompt(`New folder name under:\n${parentFolder}`);

    if (!folderName || !folderName.trim()) {
      return;
    }

    const cleanFolderName = folderName.trim().replace(/[\\/:*?"<>|]/g, '-');
    const newFolderPath = `${parentFolder}/${cleanFolderName}`;
    this.setState({ commandMessage: `Creating folder ${cleanFolderName}...` });

    try {
      const endpoint = `${this.props.siteUrl}/_api/web/folders/add('${this._escapeODataString(newFolderPath)}')`;
      const response = await this.props.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
        headers: {
          Accept: 'application/json;odata=nometadata'
        }
      });

      if (!response.ok) {
        throw new Error(`Folder creation failed (${response.status} ${response.statusText}).`);
      }

      await this._loadLibraryFiles();
      this.setState({ commandMessage: `Created folder: ${cleanFolderName}` });
    } catch (error) {
      this.setState({ commandMessage: (error as Error).message || 'Folder creation failed.' });
    }
  }

  private _renderFutureSurface(): React.ReactElement {
    const title = this.state.surfaceMode === 'sharepoint-list' ? 'SharePoint List AI view' : 'Site Pages AI view';
    const message = this.state.surfaceMode === 'sharepoint-list'
      ? 'This filter is restored for the reusable roadmap. Next step: read selected list items from a legal matter tracker, deadline tracker, or evidence log.'
      : 'This filter is restored for the reusable roadmap. Next step: read current Site Page metadata and page canvas content for legal briefing support.';

    return (
      <div className={styles.noticeBox}>
        <strong>{title}</strong>
        <span>{message}</span>
      </div>
    );
  }

  private _selectSurfaceMode(surfaceMode: SurfaceMode): void {
    this.setState({
      surfaceMode,
      question: surfaceMode === 'document-library' ? 'Summarize selected litigation document' : surfaceMode === 'sharepoint-list' ? 'Summarize high priority legal list items' : 'Summarize this legal page',
      answer: '',
      citations: [],
      suggestedActions: [],
      requestId: '',
      error: ''
    });
  }

  private _renderLibraryContent(selectedFile: ILegalDocument | undefined): React.ReactElement {
    if (this.state.libraryStatus === 'loading') {
      return <div className={styles.noticeBox}>Loading files from {escape(this.props.documentLibraryName)}...</div>;
    }

    if (this.state.libraryStatus === 'error') {
      return <div className={styles.errorBox}>{escape(this.state.libraryMessage)}</div>;
    }

    if (this.state.documents.length === 0) {
      return <div className={styles.noticeBox}>No files found. Upload documents to {escape(this.props.documentLibraryName)} and click Refresh.</div>;
    }

    if (this.state.viewMode === 'files') {
      return this._renderFileRows(selectedFile);
    }

    if (this.state.viewMode === 'children') {
      return this._renderChildFolderRows();
    }

    return this._renderParentFolderRows();
  }

  private _renderParentFolderRows(): React.ReactElement {
    const rootDocuments = this._getRootDocuments();
    const selectedFile = this._getSelectedFile();

    return (
      <div className={styles.tableWrap}>
        <div className={styles.tableHeader}>
          <span />
          <span>Name / Category</span>
          <span>Files</span>
          <span>Latest modified</span>
        </div>
        {this._getParentFolderNames().map((parentName) => {
          const files = this._getDocumentsForParentFolder(parentName);
          const latest = files[0]?.lastModified || '';

          return (
            <button className={styles.tableRow} key={parentName} type="button" onClick={() => this._selectParentFolder(parentName)}>
              <span className={styles.selectCircle} />
              <span className={styles.nameCell}><span className={styles.folderGlyph}>📁</span>{escape(parentName)}</span>
              <span>{files.length}</span>
              <span>{escape(latest)}</span>
            </button>
          );
        })}
        {rootDocuments.map((doc) => this._renderFileRow(doc, selectedFile, false))}
      </div>
    );
  }


  private _renderChildFolderRows(): React.ReactElement {
    const childFolders = this._getChildFolderPaths(this.state.selectedParentFolderName);
    const directFiles = this._getDirectDocumentsForParent(this.state.selectedParentFolderName);
    const selectedFile = this._getSelectedFile();

    return (
      <div className={styles.tableWrap}>
        {this._renderChildBreadcrumb(childFolders.length, directFiles.length)}
        <div className={styles.tableHeader}>
          <span />
          <span>Name</span>
          <span>Files</span>
          <span>Latest modified</span>
        </div>
        {childFolders.map((folderPath) => {
          const files = this._getDocumentsForFolderTree(folderPath);
          const latest = files[0]?.lastModified || '';

          return (
            <button className={styles.tableRow} key={folderPath} type="button" onClick={() => this._selectFolder(folderPath)}>
              <span className={styles.selectCircle} />
              <span className={styles.nameCell}><span className={styles.folderGlyph}>📁</span>{escape(this._getLastFolderSegment(folderPath))}</span>
              <span>{files.length}</span>
              <span>{escape(latest)}</span>
            </button>
          );
        })}
        {directFiles.map((doc) => this._renderFileRow(doc, selectedFile, false))}
        {childFolders.length === 0 && directFiles.length === 0 && <div className={styles.noticeBox}>No child folders or files found under {escape(this.state.selectedParentFolderName)}.</div>}
      </div>
    );
  }

  private _renderChildBreadcrumb(childFolderCount: number, directFileCount: number): React.ReactElement {
    return (
      <nav className={styles.breadcrumbRow} aria-label="Folder breadcrumb">
        <button type="button" onClick={() => this.setState({ viewMode: 'parents', selectedParentFolderName: '', selectedFolderPath: '', selectedFileUrl: '' })}>
          {escape(this.props.documentLibraryName)}
        </button>
        <span className={styles.breadcrumbSeparator}>›</span>
        <span className={styles.breadcrumbCurrent}>{escape(this.state.selectedParentFolderName || 'Parent folder')}</span>
        <span className={styles.breadcrumbMeta}>{childFolderCount} folder(s), {directFileCount} file(s)</span>
      </nav>
    );
  }

  private _renderFileBreadcrumb(fileCount: number): React.ReactElement {
    const segments = this._getFolderSegments(this.state.selectedFolderPath);
    const parentFolderName = segments[0] || this.state.selectedParentFolderName;
    const childFolderName = segments.length > 1 ? segments.slice(1).join(' / ') : parentFolderName;

    return (
      <nav className={styles.breadcrumbRow} aria-label="Folder breadcrumb">
        <button type="button" onClick={() => this.setState({ viewMode: 'parents', selectedParentFolderName: '', selectedFolderPath: '', selectedFileUrl: '' })}>
          {escape(this.props.documentLibraryName)}
        </button>
        <span className={styles.breadcrumbSeparator}>›</span>
        <button type="button" onClick={() => this.setState({ viewMode: 'children', selectedParentFolderName: parentFolderName, selectedFolderPath: '', selectedFileUrl: '' })}>
          {escape(parentFolderName)}
        </button>
        {childFolderName !== parentFolderName && <span className={styles.breadcrumbSeparator}>›</span>}
        {childFolderName !== parentFolderName && <span className={styles.breadcrumbCurrent}>{escape(childFolderName)}</span>}
        <span className={styles.breadcrumbMeta}>{fileCount} file(s)</span>
      </nav>
    );
  }

  private _renderFileRow(doc: ILegalDocument, selectedFile: ILegalDocument | undefined, showModifiedBy: boolean): React.ReactElement {
    const isSelected = selectedFile?.url === doc.url;

    return (
      <div key={doc.url} className={isSelected ? styles.tableRowActive : styles.tableRow}>
        <button
          className={isSelected ? styles.selectCircleActive : styles.selectCircleButton}
          type="button"
          aria-label={`Select ${doc.name}`}
          onClick={() => this.setState({ selectedFileUrl: doc.url, selectedFolderPath: doc.folderPath })}
        />
        <button
          className={styles.nameCellButton}
          type="button"
          title={`Preview ${doc.name}`}
          onClick={() => this.setState({ selectedFileUrl: doc.url, selectedFolderPath: doc.folderPath, isPreviewOpen: true })}
        >
          <span className={styles.fileGlyph}>{this._getFileGlyph(doc.type)}</span>
          <span>{escape(doc.name)}</span>
        </button>
        <span>{escape(showModifiedBy ? doc.lastModified : doc.type)}</span>
        <span>{escape(showModifiedBy ? doc.modifiedBy : doc.lastModified)}</span>
      </div>
    );
  }

  private _renderFileRows(selectedFile: ILegalDocument | undefined): React.ReactElement {
    const visibleDocuments = this._getVisibleDocuments();

    return (
      <div className={styles.tableWrap}>
        {this._renderFileBreadcrumb(visibleDocuments.length)}
        <div className={styles.tableHeader}>
          <span />
          <span>Name / Classification</span>
          <span>Recently modified</span>
          <span>Modified by</span>
        </div>
        {visibleDocuments.length === 0 && <div className={styles.noticeBox}>This folder is now empty. You are still in the same folder location.</div>}
        {visibleDocuments.map((doc) => this._renderFileRow(doc, selectedFile, true))}
      </div>
    );
  }

  private async _loadLibraryFiles(navigationState?: Partial<NavigationState>): Promise<void> {
    const libraryName = this.props.documentLibraryName || 'Litigation Documents';
    this.setState({ libraryStatus: 'loading', libraryMessage: `Loading ${libraryName}...` });

    try {
      const escapedLibraryName = libraryName.replace(/'/g, "''");
      const endpoint = `${this.props.siteUrl}/_api/web/lists/getByTitle('${escapedLibraryName}')/items?$select=Id,FileLeafRef,FileRef,FileDirRef,File_x0020_Type,Modified,FSObjType,Editor/Title&$expand=Editor&$filter=FSObjType eq 0&$orderby=FileDirRef asc,Modified desc&$top=200`;
      const response: SPHttpClientResponse = await this.props.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);

      if (!response.ok) {
        throw new Error(`Could not read library '${libraryName}' (${response.status} ${response.statusText}).`);
      }

      const payload = await response.json() as ISharePointListResponse;
      const documents = (payload.value || []).map(item => this._mapFileItem(item, libraryName));
      const folderPaths = this._getUniqueFolderPaths(documents);

      this.setState({
        documents,
        folderPaths,
        selectedFolderPath: navigationState?.selectedFolderPath || '',
        selectedParentFolderName: navigationState?.selectedParentFolderName || '',
        selectedFileUrl: navigationState?.selectedFileUrl || '',
        viewMode: navigationState?.viewMode || 'parents',
        libraryStatus: documents.length > 0 ? 'loaded' : 'empty',
        libraryMessage: documents.length > 0 ? `Loaded ${documents.length} file(s) across ${folderPaths.length} folder(s).` : `${libraryName} exists, but no files were found.`
      });
    } catch (error) {
      this.setState({
        documents: [],
        folderPaths: [],
        selectedFolderPath: '',
        selectedParentFolderName: '',
        selectedFileUrl: '',
        viewMode: 'parents',
        libraryStatus: 'error',
        libraryMessage: (error as Error).message || `Could not read ${libraryName}.`
      });
    }
  }

  private _mapFileItem(item: ISharePointFileItem, libraryName: string): ILegalDocument {
    const absoluteUrl = item.FileRef.startsWith('http') ? item.FileRef : `${window.location.origin}${item.FileRef}`;
    const fileType = item.File_x0020_Type || item.FileLeafRef.split('.').pop() || 'file';
    const displayFolderPath = this._formatFolderName(item.FileDirRef);

    return {
      id: item.Id,
      name: item.FileLeafRef,
      type: fileType.toUpperCase(),
      libraryTitle: libraryName,
      url: absoluteUrl,
      serverRelativeUrl: item.FileRef,
      folderPath: item.FileDirRef,
      displayFolderPath,
      snippet: `SharePoint file from ${displayFolderPath}. Last modified ${this._formatDate(item.Modified)}.`,
      lastModified: this._formatDate(item.Modified),
      modifiedBy: item.Editor?.Title || this.props.userDisplayName || 'Unknown'
    };
  }


  private _selectParentFolder(parentFolderName: string): void {
    this.setState({
      selectedParentFolderName: parentFolderName,
      selectedFolderPath: '',
      selectedFileUrl: '',
      viewMode: 'children',
      answer: '',
      citations: [],
      suggestedActions: [],
      requestId: '',
      error: ''
    });
  }

  private _selectFolder(folderPath: string): void {
    const firstFileInFolder = this._getDocumentsForFolder(folderPath)[0];

    this.setState({
      selectedFolderPath: folderPath,
      selectedFileUrl: firstFileInFolder?.url || '',
      viewMode: 'files',
      answer: '',
      citations: [],
      suggestedActions: [],
      requestId: '',
      error: ''
    });
  }

  private _runSuggestedPrompt(action: string): void {
    this.setState({ question: action });
    this._ask(action).catch(() => undefined);
  }

  private _getUniqueFolderPaths(documents: ILegalDocument[]): string[] {
    return documents
      .map(doc => doc.folderPath)
      .filter((folderPath, index, allPaths) => folderPath && allPaths.indexOf(folderPath) === index)
      .sort((a, b) => this._formatFolderName(a).localeCompare(this._formatFolderName(b)));
  }



  private _getParentFolderNames(): string[] {
    return this.state.folderPaths
      .map(folderPath => this._getFolderSegments(folderPath)[0])
      .filter((folderName): folderName is string => !!folderName && folderName !== 'Library root')
      .filter((folderName, index, allNames) => allNames.indexOf(folderName) === index)
      .sort((a, b) => a.localeCompare(b));
  }

  private _getChildFolderPaths(parentFolderName: string): string[] {
    const childFolderMap: { [childKey: string]: string } = {};

    this.state.folderPaths.forEach((folderPath) => {
      const segments = this._getFolderSegments(folderPath);

      if (segments[0] !== parentFolderName || segments.length < 2) {
        return;
      }

      const childName = segments[1];
      const childPath = this._getServerRelativePathForChild(parentFolderName, childName);
      childFolderMap[childPath] = childPath;
    });

    return Object.keys(childFolderMap).sort((a, b) => this._getLastFolderSegment(a).localeCompare(this._getLastFolderSegment(b)));
  }

  private _getDocumentsForParentFolder(parentFolderName: string): ILegalDocument[] {
    return this.state.documents.filter(doc => this._getFolderSegments(doc.folderPath)[0] === parentFolderName);
  }

  private _getRootDocuments(): ILegalDocument[] {
    return this.state.documents.filter(doc => this._getFolderSegments(doc.folderPath)[0] === 'Library root');
  }

  private _getDirectDocumentsForParent(parentFolderName: string): ILegalDocument[] {
    return this.state.documents.filter((doc) => {
      const segments = this._getFolderSegments(doc.folderPath);
      return segments[0] === parentFolderName && segments.length === 1;
    });
  }

  private _getDocumentsForFolderTree(folderPath: string): ILegalDocument[] {
    const folderName = this._formatFolderName(folderPath);
    return this.state.documents.filter((doc) => {
      const docFolderName = this._formatFolderName(doc.folderPath);
      return docFolderName === folderName || docFolderName.indexOf(`${folderName}/`) === 0;
    });
  }


  private _getServerRelativePathForChild(parentFolderName: string, childFolderName: string): string {
    return `${this._getLibraryRootServerRelativeUrl()}/${parentFolderName}/${childFolderName}`;
  }

  private _getFolderSegments(folderPath: string): string[] {
    const formatted = this._formatFolderName(folderPath);
    if (!formatted || formatted === 'Library root') {
      return ['Library root'];
    }

    return formatted.split('/').map(segment => segment.trim()).filter(Boolean);
  }

  private _getLastFolderSegment(folderPath: string): string {
    const segments = this._getFolderSegments(folderPath);
    return segments[segments.length - 1] || 'Library root';
  }

  private _getVisibleDocuments(): ILegalDocument[] {
    if (!this.state.selectedFolderPath) {
      return [];
    }

    return this._getDocumentsForFolderTree(this.state.selectedFolderPath);
  }

  private _getDocumentsForFolder(folderPath: string): ILegalDocument[] {
    return this.state.documents.filter(doc => doc.folderPath === folderPath);
  }

  private _getSelectedFile(): ILegalDocument | undefined {
    return this.state.documents.find(doc => doc.url === this.state.selectedFileUrl);
  }

  private _getSelectedContextLabel(): string {
    const selectedFile = this._getSelectedFile();
    if (this.state.surfaceMode === 'sharepoint-list') {
      return 'SharePoint List / future legal matter tracker or evidence log';
    }

    if (this.state.surfaceMode === 'site-pages') {
      return 'Site Pages / future legal briefing page';
    }

    if (selectedFile) {
      return `${selectedFile.displayFolderPath} / ${selectedFile.name}`;
    }

    if (this.state.selectedFolderPath) {
      return `${this.props.documentLibraryName} / ${this._formatFolderName(this.state.selectedFolderPath)} folder context`;
    }

    if (this.state.selectedParentFolderName) {
      return `${this.props.documentLibraryName} / ${this.state.selectedParentFolderName} parent folder context`;
    }

    return `${this.props.documentLibraryName} / library context`;
  }

  private _getActiveFolderPathForAsk(): string {
    if (this.state.selectedFolderPath) {
      return this.state.selectedFolderPath;
    }

    if (this.state.selectedParentFolderName) {
      return this._getServerRelativePathForParent(this.state.selectedParentFolderName);
    }

    return this._getLibraryRootServerRelativeUrl();
  }

  private _getActiveContextDocuments(): ILegalDocument[] {
    if (this.state.selectedFolderPath) {
      return this._getVisibleDocuments();
    }

    if (this.state.selectedParentFolderName) {
      return this._getDocumentsForParentFolder(this.state.selectedParentFolderName);
    }

    return this.state.documents.slice(0, 20);
  }

  private _formatFolderName(folderPath: string): string {
    if (!folderPath) {
      return 'Library root';
    }

    const marker = `/${this.props.documentLibraryName}`;
    const markerIndex = folderPath.toLowerCase().indexOf(marker.toLowerCase());

    if (markerIndex >= 0) {
      const relativePath = folderPath.substring(markerIndex + marker.length).replace(/^\//, '');
      return relativePath || 'Library root';
    }

    return folderPath.split('/').filter(Boolean).pop() || 'Library root';
  }



  private _getPreviewUrl(selectedFile: ILegalDocument): string {
    const separator = selectedFile.url.indexOf('?') >= 0 ? '&' : '?';
    return `${selectedFile.url}${separator}web=1`;
  }

  private _getCurrentFolderServerRelativeUrl(): string {
    if (this.state.selectedFolderPath) {
      return this.state.selectedFolderPath;
    }

    if (this.state.selectedParentFolderName) {
      return this._getServerRelativePathForParent(this.state.selectedParentFolderName);
    }

    return this._getLibraryRootServerRelativeUrl();
  }

  private _getLibraryRootServerRelativeUrl(): string {
    const firstFolderPath = this.state.folderPaths[0] || this.state.documents[0]?.folderPath || '';
    const marker = `/${this.props.documentLibraryName}`;
    const markerIndex = firstFolderPath.toLowerCase().indexOf(marker.toLowerCase());

    if (markerIndex >= 0) {
      return firstFolderPath.substring(0, markerIndex + marker.length);
    }

    try {
      const sitePath = new URL(this.props.siteUrl).pathname.replace(/\/$/, '');
      return `${sitePath}/${this.props.documentLibraryName}`;
    } catch {
      return `/${this.props.documentLibraryName}`;
    }
  }

  private _getServerRelativePathForParent(parentFolderName: string): string {
    const matchingPath = this.state.folderPaths.find(folderPath => this._getFolderSegments(folderPath)[0] === parentFolderName);

    if (!matchingPath) {
      return `${this._getLibraryRootServerRelativeUrl()}/${parentFolderName}`;
    }

    const rootPath = this._getLibraryRootServerRelativeUrl();
    return `${rootPath}/${parentFolderName}`;
  }

  private _escapeODataString(value: string): string {
    return value.replace(/'/g, "''");
  }

  private _formatDate(value: string): string {
    if (!value) {
      return 'Unknown modified date';
    }

    return new Date(value).toLocaleString();
  }

  private _getFileGlyph(fileType: string): string {
    const type = fileType.toLowerCase();

    if (type === 'pdf') {
      return '📄';
    }

    if (type === 'docx' || type === 'doc') {
      return '📝';
    }

    if (type === 'xlsx' || type === 'xls') {
      return '📊';
    }

    return '📎';
  }

  private async _checkBackendHealth(): Promise<void> {
    const healthEndpoint = this._getHealthEndpoint();

    try {
      const response = await fetch(healthEndpoint, { method: 'GET' });
      const result = await response.json() as { ok?: boolean; service?: string; provider?: string };

      if (!response.ok || !result.ok) {
        throw new Error(`Health check returned ${response.status}`);
      }

      this.setState({
        backendStatus: 'online',
        backendMessage: `${result.service || 'backend'} online`,
        provider: result.provider || this.state.provider
      });
    } catch (error) {
      this.setState({
        backendStatus: 'offline',
        backendMessage: (error as Error).message || 'Backend health check failed',
        provider: 'offline'
      });
    }
  }

  private _getHealthEndpoint(): string {
    return this.props.functionEndpoint
      .replace('/api/chat', '/api/health')
      .replace('/api/ask', '/api/health');
  }

  private _getStatusDotClass(): string {
    if (this.state.backendStatus === 'online') {
      return styles.statusDotOnline;
    }

    if (this.state.backendStatus === 'offline') {
      return styles.statusDotOffline;
    }

    return styles.statusDotChecking;
  }

  private _getProviderBadgeClass(): string {
    return this.state.provider === 'offline' ? styles.providerBadgeOffline : styles.providerBadge;
  }

  private async _ask(question: string): Promise<void> {
    const trimmedQuestion = question.trim();

    if (!trimmedQuestion) {
      return;
    }

    const selectedFile = this._getSelectedFile();

    const contextFiles = selectedFile ? [selectedFile] : this._getActiveContextDocuments().slice(0, 20);
    const contextSnippets = contextFiles.slice(0, 5).map(doc => doc.snippet);
    const activeFolderPath = selectedFile?.folderPath || this._getActiveFolderPathForAsk();

    this.setState({ isLoading: true, error: '', answer: '', citations: [], suggestedActions: [], requestId: '' });

    try {
      const response = await fetch(this.props.functionEndpoint, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json'
        },
        body: JSON.stringify({
          question: trimmedQuestion,
          scenario: this.state.surfaceMode === 'document-library' ? 'legal-document-library' : this.state.surfaceMode === 'sharepoint-list' ? 'legal-sharepoint-list' : 'legal-site-page',
          mode: this.state.surfaceMode === 'document-library' ? 'legal-analysis' : 'context-preview',
          contextType: this.state.surfaceMode === 'document-library' ? 'document-library' : this.state.surfaceMode === 'sharepoint-list' ? 'sharepoint-list' : 'site-page',
          siteUrl: this.props.siteUrl,
          pageUrl: this.props.pageUrl,
          pageTitle: this.props.siteTitle,
          libraryName: this.props.documentLibraryName,
          folderPath: activeFolderPath,
          listTitle: this.props.documentLibraryName,
          pageContext: {
            webTitle: this.props.siteTitle,
            userEmail: this.props.userDisplayName
          },
          selectedFiles: contextFiles,
          selectedItems: [],
          documentSnippets: contextSnippets,
          knowledgeScope: 'legal-document-library'
        })
      });

      const result = await response.json() as IAskResponse;

      if (!response.ok || result.status === 'error') {
        throw new Error(result.error || `Backend returned ${response.status}`);
      }

      this.setState({
        answer: result.answer,
        citations: result.citations || [],
        suggestedActions: result.suggestedActions || [],
        provider: result.provider || 'mock',
        requestId: result.requestId || '',
        isLoading: false
      });
    } catch (error) {
      this.setState({
        error: (error as Error).message || 'Unable to reach the AI backend.',
        provider: 'offline',
        isLoading: false
      });
    }
  }
}










