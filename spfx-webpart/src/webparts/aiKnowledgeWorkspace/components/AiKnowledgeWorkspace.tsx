import * as React from 'react';
import { AadHttpClient, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './AiKnowledgeWorkspace.module.scss';
import type { IAiKnowledgeWorkspaceProps } from './IAiKnowledgeWorkspaceProps';
import { escape } from '@microsoft/sp-lodash-subset';

interface ICitation {
  title: string;
  url: string;
  snippet?: string;
}

interface IConversationTurn {
  id: string;
  question: string;
  answer?: string;
  displayedAnswer?: string;
  citations?: ICitation[];
  requestId?: string;
  error?: string;
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
  conversationTurns: IConversationTurn[];
  suggestedActions: string[];
  provider: string;
  requestId: string;
  isLoading: boolean;
  error: string;
  commandMessage: string;
  isSettingsOpen: boolean;
  isNewMenuOpen: boolean;
  isCreateFolderOpen: boolean;
  newFolderName: string;
  isCreatingFolder: boolean;
  confirmationDialog: 'delete' | 'ingest' | '';
  isPreviewOpen: boolean;
  surfaceMode: SurfaceMode;
  documents: ILegalDocument[];
  folderPaths: string[];
  selectedFolderPath: string;
  selectedParentFolderName: string;
  selectedFileUrl: string;
  selectedFolderForAction: string;
  viewMode: ViewMode;
  isAiPanelOpen: boolean;
  libraryStatus: LibraryStatus;
  libraryMessage: string;
  backendStatus: 'checking' | 'online' | 'offline';
  backendMessage: string;
  ragStatus: 'checking' | 'ready' | 'not-indexed' | 'error';
  ragMessage: string;
  isIngesting: boolean;
  ingestionMessage: string;
}

const MAX_SELECTED_PDF_BYTES = 4 * 1024 * 1024;

export default class AiKnowledgeWorkspace extends React.Component<IAiKnowledgeWorkspaceProps, IAiKnowledgeWorkspaceState> {
  private readonly _fileInputRef: React.RefObject<HTMLInputElement> = React.createRef<HTMLInputElement>();
  private readonly _conversationAreaRef: React.RefObject<HTMLDivElement> = React.createRef<HTMLDivElement>();
  private _answerRevealTimer: number | undefined;

  public constructor(props: IAiKnowledgeWorkspaceProps) {
    super(props);

    this.state = {
      question: '',
      answer: '',
      citations: [],
      conversationTurns: [],
      suggestedActions: [],
      provider: 'mock',
      requestId: '',
      isLoading: false,
      error: '',
      commandMessage: '',
      isSettingsOpen: false,
      isNewMenuOpen: false,
      isCreateFolderOpen: false,
      newFolderName: '',
      isCreatingFolder: false,
      confirmationDialog: '',
      isPreviewOpen: false,
      surfaceMode: 'document-library',
      documents: [],
      folderPaths: [],
      selectedFolderPath: '',
      selectedParentFolderName: '',
      selectedFileUrl: '',
      selectedFolderForAction: '',
      viewMode: 'parents',
      isAiPanelOpen: false,
      libraryStatus: 'loading',
      libraryMessage: `Loading ${props.documentLibraryName || 'Litigation Documents'}...`,
      backendStatus: 'checking',
      backendMessage: 'Checking backend health...',
      ragStatus: 'checking',
      ragMessage: 'Checking Azure AI Search RAG...',
      isIngesting: false,
      ingestionMessage: 'No library ingestion has run from this workspace yet.'
    };
  }

  public componentDidMount(): void {
    this._checkBackendHealth().catch(() => undefined);
    this._checkRagHealth().catch(() => undefined);
    this._loadLibraryFiles().catch(() => undefined);
  }

  public componentWillUnmount(): void {
    if (this._answerRevealTimer !== undefined) {
      window.clearInterval(this._answerRevealTimer);
    }
  }

  public componentDidUpdate(prevProps: IAiKnowledgeWorkspaceProps): void {
    if (prevProps.documentLibraryName !== this.props.documentLibraryName || prevProps.siteUrl !== this.props.siteUrl) {
      this._loadLibraryFiles().catch(() => undefined);
    }
  }

  public render(): React.ReactElement<IAiKnowledgeWorkspaceProps> {
    const selectedFile = this._getSelectedFile();
    const selectedFolderForAction = this.state.selectedFolderForAction;
    const deleteTargetName = selectedFile?.name || (selectedFolderForAction ? this._getLastFolderSegment(selectedFolderForAction) : '');
    const originalLibraryUrl = this._getOriginalLibraryUrl();

    return (
      <section className={styles.aiKnowledgeWorkspace}>
        <div className={styles.commandBar}>
          <button className={styles.primaryCommand} type="button" onClick={() => this._openUploadPicker()}>↑ Upload</button>
          <div className={styles.newCommandWrap}>
            <button
              className={styles.commandButton}
              type="button"
              aria-haspopup="menu"
              aria-expanded={this.state.isNewMenuOpen}
              onClick={() => this.setState((currentState) => ({ isNewMenuOpen: !currentState.isNewMenuOpen }))}
            >
              ＋ New <span className={styles.commandChevron} aria-hidden="true" />
            </button>
            {this.state.isNewMenuOpen && (
              <div className={styles.newMenu} role="menu" aria-label="New item menu">
                <button
                  type="button"
                  role="menuitem"
                  onClick={() => this.setState({ isNewMenuOpen: false, isCreateFolderOpen: true, newFolderName: '' })}
                >
                  <span className={styles.newMenuIcon}>📁</span>
                  <span><strong>Folder</strong><small>Create a folder in the current location</small></span>
                </button>
              </div>
            )}
          </div>
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
            disabled={!selectedFile && !selectedFolderForAction}
            title={deleteTargetName ? `Delete ${deleteTargetName}` : 'Select a file or folder before deleting'}
            onClick={() => this._deleteSelectedItem()}
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
              <a className={styles.openLibraryLink} href={originalLibraryUrl} target="_blank" rel="noreferrer">
                ↗ Open in SharePoint
              </a>
            </div>
            {this.state.commandMessage && <div className={styles.commandMessage}>{escape(this.state.commandMessage)}</div>}

            {this.state.surfaceMode === 'document-library' ? this._renderLibraryContent(selectedFile) : this._renderFutureSurface()} 
          </main>

          {this.state.isAiPanelOpen && this._renderAiPopup(selectedFile)}
          {this.state.isCreateFolderOpen && this._renderCreateFolderDialog()}
          {this.state.confirmationDialog && this._renderConfirmationDialog()}
          {this.state.isSettingsOpen && this._renderSettingsPopup()}
          {this.state.isPreviewOpen && selectedFile && this._renderPreviewPopup(selectedFile)}
        </div>
      </section>
    );
  }


  private _renderAiPopup(selectedFile: ILegalDocument | undefined): React.ReactElement {
    const isFileAnalysis = Boolean(selectedFile);
    const scope = isFileAnalysis ? selectedFile!.name : `Entire ${this.props.documentLibraryName} library`;
    const sampleQuestions = isFileAnalysis
      ? ['Summarize this document', 'List key dates and deadlines', 'What actions are required?']
      : ['Find documents about this topic', 'Summarize key next actions', 'Identify important dates and deadlines'];
    const introduction = isFileAnalysis
      ? `Ask questions about this file in ${this.props.documentLibraryName}. Get a summary, find key dates, or identify important details.`
      : `Ask questions across ${this.props.documentLibraryName}. Answers use relevant document evidence and include sources.`;
    const placeholder = isFileAnalysis
      ? 'Ask a question about this file…'
      : 'Describe the document, topic, person, date, or evidence you need…';

    return (
      <div className={styles.aiOverlay} role="dialog" aria-modal="true" aria-label="Ask AI">
        <div className={styles.aiDialog}>
          <aside className={styles.aiPanel}>
            <div className={styles.aiPanelHeader}>
              <div>
                <h2>Ask AI</h2>
                <span>{escape(this.props.documentLibraryName)}</span>
              </div>
              <button className={styles.closeButton} type="button" aria-label="Close Ask AI panel" onClick={() => this.setState({ isAiPanelOpen: false })}>×</button>
            </div>

            <div className={styles.aiScope}>
              <span className={styles.aiScopeDot} />
              <div>
                <strong>{isFileAnalysis ? 'Selected file' : 'Library search'}</strong>
                <span>{escape(scope)}</span>
              </div>
            </div>

            <div className={styles.aiIntro}>{introduction}</div>

            <div ref={this._conversationAreaRef} className={styles.conversationArea} aria-live="polite">
              {this.state.conversationTurns.map((turn) => (
                <div key={turn.id} className={styles.conversationTurn}>
                  <div className={styles.questionBubble}>{escape(turn.question)}</div>
                  {(turn.displayedAnswer || turn.answer) && (
                    <div className={styles.answerBox}>
                      <p>{turn.displayedAnswer || turn.answer}</p>
                      {turn.requestId && <span className={styles.requestId}>Request ID: {escape(turn.requestId)}</span>}
                    </div>
                  )}
                  {turn.error && <div className={styles.errorBox}>{escape(turn.error)}</div>}
                  {(turn.citations || []).length > 0 && (
                    <details className={styles.citationBox} open>
                      <summary>Sources ({turn.citations!.length})</summary>
                      {turn.citations!.map((citation) => (
                        <a key={`${turn.id}-${citation.title}-${citation.url}`} href={citation.url} target="_blank" rel="noreferrer">
                          <strong>{citation.title}</strong>
                          <span>{citation.snippet || citation.url}</span>
                        </a>
                      ))}
                    </details>
                  )}
                </div>
              ))}
              {this.state.isLoading && <div className={styles.loadingMessage}>Searching the available document evidence…</div>}
            </div>

            <div className={styles.aiComposer}>
              <div className={styles.sampleQuestions} aria-label="Sample questions">
                {sampleQuestions.map((sampleQuestion) => (
                  <button
                    key={sampleQuestion}
                    type="button"
                    onClick={() => {
                      this.setState({ question: sampleQuestion });
                      this._ask(sampleQuestion).catch(() => undefined);
                    }}
                  >
                    {sampleQuestion}
                  </button>
                ))}
              </div>
              <textarea
                id="ai-question"
                className={styles.questionInput}
                value={this.state.question}
                onChange={(event) => this.setState({ question: event.currentTarget.value })}
                onKeyDown={(event) => {
                  if (event.key === 'Enter' && !event.shiftKey && this.state.question.trim() && !this.state.isLoading) {
                    event.preventDefault();
                    this._ask(this.state.question).catch(() => undefined);
                  }
                }}
                placeholder={placeholder}
                aria-label="Message AI"
              />
              <div className={styles.composerFooter}>
                <span>Press Enter to send · Shift + Enter for a new line</span>
              </div>
              <p className={styles.aiDisclaimer}>AI-generated content may be incorrect.</p>
            </div>
          </aside>
        </div>
      </div>
    );
  }



  private _renderCreateFolderDialog(): React.ReactElement {
    const currentLocation = this._getCurrentFolderServerRelativeUrl();

    return (
      <div className={styles.folderDialogOverlay} role="dialog" aria-modal="true" aria-label="Create new folder">
        <form className={styles.folderDialog} onSubmit={(event) => { event.preventDefault(); this._createFolder(this.state.newFolderName).catch(() => undefined); }}>
          <div className={styles.folderDialogHeader}>
            <div>
              <span>NEW</span>
              <h2>Create a folder</h2>
            </div>
            <button type="button" className={styles.closeButton} aria-label="Close create folder" onClick={() => this.setState({ isCreateFolderOpen: false, newFolderName: '' })}>×</button>
          </div>
          <label className={styles.folderInputLabel} htmlFor="folder-name">Folder name</label>
          <input
            id="folder-name"
            className={styles.folderNameInput}
            type="text"
            autoFocus
            maxLength={128}
            value={this.state.newFolderName}
            onChange={(event) => this.setState({ newFolderName: event.currentTarget.value })}
            placeholder="Enter a folder name"
          />
          <div className={styles.folderLocation}>
            <span>Location</span>
            <strong>{escape(currentLocation)}</strong>
          </div>
          <div className={styles.folderDialogActions}>
            <button type="button" className={styles.folderCancelButton} onClick={() => this.setState({ isCreateFolderOpen: false, newFolderName: '' })}>Cancel</button>
            <button type="submit" className={styles.folderCreateButton} disabled={this.state.isCreatingFolder || !this.state.newFolderName.trim()}>
              {this.state.isCreatingFolder ? 'Creating…' : 'Create'}
            </button>
          </div>
        </form>
      </div>
    );
  }

  private _renderConfirmationDialog(): React.ReactElement {
    const isDelete = this.state.confirmationDialog === 'delete';
    const selectedFile = this._getSelectedFile();
    const selectedFolderPath = this.state.selectedFolderForAction;
    const isFolder = !selectedFile && !!selectedFolderPath;
    const targetName = selectedFile?.name || (selectedFolderPath ? this._getLastFolderSegment(selectedFolderPath) : '');
    const title = isDelete ? `Delete ${isFolder ? 'folder' : 'file'}?` : 'Index library text PDFs?';
    const message = isDelete
      ? `“${targetName}” will be moved to the SharePoint recycle bin.`
      : `Index supported text-layer PDFs from ${this.props.documentLibraryName}. Scanned/image PDFs will be reported as OCR required.`;
    const confirmLabel = isDelete ? 'Move to recycle bin' : 'Start indexing';

    return (
      <div className={styles.folderDialogOverlay} role="dialog" aria-modal="true" aria-label={title}>
        <div className={styles.confirmationDialog}>
          <div className={styles.folderDialogHeader}>
            <div>
              <span>{isDelete ? 'CONFIRM DELETE' : 'CONFIRM INDEXING'}</span>
              <h2>{title}</h2>
            </div>
            <button type="button" className={styles.closeButton} aria-label="Close confirmation" onClick={() => this.setState({ confirmationDialog: '' })}>×</button>
          </div>
          <p>{message}</p>
          <div className={styles.folderDialogActions}>
            <button type="button" className={styles.folderCancelButton} onClick={() => this.setState({ confirmationDialog: '' })}>Cancel</button>
            <button
              type="button"
              className={isDelete ? styles.confirmDeleteButton : styles.folderCreateButton}
              onClick={() => {
                this.setState({ confirmationDialog: '' });
                if (isDelete) {
                  this._confirmDeleteSelectedItem().catch(() => undefined);
                } else {
                  this._confirmLibraryIngest().catch(() => undefined);
                }
              }}
            >
              {confirmLabel}
            </button>
          </div>
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
            <strong>Azure AI Search RAG</strong>
            <span>{escape(this.state.ragMessage)}</span>
            <strong>Library ingestion</strong>
            <span>{escape(this.state.ingestionMessage)}</span>
            <strong>Current upload target</strong>
            <span>{escape(this._getCurrentFolderServerRelativeUrl())}</span>
            <strong>Loaded metadata</strong>
            <span>{this.state.documents.length} file(s), {this.state.folderPaths.length} folder path(s)</span>
          </div>
          <button
            className={styles.askButton}
            type="button"
            disabled={this.state.isIngesting || this.state.ragStatus === 'error'}
            onClick={() => this._ingestLibrary()}
          >
            {this.state.isIngesting ? 'Indexing text PDFs...' : 'Index text PDFs now'}
          </button>
          <button className={styles.askButton} type="button" onClick={() => { this._checkBackendHealth().catch(() => undefined); this._checkRagHealth().catch(() => undefined); }}>
            Recheck backend and RAG
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
    const navigationState: NavigationState = {
      selectedParentFolderName: this.state.selectedParentFolderName,
      selectedFolderPath: this.state.selectedFolderPath,
      selectedFileUrl: this.state.selectedFileUrl,
      viewMode: this.state.viewMode
    };
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

      this.setState({ commandMessage: `Uploaded ${files.length} file(s). Refreshing the current folder...` });
      await this._loadLibraryFiles(navigationState);
      this.setState({ commandMessage: `Uploaded ${files.length} file(s) to ${targetFolder}.` });
    } catch (error) {
      this.setState({ commandMessage: (error as Error).message || 'Upload failed.' });
    }
  }


  private _deleteSelectedItem(): void {
    const selectedFile = this._getSelectedFile();
    const selectedFolderPath = this.state.selectedFolderForAction;

    if (!selectedFile && !selectedFolderPath) {
      this.setState({ commandMessage: 'Select a file or folder before deleting.' });
      return;
    }

    this.setState({ confirmationDialog: 'delete' });
  }

  private async _confirmDeleteSelectedItem(): Promise<void> {
    const selectedFile = this._getSelectedFile();
    const selectedFolderPath = this.state.selectedFolderForAction;
    const isFolder = !selectedFile && !!selectedFolderPath;
    const targetName = selectedFile?.name || (selectedFolderPath ? this._getLastFolderSegment(selectedFolderPath) : '');

    if (!targetName) {
      this.setState({ commandMessage: 'Select a file or folder before deleting.' });
      return;
    }

    const navigationState: NavigationState = {
      selectedParentFolderName: this.state.selectedParentFolderName,
      selectedFolderPath: this.state.selectedFolderPath,
      selectedFileUrl: '',
      viewMode: this.state.viewMode
    };
    this.setState({ commandMessage: `Deleting ${targetName}...` });

    try {
      const targetUrl = isFolder ? selectedFolderPath : selectedFile!.serverRelativeUrl;
      const api = isFolder ? 'GetFolderByServerRelativeUrl' : 'GetFileByServerRelativeUrl';
      const endpoint = `${this.props.siteUrl}/_api/web/${api}('${this._escapeODataString(targetUrl)}')/recycle()`;
      const response = await this.props.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
        headers: { Accept: 'application/json;odata=nometadata' }
      });

      if (!response.ok) {
        throw new Error(`Delete failed for ${targetName} (${response.status} ${response.statusText}).`);
      }

      await this._loadLibraryFiles(navigationState);
      this.setState({
        isPreviewOpen: false,
        selectedFolderForAction: '',
        answer: '',
        citations: [],
        suggestedActions: [],
        requestId: '',
        commandMessage: `Moved to recycle bin: ${targetName}`
      });
    } catch (error) {
      this.setState({ commandMessage: (error as Error).message || 'Delete failed.' });
    }
  }

  private async _createFolder(folderName: string): Promise<void> {
    const parentFolder = this._getCurrentFolderServerRelativeUrl();
    const cleanFolderName = folderName.trim().replace(/[\\/:*?"<>|]/g, '-');

    if (!cleanFolderName) {
      return;
    }

    const newFolderPath = `${parentFolder}/${cleanFolderName}`;
    const navigationState: NavigationState = {
      selectedFolderPath: this.state.selectedFolderPath,
      selectedParentFolderName: this.state.selectedParentFolderName,
      selectedFileUrl: this.state.selectedFileUrl,
      viewMode: this.state.viewMode
    };
    this.setState({ isCreatingFolder: true, commandMessage: `Creating folder ${cleanFolderName}...` });

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

      await this._loadLibraryFiles(navigationState);
      this.setState({
        isCreatingFolder: false,
        isCreateFolderOpen: false,
        newFolderName: '',
        commandMessage: `Created folder: ${cleanFolderName}`
      });
    } catch (error) {
      this.setState({
        isCreatingFolder: false,
        commandMessage: (error as Error).message || 'Folder creation failed.'
      });
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

    if (this.state.documents.length === 0 && this.state.folderPaths.length === 0) {
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
          const folderPath = this._getServerRelativePathForParent(parentName);
          const isFolderSelected = this.state.selectedFolderForAction === folderPath;

          return (
            <div className={isFolderSelected ? styles.tableRowActive : styles.tableRow} key={parentName}>
              <button
                className={isFolderSelected ? styles.selectCircleActive : styles.selectCircleButton}
                type="button"
                aria-label={`Select folder ${parentName}`}
                onClick={() => this.setState({ selectedFolderForAction: isFolderSelected ? '' : folderPath, selectedFileUrl: '' })}
              />
              <button className={styles.nameCellButton} type="button" onClick={() => this._selectParentFolder(parentName)}>
                <span className={styles.folderGlyph}>📁</span>{escape(parentName)}
              </button>
              <span>{files.length}</span>
              <span>{escape(latest)}</span>
            </div>
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
          const isFolderSelected = this.state.selectedFolderForAction === folderPath;

          return (
            <div className={isFolderSelected ? styles.tableRowActive : styles.tableRow} key={folderPath}>
              <button
                className={isFolderSelected ? styles.selectCircleActive : styles.selectCircleButton}
                type="button"
                aria-label={`Select folder ${this._getLastFolderSegment(folderPath)}`}
                onClick={() => this.setState({ selectedFolderForAction: isFolderSelected ? '' : folderPath, selectedFileUrl: '' })}
              />
              <button className={styles.nameCellButton} type="button" onClick={() => this._selectFolder(folderPath)}>
                <span className={styles.folderGlyph}>📁</span>{escape(this._getLastFolderSegment(folderPath))}
              </button>
              <span>{files.length}</span>
              <span>{escape(latest)}</span>
            </div>
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
          onClick={() => this.setState({ selectedFileUrl: doc.url, selectedFolderPath: doc.folderPath, selectedFolderForAction: '', conversationTurns: [] })}
        />
        <button
          className={styles.nameCellButton}
          type="button"
          title={`Preview ${doc.name}`}
          onClick={() => this.setState({ selectedFileUrl: doc.url, selectedFolderPath: doc.folderPath, selectedFolderForAction: '', isPreviewOpen: true, conversationTurns: [] })}
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
      const endpoint = `${this.props.siteUrl}/_api/web/lists/getByTitle('${escapedLibraryName}')/items?$select=Id,FileLeafRef,FileRef,FileDirRef,File_x0020_Type,Modified,FSObjType,Editor/Title&$expand=Editor&$orderby=FileDirRef asc,Modified desc&$top=200`;
      const response: SPHttpClientResponse = await this.props.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);

      if (!response.ok) {
        throw new Error(`Could not read library '${libraryName}' (${response.status} ${response.statusText}).`);
      }

      const payload = await response.json() as ISharePointListResponse;
      const items = payload.value || [];
      const documents = items
        .filter((item) => item.FSObjType === 0)
        .map(item => this._mapFileItem(item, libraryName));
      const explicitFolderPaths = items
        .filter((item) => item.FSObjType === 1 && !!item.FileRef)
        .map((item) => item.FileRef);
      const folderPaths = this._getUniqueFolderPaths(documents, explicitFolderPaths);

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
      selectedFolderForAction: '',
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
      selectedFolderForAction: '',
      viewMode: 'files',
      answer: '',
      citations: [],
      suggestedActions: [],
      requestId: '',
      error: ''
    });
  }

  private _getUniqueFolderPaths(documents: ILegalDocument[], explicitFolderPaths: string[] = []): string[] {
    return [...documents.map(doc => doc.folderPath), ...explicitFolderPaths]
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

  private _getOriginalLibraryUrl(): string {
    const serverRelativeUrl = this._getLibraryRootServerRelativeUrl();
    try {
      return `${new URL(this.props.siteUrl).origin}${serverRelativeUrl}`;
    } catch {
      return `${this.props.siteUrl.replace(/\/$/, '')}/${this.props.documentLibraryName}`;
    }
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
      const apiClient = await this.props.aadHttpClientFactory.getClient(this.props.functionApiResource);
      const response = await apiClient.fetch(healthEndpoint, AadHttpClient.configurations.v1, { method: 'GET' });
      if (!response.ok) {
        throw new Error(`Backend health check returned ${response.status} ${response.statusText || 'without a response body'}.`);
      }

      const result = await response.json() as { ok?: boolean; service?: string; provider?: string };

      if (!result.ok) {
        throw new Error('Backend health check returned an invalid success response.');
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

  private _getRagHealthEndpoint(): string {
    return this.props.functionEndpoint
      .replace('/api/chat', '/api/rag/health')
      .replace('/api/ask', '/api/rag/health');
  }

  private _getRagIngestEndpoint(): string {
    return this.props.functionEndpoint
      .replace('/api/chat', '/api/rag/ingest')
      .replace('/api/ask', '/api/rag/ingest');
  }

  private _getRagAnswerEndpoint(): string {
    return this.props.functionEndpoint
      .replace('/api/chat', '/api/rag/answer')
      .replace('/api/ask', '/api/rag/answer');
  }

  private _ingestLibrary(): void {
    this.setState({ confirmationDialog: 'ingest' });
  }

  private async _confirmLibraryIngest(): Promise<void> {
    this.setState({ isIngesting: true, ingestionMessage: `Indexing text PDFs from ${this.props.documentLibraryName}...` });
    try {
      const apiClient = await this.props.aadHttpClientFactory.getClient(this.props.functionApiResource);
      const response = await apiClient.fetch(this._getRagIngestEndpoint(), AadHttpClient.configurations.v1, { method: 'POST' });
      const payload = await response.json() as {
        error?: string;
        detail?: string;
        ingestion?: {
          discoveredFiles: number;
          indexedDocuments: number;
          indexedChunks: number;
          skippedUnsupported: number;
          skippedTooLarge: number;
          requiresOcr: number;
          failedFiles: Array<{ name: string; reason: string }>;
        };
      };
      if (!response.ok || !payload.ingestion) {
        throw new Error(payload.detail || payload.error || `Library ingestion returned ${response.status}.`);
      }

      const ingestion = payload.ingestion;
      this.setState({
        isIngesting: false,
        ragStatus: ingestion.indexedChunks > 0 ? 'ready' : this.state.ragStatus,
        ingestionMessage: `Indexed ${ingestion.indexedChunks} chunk(s) from ${ingestion.indexedDocuments}/${ingestion.discoveredFiles} file(s). OCR required: ${ingestion.requiresOcr}; unsupported: ${ingestion.skippedUnsupported}; over size limit: ${ingestion.skippedTooLarge}; failed: ${ingestion.failedFiles.length}.`
      });
      await this._checkRagHealth();
    } catch (error) {
      this.setState({
        isIngesting: false,
        ingestionMessage: (error as Error).message || 'Library ingestion failed.'
      });
    }
  }

  private async _checkRagHealth(): Promise<void> {
    try {
      const apiClient = await this.props.aadHttpClientFactory.getClient(this.props.functionApiResource);
      const response = await apiClient.fetch(this._getRagHealthEndpoint(), AadHttpClient.configurations.v1, { method: 'GET' });

      if (!response.ok) {
        let detail = '';
        try {
          const failure = await response.json() as { error?: string; detail?: string };
          detail = failure.detail || failure.error || '';
        } catch {
          detail = '';
        }
        throw new Error(`RAG health check returned ${response.status} ${response.statusText || 'without a response body'}${detail ? `: ${detail}` : '.'}`);
      }

      const result = await response.json() as { ok?: boolean; rag?: { configured?: boolean; indexName?: string; indexExists?: boolean } };
      if (!result.ok || !result.rag?.configured) {
        throw new Error('Azure AI Search RAG is not configured.');
      }

      this.setState({
        ragStatus: result.rag.indexExists ? 'ready' : 'not-indexed',
        ragMessage: result.rag.indexExists
          ? `Azure AI Search connected: ${result.rag.indexName || 'library index'} is ready.`
          : `Azure AI Search connected: ${result.rag.indexName || 'library index'} will be created during ingestion.`
      });
    } catch (error) {
      this.setState({
        ragStatus: 'error',
        ragMessage: (error as Error).message || 'Azure AI Search RAG health check failed.'
      });
    }
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

  private async _getSelectedPdfPayload(selectedFile: ILegalDocument | undefined): Promise<{ name: string; fileType: string; contentBase64: string } | undefined> {
    if (!selectedFile) {
      return undefined;
    }

    if (selectedFile.type.toLowerCase() !== 'pdf') {
      throw new Error('This proof of concept currently supports selected PDF files only.');
    }

    const response: SPHttpClientResponse = await this.props.spHttpClient.get(selectedFile.url, SPHttpClient.configurations.v1, {
      headers: { Accept: 'application/pdf' }
    });

    if (!response.ok) {
      throw new Error(`Unable to read ${selectedFile.name} from SharePoint (${response.status} ${response.statusText}).`);
    }

    const fileBytes = new Uint8Array(await (await response.blob()).arrayBuffer());
    if (fileBytes.byteLength > MAX_SELECTED_PDF_BYTES) {
      throw new Error(`${selectedFile.name} exceeds the 4 MB proof-of-concept limit.`);
    }

    let binary = '';
    const chunkSize = 0x8000;
    for (let offset = 0; offset < fileBytes.length; offset += chunkSize) {
      const chunk = Array.from(fileBytes.subarray(offset, offset + chunkSize));
      binary += String.fromCharCode.apply(null, chunk);
    }

    return {
      name: selectedFile.name,
      fileType: 'pdf',
      contentBase64: window.btoa(binary)
    };
  }

  private _scrollConversationToBottom(): void {
    const conversationArea = this._conversationAreaRef.current;
    if (!conversationArea) {
      return;
    }

    window.requestAnimationFrame(() => {
      conversationArea.scrollTo({ top: conversationArea.scrollHeight, behavior: 'smooth' });
    });
  }

  private _revealAnswer(turnId: string, result: IAskResponse): void {
    if (this._answerRevealTimer !== undefined) {
      window.clearInterval(this._answerRevealTimer);
    }

    const answer = result.answer || '';
    const revealChunkSize = 14;
    let visibleLength = 0;

    const revealNextChunk = (): void => {
      visibleLength = Math.min(answer.length, visibleLength + revealChunkSize);
      const isComplete = visibleLength >= answer.length;
      const displayedAnswer = answer.slice(0, visibleLength);

      this.setState((currentState) => ({
        answer: isComplete ? answer : '',
        citations: isComplete ? result.citations || [] : [],
        suggestedActions: isComplete ? result.suggestedActions || [] : [],
        provider: result.provider || 'mock',
        requestId: isComplete ? result.requestId || '' : '',
        isLoading: !isComplete,
        conversationTurns: currentState.conversationTurns.map((turn) => turn.id === turnId
          ? {
            ...turn,
            displayedAnswer,
            answer: isComplete ? answer : undefined,
            citations: isComplete ? result.citations || [] : [],
            requestId: isComplete ? result.requestId || '' : ''
          }
          : turn)
      }), () => this._scrollConversationToBottom());

      if (isComplete && this._answerRevealTimer !== undefined) {
        window.clearInterval(this._answerRevealTimer);
        this._answerRevealTimer = undefined;
      }
    };

    revealNextChunk();
    if (visibleLength < answer.length) {
      this._answerRevealTimer = window.setInterval(revealNextChunk, 28);
    }
  }

  private async _ask(question: string): Promise<void> {
    const trimmedQuestion = question.trim();
    const turnId = `${Date.now()}-${Math.random().toString(36).slice(2)}`;

    if (!trimmedQuestion) {
      return;
    }

    const selectedFile = this._getSelectedFile();
    const useLibraryRag = !selectedFile && this.state.surfaceMode === 'document-library';
    const contextFiles = selectedFile ? [selectedFile] : [];
    const contextSnippets = contextFiles.map(doc => doc.snippet);
    const activeFolderPath = selectedFile?.folderPath || this._getActiveFolderPathForAsk();

    this.setState((currentState) => ({
      isLoading: true,
      error: '',
      question: '',
      answer: '',
      citations: [],
      suggestedActions: [],
      requestId: '',
      conversationTurns: [...currentState.conversationTurns, { id: turnId, question: trimmedQuestion }]
    }));

    try {
      const selectedDocument = selectedFile ? await this._getSelectedPdfPayload(selectedFile) : undefined;
      const apiClient = await this.props.aadHttpClientFactory.getClient(this.props.functionApiResource);
      const response = await apiClient.fetch(
        useLibraryRag ? this._getRagAnswerEndpoint() : this.props.functionEndpoint,
        AadHttpClient.configurations.v1, {
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
          selectedDocument,
          knowledgeScope: useLibraryRag ? 'library-wide-rag' : 'legal-document-library'
        })
      });

      if (!response.ok) {
        let detail = '';
        try {
          const failure = await response.json() as IAskResponse;
          detail = failure.error || '';
        } catch {
          detail = '';
        }
        throw new Error(detail || `Backend returned ${response.status} ${response.statusText || 'without a response body'}.`);
      }

      const result = await response.json() as IAskResponse;

      if (result.status === 'error') {
        throw new Error(result.error || 'Backend returned an error response.');
      }

      this._revealAnswer(turnId, result);
    } catch (error) {
      const message = (error as Error).message || 'Unable to reach the AI backend.';
      this.setState((currentState) => ({
        error: '',
        provider: 'offline',
        isLoading: false,
        conversationTurns: currentState.conversationTurns.map((turn) => turn.id === turnId
          ? { ...turn, error: message }
          : turn)
      }));
    }
  }
}










