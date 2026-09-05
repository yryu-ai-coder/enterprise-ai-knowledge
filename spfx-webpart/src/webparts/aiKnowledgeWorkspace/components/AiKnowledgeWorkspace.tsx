import * as React from 'react';
import * as ReactDOM from 'react-dom';
import JSZip from 'jszip';
import { AadHttpClient, SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import styles from './AiKnowledgeWorkspace.module.scss';
import type { IAiKnowledgeWorkspaceProps } from './IAiKnowledgeWorkspaceProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { formatOcrIngestionMessage, getOcrIngestEndpoint, IOcrIngestionResult } from './ocrIngestion';
import { buildListItemSelect } from './documentLibraryColumns';
import type { IDocumentLibraryDisplayColumn } from './documentLibraryColumns';
import { buildRenameEndpoint, buildRenameRequest, RenameItemType, validateRenameName } from './renameItem';
import { formatAnswerForDisplay } from './answerPresentation';
import { formatCitationSnippet } from './citationPresentation';
import { getFolderOpenSelection, getSelectionCommandState, toggleSelection } from './librarySelection';
import { getFileIconKind } from './fileIcon';
import { shouldDeferQuestionSubmitForComposition, shouldSubmitQuestionOnEnter } from './questionSubmission';
import { buildFolderSummaryQuestion } from './folderSummary';

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
  fieldValues: Record<string, unknown>;
}

interface ISharePointFileItem {
  [internalName: string]: unknown;
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

interface ISharePointLibraryMetadata {
  RootFolder?: {
    ServerRelativeUrl?: string;
  };
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
  isRenameOpen: boolean;
  renameName: string;
  isRenaming: boolean;
  isDownloading: boolean;
  confirmationDialog: 'delete' | 'ingest-text' | 'ingest-ocr' | '';
  isPreviewOpen: boolean;
  surfaceMode: SurfaceMode;
  documents: ILegalDocument[];
  folderPaths: string[];
  libraryRootServerRelativeUrl: string;
  selectedFolderPath: string;
  selectedParentFolderName: string;
  selectedFileUrl: string;
  selectedFolderForAction: string;
  selectedFileUrls: string[];
  selectedFolderPaths: string[];
  viewMode: ViewMode;
  isAiPanelOpen: boolean;
  aiPanelWidth: number;
  libraryStatus: LibraryStatus;
  libraryMessage: string;
  backendStatus: 'checking' | 'online' | 'offline';
  backendMessage: string;
  ragStatus: 'checking' | 'ready' | 'not-indexed' | 'error';
  ragMessage: string;
  isIngesting: boolean;
  ingestionMessage: string;
  isOcrIngesting: boolean;
  ocrIngestionMessage: string;
}

const MAX_SELECTED_PDF_BYTES = 4 * 1024 * 1024;

export default class AiKnowledgeWorkspace extends React.Component<IAiKnowledgeWorkspaceProps, IAiKnowledgeWorkspaceState> {
  private readonly _fileInputRef: React.RefObject<HTMLInputElement> = React.createRef<HTMLInputElement>();
  private readonly _conversationAreaRef: React.RefObject<HTMLDivElement> = React.createRef<HTMLDivElement>();
  private readonly _workspaceRef: React.RefObject<HTMLElement> = React.createRef<HTMLElement>();
  private _answerRevealTimer: number | undefined;
  private _commandMessageTimer: number | undefined;
  private _submitAfterComposition = false;
  private _isResizingAiPanel = false;
  private _aiResizeStartX = 0;
  private _aiResizeStartWidth = 560;
  private readonly _conversationHistoryByScope: Record<string, IConversationTurn[]> = {};

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
      isRenameOpen: false,
      renameName: '',
      isRenaming: false,
      isDownloading: false,
      confirmationDialog: '',
      isPreviewOpen: false,
      surfaceMode: 'document-library',
      documents: [],
      folderPaths: [],
      libraryRootServerRelativeUrl: '',
      selectedFolderPath: '',
      selectedParentFolderName: '',
      selectedFileUrl: '',
      selectedFolderForAction: '',
      selectedFileUrls: [],
      selectedFolderPaths: [],
      viewMode: 'parents',
      isAiPanelOpen: false,
      aiPanelWidth: 560,
      libraryStatus: 'loading',
      libraryMessage: props.documentLibraryName ? `Loading ${props.documentLibraryName}...` : 'Select a document library in the Web Part properties.',
      backendStatus: 'checking',
      backendMessage: 'Checking backend health...',
      ragStatus: 'checking',
      ragMessage: 'Checking Azure AI Search RAG...',
      isIngesting: false,
      ingestionMessage: 'No text-PDF library ingestion has run from this workspace yet.',
      isOcrIngesting: false,
      ocrIngestionMessage: 'No OCR-required PDF processing has run from this workspace yet.'
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
    if (this._commandMessageTimer !== undefined) {
      window.clearTimeout(this._commandMessageTimer);
    }
    window.removeEventListener('mousemove', this._resizeAiPanel);
    window.removeEventListener('mouseup', this._endAiPanelResize);
  }

  public componentDidUpdate(prevProps: IAiKnowledgeWorkspaceProps, prevState: IAiKnowledgeWorkspaceState): void {
    if (prevProps.documentLibraryName !== this.props.documentLibraryName || prevProps.siteUrl !== this.props.siteUrl || prevProps.displayColumns !== this.props.displayColumns) {
      this._loadLibraryFiles().catch(() => undefined);
    }

    if (prevState.conversationTurns !== this.state.conversationTurns && prevState.selectedFileUrl === this.state.selectedFileUrl) {
      this._saveConversationHistory(this.state.selectedFileUrl, this.state.conversationTurns);
    }

    if (prevState.commandMessage !== this.state.commandMessage) {
      if (this._commandMessageTimer !== undefined) {
        window.clearTimeout(this._commandMessageTimer);
      }
      if (this.state.commandMessage) {
        this._commandMessageTimer = window.setTimeout(() => this.setState({ commandMessage: '' }), 5000);
      }
    }
  }

  private _getSharePointFontFamily(): string | undefined {
    const workspace = this._workspaceRef.current;
    return workspace ? window.getComputedStyle(workspace).fontFamily : undefined;
  }

  private _getConversationHistoryKey(fileUrl: string): string {
    return `nextcore.ask-ai.history:${this.props.siteUrl}:${this.props.documentLibraryName}:${fileUrl || '__library__'}`;
  }

  private _saveConversationHistory(fileUrl: string, turns: IConversationTurn[]): void {
    const key = this._getConversationHistoryKey(fileUrl);
    this._conversationHistoryByScope[key] = turns;
    try {
      window.sessionStorage.setItem(key, JSON.stringify(turns));
    } catch {
      // Keep the in-memory history when browser storage is unavailable.
    }
  }

  private _getConversationHistory(fileUrl: string): IConversationTurn[] {
    const key = this._getConversationHistoryKey(fileUrl);
    if (this._conversationHistoryByScope[key]) {
      return this._conversationHistoryByScope[key];
    }
    try {
      const saved = window.sessionStorage.getItem(key);
      if (saved) {
        const turns = JSON.parse(saved) as IConversationTurn[];
        this._conversationHistoryByScope[key] = turns;
        return turns;
      }
    } catch {
      // Start a fresh history if the browser has cleared or blocked storage.
    }
    return [];
  }

  private _selectFile(doc: ILegalDocument, preview = false): void {
    this._saveConversationHistory(this.state.selectedFileUrl, this.state.conversationTurns);
    this.setState({
      selectedFileUrl: doc.url,
      selectedFolderPath: doc.folderPath,
      selectedFolderForAction: '',
      selectedFileUrls: [doc.url],
      selectedFolderPaths: [],
      isPreviewOpen: preview,
      conversationTurns: this._getConversationHistory(doc.url)
    });
  }

  private _toggleFileSelection(doc: ILegalDocument): void {
    this.setState((currentState) => {
      const selectedFileUrls = toggleSelection(currentState.selectedFileUrls, doc.url);
      const commandState = getSelectionCommandState(selectedFileUrls, currentState.selectedFolderPaths);
      return {
        selectedFileUrls,
        selectedFileUrl: commandState.selectedFileUrl,
        selectedFolderForAction: commandState.selectedFolderPath,
        isPreviewOpen: false
      };
    });
  }

  private _toggleFolderSelection(folderPath: string): void {
    this.setState((currentState) => {
      const selectedFolderPaths = toggleSelection(currentState.selectedFolderPaths, folderPath);
      const commandState = getSelectionCommandState(currentState.selectedFileUrls, selectedFolderPaths);
      return {
        selectedFolderPaths,
        selectedFileUrl: commandState.selectedFileUrl,
        selectedFolderForAction: commandState.selectedFolderPath,
        isPreviewOpen: false
      };
    });
  }

  private _selectFolderForAction(folderPath: string): void {
    this.setState({
      selectedFileUrl: '',
      selectedFolderForAction: folderPath,
      selectedFileUrls: [],
      selectedFolderPaths: [folderPath],
      isPreviewOpen: false
    });
  }

  private _openRenameForFile(doc: ILegalDocument): void {
    this.setState({
      selectedFileUrl: doc.url,
      selectedFolderForAction: '',
      selectedFileUrls: [doc.url],
      selectedFolderPaths: []
    }, () => this._openRenameDialog());
  }

  private _openRenameForFolder(folderPath: string): void {
    this.setState({
      selectedFileUrl: '',
      selectedFolderForAction: folderPath,
      selectedFileUrls: [],
      selectedFolderPaths: [folderPath]
    }, () => this._openRenameDialog());
  }

  public render(): React.ReactElement<IAiKnowledgeWorkspaceProps> {
    const selectedFile = this._getSelectedFile();
    const selectionCommandState = getSelectionCommandState(this.state.selectedFileUrls, this.state.selectedFolderPaths);
    const originalLibraryUrl = this._getOriginalLibraryUrl();

    return (
      <section
        ref={this._workspaceRef}
        className={styles.aiKnowledgeWorkspace}
        onClick={() => this.state.isNewMenuOpen && this.setState({ isNewMenuOpen: false })}
      >
        <div className={styles.surfaceTabs} aria-label="SharePoint AI surface filters">
          <button
            className={this.state.surfaceMode === 'document-library' ? styles.surfaceTabActive : styles.surfaceTab}
            type="button"
            onClick={() => this._selectSurfaceMode('document-library')}
          >
            Documents
          </button>
          <button
            className={this.state.surfaceMode === 'sharepoint-list' ? styles.surfaceTabActive : styles.surfaceTab}
            type="button"
            onClick={() => this._selectSurfaceMode('sharepoint-list')}
          >
            Lists
          </button>
          <button
            className={this.state.surfaceMode === 'site-pages' ? styles.surfaceTabActive : styles.surfaceTab}
            type="button"
            onClick={() => this._selectSurfaceMode('site-pages')}
          >
            Pages
          </button>
        </div>

        <div className={styles.commandBar}>
          <button className={styles.primaryCommand} type="button" onClick={() => this._openUploadPicker()}>↑ Upload</button>
          <div className={styles.newCommandWrap} onClick={(event) => event.stopPropagation()}>
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
            disabled={!selectionCommandState.canPreview}
            title={selectionCommandState.previewTitle}
            onClick={() => this.setState({ isPreviewOpen: true })}
          >
            👁 Preview
          </button>
          <button
            className={styles.commandButton}
            type="button"
            disabled={!selectionCommandState.canDownload || this.state.isDownloading}
            title={selectionCommandState.downloadTitle}
            onClick={() => this._downloadSelectedFiles().catch(() => undefined)}
          >
            {this.state.isDownloading ? '↓ Preparing…' : '↓ Download'}
          </button>
          <button
            className={styles.commandButton}
            type="button"
            disabled={!selectionCommandState.canRename}
            title={selectionCommandState.renameTitle}
            onClick={() => this._openRenameDialog()}
          >
            ✎ Rename
          </button>
          <button
            className={styles.dangerCommand}
            type="button"
            disabled={!selectionCommandState.canDelete}
            title={selectionCommandState.deleteTitle}
            onClick={() => this._deleteSelectedItem()}
          >
            🗑 Delete
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

        <div className={styles.pageLayout}>
          <main className={styles.libraryCard}>
            <div className={styles.libraryHeader}>
              <div>
                <div className={styles.backendPill}>
                  <span className={this._getStatusDotClass()} />
                  <strong>{escape(this.state.backendMessage)}</strong>
                </div>
                <h1>{escape(this.props.documentLibraryName || 'Select a document library')}</h1>
              </div>
              <div className={styles.headerActions}>
                <button className={styles.askCommand} type="button" onClick={() => this.setState({ isAiPanelOpen: true })}>
                  ✨ Ask AI
                </button>
                <a className={styles.openLibraryLink} href={originalLibraryUrl} target="_blank" rel="noreferrer">
                  ↗ Open in SharePoint
                </a>
              </div>
            </div>
            {this.state.commandMessage && <div className={styles.commandMessage}>{escape(this.state.commandMessage)}</div>}

            {this.state.surfaceMode === 'document-library' ? this._renderLibraryContent(selectedFile) : this._renderFutureSurface()} 
          </main>
          {this.state.surfaceMode === 'document-library' && this._renderDocumentSummary(selectedFile)}

          {this.state.isAiPanelOpen && this._renderAiPopup(selectedFile)}
          {this.state.isCreateFolderOpen && this._renderCreateFolderDialog()}
          {this.state.isRenameOpen && this._renderRenameDialog()}
          {this.state.confirmationDialog && this._renderConfirmationDialog()}
          {this.state.isSettingsOpen && this._renderSettingsPopup()}
          {this.state.isPreviewOpen && selectedFile && this._renderPreviewPopup(selectedFile)}
        </div>
      </section>
    );
  }


  private _renderDocumentSummary(selectedFile: ILegalDocument | undefined): React.ReactElement {
    const selectedDocuments = this.state.documents.filter((document) => this.state.selectedFileUrls.indexOf(document.url) >= 0);
    const documentForSummary = selectedFile || (selectedDocuments.length === 1 ? selectedDocuments[0] : undefined);

    const activeFolderPath = this.state.selectedFolderPath;
    if (!documentForSummary && activeFolderPath) {
      const folderDocuments = this._getDocumentsForFolderTree(activeFolderPath);
      const folderName = this._formatFolderName(activeFolderPath);
      return (
        <aside className={styles.documentSummary} aria-label="Folder details">
          <header className={styles.documentSummaryHeader}><h2>Folder details</h2><span>Folder and subfolders</span></header>
          <div className={styles.summaryFile}><span className={styles.fileGlyph}>▣</span><div><strong>{escape(folderName)}</strong><span>{folderDocuments.length} file(s) in this folder</span></div></div>
          <div className={styles.summaryBody}><h3>ASK AI</h3><p>Generate a summary or ask questions about the files in this folder.</p><div className={styles.summaryActions}><button type="button" onClick={() => this._openFolderSummaryAsk(activeFolderPath, true)}>Generate summary</button><button className={styles.summaryAskButton} type="button" onClick={() => this._openFolderSummaryAsk(activeFolderPath, false)}>Ask AI</button></div></div>
        </aside>
      );
    }

    if (!documentForSummary) {
      return (
        <aside className={styles.documentSummary} aria-label="Details">
          <header className={styles.documentSummaryHeader}><h2>Details</h2><span>Select a file or folder</span></header>
          <div className={styles.documentSummaryEmpty}>Select a file to view its details, generate a summary, or ask AI questions.</div>
        </aside>
      );
    }

    return (
      <aside className={styles.documentSummary} aria-label="File details">
        <header className={styles.documentSummaryHeader}><h2>File details</h2><span>Selected file</span></header>
        <div className={styles.summaryFile}><span className={styles.fileGlyph}>{this._renderFileGlyph(documentForSummary.type)}</span><div><strong>{escape(documentForSummary.name)}</strong><span>{escape(documentForSummary.type.toUpperCase())} file</span></div></div>
        <div className={styles.summaryMetadata}><div><span>MODIFIED</span><strong>{escape(documentForSummary.lastModified)}</strong></div><div><span>MODIFIED BY</span><strong>{escape(documentForSummary.modifiedBy || '—')}</strong></div></div>
        <div className={styles.summaryBody}><h3>ASK AI</h3><p>Generate a summary or ask questions about this file.</p><div className={styles.summaryActions}><button type="button" onClick={() => this._openSummaryAsk(documentForSummary, 'Summarize this document')}>Generate summary</button><button className={styles.summaryAskButton} type="button" onClick={() => this._openSummaryAsk(documentForSummary)}>Ask AI</button></div></div>
      </aside>
    );
  }

  private _getFolderConversationScopeKey(folderPath: string): string {
    return `folder:${folderPath}`;
  }

  private _openFolderSummaryAsk(folderPath: string, generateSummary: boolean): void {
    const question = buildFolderSummaryQuestion(this._formatFolderName(folderPath));
    const folderScopeKey = this._getFolderConversationScopeKey(folderPath);
    const folderSelection = getFolderOpenSelection();
    this._saveConversationHistory(this.state.selectedFileUrl, this.state.conversationTurns);
    this.setState({
      selectedFolderPath: folderPath,
      selectedFileUrl: folderScopeKey,
      selectedFileUrls: folderSelection.selectedFileUrls,
      selectedFolderPaths: folderSelection.selectedFolderPaths,
      selectedFolderForAction: '',
      isAiPanelOpen: true,
      question: generateSummary ? question : '',
      conversationTurns: this._getConversationHistory(folderScopeKey)
    }, () => {
      if (generateSummary) this._ask(question).catch(() => undefined);
    });
  }

  private _openSummaryAsk(document: ILegalDocument, question?: string): void {
    this.setState({
      selectedFileUrl: document.url,
      selectedFileUrls: [document.url],
      selectedFolderPaths: [],
      selectedFolderForAction: '',
      isAiPanelOpen: true,
      question: question || this.state.question,
      conversationTurns: this._getConversationHistory(document.url)
    }, () => {
      if (question) this._ask(question).catch(() => undefined);
    });
  }

  private _beginAiPanelResize = (event: React.MouseEvent<HTMLDivElement>): void => {
    event.preventDefault();
    this._isResizingAiPanel = true;
    this._aiResizeStartX = event.clientX;
    this._aiResizeStartWidth = this.state.aiPanelWidth;
    window.addEventListener('mousemove', this._resizeAiPanel);
    window.addEventListener('mouseup', this._endAiPanelResize);
  };

  private _resizeAiPanel = (event: MouseEvent): void => {
    if (!this._isResizingAiPanel) return;
    const maxWidth = Math.max(420, Math.min(1_000, window.innerWidth - 48));
    const width = Math.max(420, Math.min(maxWidth, this._aiResizeStartWidth + this._aiResizeStartX - event.clientX));
    this.setState({ aiPanelWidth: width });
  };

  private _endAiPanelResize = (): void => {
    this._isResizingAiPanel = false;
    window.removeEventListener('mousemove', this._resizeAiPanel);
    window.removeEventListener('mouseup', this._endAiPanelResize);
  };

  private _renderAiPopup(selectedFile: ILegalDocument | undefined): React.ReactElement {
    const selectedDocuments = this.state.documents.filter((document) => this.state.selectedFileUrls.indexOf(document.url) >= 0);
    const isFileAnalysis = selectedDocuments.length > 0;
    const activeFolderPath = this._getActiveFolderPathForAsk();
    const isFolderAnalysis = !isFileAnalysis && activeFolderPath.toLowerCase() !== this._getLibraryRootServerRelativeUrl().toLowerCase();
    const scope = selectedDocuments.length === 1
      ? selectedDocuments[0].name
      : selectedDocuments.length > 1
        ? `${selectedDocuments.length} selected documents`
        : isFolderAnalysis
          ? `Current folder: ${this._formatFolderName(activeFolderPath)}`
          : `Entire ${this.props.documentLibraryName} library`;
    const sampleQuestions = isFileAnalysis
      ? ['Summarize this document', 'List key dates and deadlines', 'What actions are required?']
      : isFolderAnalysis
        ? ['Summarize documents in this folder', 'What are the main topics in this folder?', 'Identify important dates and deadlines']
        : ['Find documents about this topic', 'Summarize key next actions', 'Identify important dates and deadlines'];
    const introduction = selectedDocuments.length === 1
      ? `Ask questions about this file in ${this.props.documentLibraryName}. Answers use evidence from this file and include sources.`
      : selectedDocuments.length > 1
        ? `Ask questions across these ${selectedDocuments.length} selected documents. Answers use only their indexed evidence and include sources.`
        : isFolderAnalysis
          ? `Ask questions about documents in the current folder. Answers use only indexed evidence from this folder and include sources.`
          : `Ask questions across ${this.props.documentLibraryName}. Answers use relevant document evidence and include sources.`;
    const placeholder = selectedDocuments.length > 0
      ? `Ask a question about ${selectedDocuments.length === 1 ? 'this file' : 'these selected documents'}…`
      : 'Describe the document, topic, person, date, or evidence you need…';
    const isWorkbench = /\/_layouts\/15\/workbench\.aspx$/i.test(window.location.pathname);
    const overlayClassName = isWorkbench ? styles.aiOverlay : `${styles.aiOverlay} ${styles.askAiOverlay}`;
    // The portal is attached to document.body, outside the SharePoint canvas
    // where the theme variables are inherited. Copy the resolved canvas font.
    const sharePointFontFamily = this._getSharePointFontFamily();

    // A SharePoint canvas web part lives in its own stacking context. Render the
    // production drawer at document.body so it can genuinely sit above the
    // Office suite bar and the edit footer.
    return ReactDOM.createPortal(
      <div className={overlayClassName} style={sharePointFontFamily ? { fontFamily: sharePointFontFamily } : undefined} role="dialog" aria-modal="true" aria-label="Ask AI">
        <div className={styles.aiDialog} style={{ width: `${this.state.aiPanelWidth}px` }}>
          <div className={styles.aiResizeHandle} role="separator" aria-label="Resize Ask AI panel" aria-orientation="vertical" onMouseDown={this._beginAiPanelResize} />
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
                <strong>{isFileAnalysis ? 'Selected file' : isFolderAnalysis ? 'Current folder' : 'Library search'}</strong>
                <span>{escape(scope)}</span>
              </div>
            </div>

            <div className={styles.aiIntro}>{introduction}</div>

            <div ref={this._conversationAreaRef} className={styles.conversationArea} aria-live="polite">
              {this.state.conversationTurns.map((turn) => (
                <div key={turn.id} className={styles.conversationTurn}>
                  <div className={styles.questionBubble}>{formatAnswerForDisplay(turn.question)}</div>
                  {(turn.displayedAnswer || turn.answer) && (
                    <div className={styles.answerBox}>
                      <p>{formatAnswerForDisplay(turn.displayedAnswer || turn.answer || '')}</p>
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
                          <span>{formatCitationSnippet(citation.snippet || citation.url)}</span>
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
                  const questionEvent = {
                    key: event.key,
                    shiftKey: event.shiftKey,
                    isComposing: event.nativeEvent.isComposing,
                    keyCode: event.keyCode,
                    question: this.state.question,
                    isLoading: this.state.isLoading
                  };
                  if (shouldDeferQuestionSubmitForComposition(questionEvent)) {
                    // The first Enter completes an IME syllable. Submit on composition end,
                    // when React has the complete controlled-textarea value.
                    this._submitAfterComposition = true;
                    return;
                  }
                  if (shouldSubmitQuestionOnEnter(questionEvent)) {
                    event.preventDefault();
                    this._ask(this.state.question).catch(() => undefined);
                  }
                }}
                onCompositionEnd={(event) => {
                  if (!this._submitAfterComposition) return;
                  this._submitAfterComposition = false;
                  const completedQuestion = event.currentTarget.value;
                  if (!completedQuestion.trim() || this.state.isLoading) return;
                  this.setState({ question: completedQuestion }, () => this._ask(completedQuestion).catch(() => undefined));
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
      </div>,
      document.body
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

  private _renderRenameDialog(): React.ReactElement {
    const selectedFile = this._getSelectedFile();
    const selectedFolderPath = this.state.selectedFolderForAction;
    const isFolder = !selectedFile && !!selectedFolderPath;
    const currentName = selectedFile?.name || (selectedFolderPath ? this._getLastFolderSegment(selectedFolderPath) : '');
    const validation = validateRenameName(this.state.renameName, currentName);

    return (
      <div className={styles.folderDialogOverlay} role="dialog" aria-modal="true" aria-label={`Rename ${isFolder ? 'folder' : 'file'}`}>
        <form className={styles.folderDialog} onSubmit={(event) => { event.preventDefault(); this._renameSelectedItem().catch(() => undefined); }}>
          <div className={styles.folderDialogHeader}>
            <div>
              <span>RENAME {isFolder ? 'FOLDER' : 'FILE'}</span>
              <h2>Rename {escape(currentName)}</h2>
            </div>
            <button type="button" className={styles.closeButton} aria-label="Close rename" onClick={() => this._closeRenameDialog()}>×</button>
          </div>
          <label className={styles.folderInputLabel} htmlFor="rename-name">New name</label>
          <input
            id="rename-name"
            className={styles.folderNameInput}
            type="text"
            autoFocus
            maxLength={128}
            value={this.state.renameName}
            onChange={(event) => this.setState({ renameName: event.currentTarget.value })}
            aria-describedby="rename-name-hint"
          />
          <p id="rename-name-hint" className={styles.renameHint}>Names cannot contain \ / : * ? &quot; &lt; &gt; | # % {'{'} {'}'} ~ &amp;.</p>
          {!validation.valid && this.state.renameName.trim() && <p className={styles.renameError} role="alert">{validation.error}</p>}
          <div className={styles.folderDialogActions}>
            <button type="button" className={styles.folderCancelButton} onClick={() => this._closeRenameDialog()}>Cancel</button>
            <button type="submit" className={styles.folderCreateButton} disabled={this.state.isRenaming || !validation.valid}>
              {this.state.isRenaming ? 'Renaming…' : 'Rename'}
            </button>
          </div>
        </form>
      </div>
    );
  }

  private _renderConfirmationDialog(): React.ReactElement {
    const isDelete = this.state.confirmationDialog === 'delete';
    const isOcrIngest = this.state.confirmationDialog === 'ingest-ocr';
    const selectionCommandState = getSelectionCommandState(this.state.selectedFileUrls, this.state.selectedFolderPaths);
    const selectedFile = this._getSelectedFile();
    const selectedFolderPath = selectionCommandState.selectedFolderPath;
    const isFolder = !selectedFile && !!selectedFolderPath;
    const targetName = selectedFile?.name || (selectedFolderPath ? this._getLastFolderSegment(selectedFolderPath) : '');
    const title = isDelete
      ? selectionCommandState.total > 1 ? `Delete ${selectionCommandState.total} items?` : `Delete ${isFolder ? 'folder' : 'file'}?`
      : isOcrIngest
        ? 'Process OCR-required PDFs?'
        : 'Index supported documents?';
    const message = isDelete
      ? selectionCommandState.total > 1
        ? `${selectionCommandState.total} selected files and folders will be moved to the SharePoint recycle bin.`
        : `“${targetName}” will be moved to the SharePoint recycle bin.`
      : isOcrIngest
        ? `Process only OCR-required PDFs in ${this.props.documentLibraryName}. This is a separate, explicit action and does not run automatically.`
        : `Index supported PDF, Word, PowerPoint, and Excel files from ${this.props.documentLibraryName}. Scanned/image PDFs will be reported as OCR required.`;
    const confirmLabel = isDelete ? 'Move to recycle bin' : isOcrIngest ? 'Process OCR PDFs' : 'Start indexing';

    return (
      <div className={styles.folderDialogOverlay} role="dialog" aria-modal="true" aria-label={title}>
        <div className={styles.confirmationDialog}>
          <div className={styles.folderDialogHeader}>
            <div>
              <span>{isDelete ? 'CONFIRM DELETE' : isOcrIngest ? 'CONFIRM OCR PROCESSING' : 'CONFIRM INDEXING'}</span>
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
                } else if (isOcrIngest) {
                  this._confirmOcrLibraryIngest().catch(() => undefined);
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
            <strong>OCR PDF processing</strong>
            <span>{escape(this.state.ocrIngestionMessage)}</span>
            <strong>Current upload target</strong>
            <span>{escape(this._getCurrentFolderServerRelativeUrl())}</span>
            <strong>Loaded metadata</strong>
            <span>{this.state.documents.length} file(s), {this.state.folderPaths.length} folder path(s)</span>
          </div>
          <button
            className={styles.askButton}
            type="button"
            disabled={this.state.isIngesting || this.state.isOcrIngesting || this.state.ragStatus === 'error'}
            onClick={() => this._ingestLibrary()}
          >
            {this.state.isIngesting ? 'Indexing supported documents...' : 'Index supported documents now'}
          </button>
          <button
            className={styles.askButton}
            type="button"
            disabled={this.state.isIngesting || this.state.isOcrIngesting || this.state.ragStatus === 'error'}
            onClick={() => this._ingestOcrLibrary()}
          >
            {this.state.isOcrIngesting ? 'Processing OCR-required PDFs...' : 'Process OCR-required PDFs'}
          </button>
          <p className={styles.ocrLimitHint}>OCR applies only to PDF files up to 4 MB and 15 pages with printed text; it is not for signatures.</p>
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


  private _openRenameDialog(): void {
    const selectedFile = this._getSelectedFile();
    const selectedFolderPath = this.state.selectedFolderForAction;
    const currentName = selectedFile?.name || (selectedFolderPath ? this._getLastFolderSegment(selectedFolderPath) : '');

    if (!currentName) {
      this.setState({ commandMessage: 'Select a file or folder before renaming.' });
      return;
    }

    this.setState({ isRenameOpen: true, renameName: currentName });
  }

  private _closeRenameDialog(): void {
    if (!this.state.isRenaming) {
      this.setState({ isRenameOpen: false, renameName: '' });
    }
  }

  private async _renameSelectedItem(): Promise<void> {
    const selectedFile = this._getSelectedFile();
    const selectedFolderPath = this.state.selectedFolderForAction;
    const itemType: RenameItemType | undefined = selectedFile ? 'file' : selectedFolderPath ? 'folder' : undefined;
    const sourceServerRelativeUrl = selectedFile?.serverRelativeUrl || selectedFolderPath;
    const currentName = selectedFile?.name || (selectedFolderPath ? this._getLastFolderSegment(selectedFolderPath) : '');
    const validation = validateRenameName(this.state.renameName, currentName);

    if (!itemType || !sourceServerRelativeUrl) {
      this.setState({ isRenameOpen: false, commandMessage: 'Select a file or folder before renaming.' });
      return;
    }

    if (!validation.valid) {
      this.setState({ commandMessage: validation.error });
      return;
    }

    const navigationState: NavigationState = {
      selectedFolderPath: this.state.selectedFolderPath,
      selectedParentFolderName: this.state.selectedParentFolderName,
      selectedFileUrl: selectedFile ? this._getAbsoluteSharePointUrl(this._getRenamedPath(sourceServerRelativeUrl, validation.value)) : '',
      viewMode: this.state.viewMode
    };
    this.setState({ isRenaming: true, commandMessage: `Renaming ${currentName}...` });

    try {
      const endpoint = buildRenameEndpoint(this.props.siteUrl, itemType, sourceServerRelativeUrl, validation.value);
      const request = buildRenameRequest(itemType, validation.value);
      const response = await this.props.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, request);

      if (!response.ok) {
        throw new Error(`Rename failed for ${currentName} (${response.status} ${response.statusText}).`);
      }

      await this._loadLibraryFiles(navigationState);
      this.setState({
        isRenaming: false,
        isRenameOpen: false,
        renameName: '',
        selectedFolderForAction: '',
        isPreviewOpen: false,
        commandMessage: `Renamed ${itemType}: ${validation.value}`
      });
    } catch (error) {
      this.setState({
        isRenaming: false,
        commandMessage: (error as Error).message || 'Rename failed.'
      });
    }
  }

  private _getRenamedPath(sourceServerRelativeUrl: string, newName: string): string {
    const sourcePath = sourceServerRelativeUrl.replace(/\/+$/, '');
    return `${sourcePath.substring(0, sourcePath.lastIndexOf('/'))}/${newName}`;
  }

  private _getAbsoluteSharePointUrl(serverRelativeUrl: string): string {
    try {
      return `${new URL(this.props.siteUrl).origin}${serverRelativeUrl}`;
    } catch {
      return `${this.props.siteUrl.replace(/\/$/, '')}${serverRelativeUrl}`;
    }
  }

  private _triggerBrowserDownload(blob: Blob, fileName: string): void {
    const objectUrl = window.URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = objectUrl;
    link.download = fileName;
    link.style.display = 'none';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    window.setTimeout(() => window.URL.revokeObjectURL(objectUrl), 0);
  }

  private _getDownloadFileName(): string {
    const date = new Date().toISOString().slice(0, 10);
    const libraryName = (this.props.documentLibraryName || 'documents').replace(/[\\/:*?"<>|]/g, '-');
    return `${libraryName}-${date}.zip`;
  }

  private async _downloadSelectedFiles(): Promise<void> {
    const selectedFiles = this.state.selectedFileUrls
      .map((fileUrl) => this.state.documents.find((document) => document.url === fileUrl))
      .filter((document): document is ILegalDocument => !!document);

    if (selectedFiles.length === 0) {
      this.setState({ commandMessage: 'Select one or more files before downloading.' });
      return;
    }

    this.setState({
      isDownloading: true,
      commandMessage: selectedFiles.length === 1 ? `Downloading ${selectedFiles[0].name}...` : `Preparing ZIP for ${selectedFiles.length} selected files...`
    });

    try {
      const downloadedFiles: Array<{ name: string; blob: Blob }> = [];
      const usedNames: Record<string, number> = {};
      for (const document of selectedFiles) {
        const endpoint = `${this.props.siteUrl}/_api/web/GetFileByServerRelativeUrl('${this._escapeODataString(document.serverRelativeUrl)}')/$value`;
        const response = await this.props.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
        if (!response.ok) {
          throw new Error(`Download failed for ${document.name} (${response.status} ${response.statusText}).`);
        }

        const originalName = document.name;
        const duplicateNumber = (usedNames[originalName] || 0) + 1;
        usedNames[originalName] = duplicateNumber;
        const extensionIndex = originalName.lastIndexOf('.');
        const uniqueName = duplicateNumber === 1 ? originalName : extensionIndex > 0
          ? `${originalName.slice(0, extensionIndex)} (${duplicateNumber})${originalName.slice(extensionIndex)}`
          : `${originalName} (${duplicateNumber})`;
        downloadedFiles.push({ name: uniqueName, blob: await response.blob() });
      }

      if (downloadedFiles.length === 1) {
        this._triggerBrowserDownload(downloadedFiles[0].blob, downloadedFiles[0].name);
        this.setState({ commandMessage: `Download started: ${downloadedFiles[0].name}` });
      } else {
        const zip = new JSZip();
        downloadedFiles.forEach((file) => zip.file(file.name, file.blob));
        const zipBlob = await zip.generateAsync({ type: 'blob' });
        const zipName = this._getDownloadFileName();
        this._triggerBrowserDownload(zipBlob, zipName);
        this.setState({ commandMessage: `ZIP download started: ${zipName} (${downloadedFiles.length} files).` });
      }
    } catch (error) {
      this.setState({ commandMessage: (error as Error).message || 'Download failed.' });
    } finally {
      this.setState({ isDownloading: false });
    }
  }

  private _deleteSelectedItem(): void {
    const selectionCommandState = getSelectionCommandState(this.state.selectedFileUrls, this.state.selectedFolderPaths);

    if (!selectionCommandState.canDelete) {
      this.setState({ commandMessage: 'Select a file or folder before deleting.' });
      return;
    }

    this.setState({ confirmationDialog: 'delete' });
  }

  private async _confirmDeleteSelectedItem(): Promise<void> {
    const selectedFiles = this.state.selectedFileUrls
      .map((fileUrl) => this.state.documents.find((document) => document.url === fileUrl))
      .filter((document): document is ILegalDocument => !!document);
    const deleteTargets = [
      ...selectedFiles.map((document) => ({ name: document.name, url: document.serverRelativeUrl, isFolder: false })),
      ...this.state.selectedFolderPaths.map((folderPath) => ({ name: this._getLastFolderSegment(folderPath), url: folderPath, isFolder: true }))
    ];

    if (deleteTargets.length === 0) {
      this.setState({ commandMessage: 'Select a file or folder before deleting.' });
      return;
    }

    const navigationState: NavigationState = {
      selectedParentFolderName: this.state.selectedParentFolderName,
      selectedFolderPath: this.state.selectedFolderPath,
      selectedFileUrl: '',
      viewMode: this.state.viewMode
    };
    this.setState({ commandMessage: `Deleting ${deleteTargets.length} selected item(s)...` });

    try {
      for (const target of deleteTargets) {
        const api = target.isFolder ? 'GetFolderByServerRelativeUrl' : 'GetFileByServerRelativeUrl';
        const endpoint = `${this.props.siteUrl}/_api/web/${api}('${this._escapeODataString(target.url)}')/recycle()`;
        const response = await this.props.spHttpClient.post(endpoint, SPHttpClient.configurations.v1, {
          headers: { Accept: 'application/json;odata=nometadata' }
        });

        if (!response.ok) {
          throw new Error(`Delete failed for ${target.name} (${response.status} ${response.statusText}).`);
        }
      }

      await this._loadLibraryFiles(navigationState);
      this.setState({
        isPreviewOpen: false,
        selectedFolderForAction: '',
        selectedFileUrls: [],
        selectedFolderPaths: [],
        answer: '',
        citations: [],
        suggestedActions: [],
        requestId: '',
        commandMessage: `Moved ${deleteTargets.length} item(s) to the recycle bin.`
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
    const displayColumns = this.props.displayColumns;
    const gridTemplateColumns = `34px minmax(260px, 1fr)${displayColumns.map(() => ' minmax(116px, 150px)').join('')}`;

    return (
      <div className={styles.tableWrap}>
        <div className={styles.tableHeader} style={{ gridTemplateColumns }}>
          <span />
          <span>Name / Category</span>
          {displayColumns.map((column) => <span key={column.internalName}>{escape(column.title)}</span>)}
        </div>
        {this._getParentFolderNames().map((parentName) => {
          const files = this._getDocumentsForParentFolder(parentName);
          const latest = files[0]?.lastModified || '';
          const folderPath = this._getServerRelativePathForParent(parentName);
          const isFolderSelected = this.state.selectedFolderPaths.indexOf(folderPath) >= 0;

          return (
            <div className={isFolderSelected ? styles.tableRowActive : styles.tableRow} style={{ gridTemplateColumns }} key={parentName}>
              <button
                className={isFolderSelected ? styles.selectCircleActive : styles.selectCircleButton}
                type="button"
                aria-label={`Select folder ${parentName}`}
                onClick={() => this._toggleFolderSelection(folderPath)}
              />
              <div className={styles.nameCellActions}>
                <button className={styles.nameCellButton} type="button" onClick={() => this._selectParentFolder(parentName)}>
                  <span className={styles.folderGlyph}>📁</span>{escape(parentName)}
                </button>
                <button className={styles.rowMoreButton} type="button" aria-label={`Rename folder ${parentName}`} title="Rename folder" onClick={() => this._openRenameForFolder(folderPath)}>⋯</button>
              </div>
              {displayColumns.map((column) => <span key={column.internalName}>{escape(this._formatFolderColumnValue(column.internalName, files.length, latest))}</span>)}
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
          const isFolderSelected = this.state.selectedFolderPaths.indexOf(folderPath) >= 0;

          return (
            <div className={isFolderSelected ? styles.tableRowActive : styles.tableRow} key={folderPath}>
              <button
                className={isFolderSelected ? styles.selectCircleActive : styles.selectCircleButton}
                type="button"
                aria-label={`Select folder ${this._getLastFolderSegment(folderPath)}`}
                onClick={() => this._toggleFolderSelection(folderPath)}
              />
              <div className={styles.nameCellActions}>
                <button className={styles.nameCellButton} type="button" onClick={() => this._selectFolder(folderPath)}>
                  <span className={styles.folderGlyph}>📁</span>{escape(this._getLastFolderSegment(folderPath))}
                </button>
                <button className={styles.rowMoreButton} type="button" aria-label={`Rename folder ${this._getLastFolderSegment(folderPath)}`} title="Rename folder" onClick={() => this._openRenameForFolder(folderPath)}>⋯</button>
              </div>
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
    const isSelected = this.state.selectedFileUrls.indexOf(doc.url) >= 0;

    return (
      <div key={doc.url} className={isSelected ? styles.tableRowActive : styles.tableRow}>
        <button
          className={isSelected ? styles.selectCircleActive : styles.selectCircleButton}
          type="button"
          aria-label={`Select ${doc.name}`}
          onClick={() => this._toggleFileSelection(doc)}
        />
        <div className={styles.nameCellActions}>
          <button
            className={styles.nameCellButton}
            type="button"
            title={`Preview ${doc.name}`}
            onClick={() => this._selectFile(doc, true)}
          >
            <span className={styles.fileGlyph}>{this._renderFileGlyph(doc.type)}</span>
            <span>{escape(doc.name)}</span>
          </button>
          <button className={styles.rowMoreButton} type="button" aria-label={`Rename ${doc.name}`} title="Rename file" onClick={() => this._openRenameForFile(doc)}>⋯</button>
        </div>
        <span>{escape(showModifiedBy ? doc.lastModified : doc.type)}</span>
        <span>{escape(showModifiedBy ? doc.modifiedBy : doc.lastModified)}</span>
      </div>
    );
  }

  private _renderFileRows(selectedFile: ILegalDocument | undefined): React.ReactElement {
    const visibleDocuments = this._getVisibleDocuments();
    const displayColumns: IDocumentLibraryDisplayColumn[] = this.props.displayColumns;
    const gridTemplateColumns = `34px minmax(260px, 1fr)${displayColumns.map(() => ' minmax(116px, 150px)').join('')}`;

    return (
      <div className={styles.tableWrap}>
        {this._renderFileBreadcrumb(visibleDocuments.length)}
        <div className={styles.tableHeader} style={{ gridTemplateColumns }}>
          <span />
          <span>Name</span>
          {displayColumns.map((column) => <span key={column.internalName}>{escape(column.title)}</span>)}
        </div>
        {visibleDocuments.length === 0 && <div className={styles.noticeBox}>This folder is now empty. You are still in the same folder location.</div>}
        {visibleDocuments.map((doc) => {
          const isSelected = this.state.selectedFileUrls.indexOf(doc.url) >= 0;
          return <div key={doc.url} className={isSelected ? styles.tableRowActive : styles.tableRow} style={{ gridTemplateColumns }}>
            <button className={isSelected ? styles.selectCircleActive : styles.selectCircleButton} type="button" aria-label={`Select ${doc.name}`} onClick={() => this._toggleFileSelection(doc)} />
            <div className={styles.nameCellActions}>
              <button className={styles.nameCellButton} type="button" title={`Preview ${doc.name}`} onClick={() => this._selectFile(doc, true)}><span className={styles.fileGlyph}>{this._renderFileGlyph(doc.type)}</span><span>{escape(doc.name)}</span></button>
              <button className={styles.rowMoreButton} type="button" aria-label={`Rename ${doc.name}`} title="Rename file" onClick={() => this._openRenameForFile(doc)}>⋯</button>
            </div>
            {displayColumns.map((column) => <span key={column.internalName}>{escape(this._formatConfiguredColumnValue(doc, column.internalName))}</span>)}
          </div>;
        })}
      </div>
    );
  }

  private _formatFolderColumnValue(internalName: string, fileCount: number, latestModified: string): string {
    if (internalName === 'File_x0020_Type') return 'Folder';
    if (internalName === 'Modified') return latestModified || '—';
    if (internalName === 'Editor' || internalName === 'Created') return '—';
    return internalName === 'ItemChildCount' ? String(fileCount) : '—';
  }

  private _formatConfiguredColumnValue(doc: ILegalDocument, internalName: string): string {
    const value = doc.fieldValues[internalName];
    if (value === undefined || value === null || value === '') return '—';
    if (internalName === 'Modified' || internalName === 'Created') return this._formatDate(String(value));
    if (internalName === 'Editor') return doc.modifiedBy;
    if (typeof value === 'object') {
      const person = value as { Title?: string };
      return person.Title || '—';
    }
    return String(value);
  }

  private async _loadLibraryFiles(navigationState?: Partial<NavigationState>): Promise<void> {
    const libraryName = this.props.documentLibraryName;
    if (!libraryName) {
      this.setState({
        documents: [],
        folderPaths: [],
        libraryRootServerRelativeUrl: '',
        selectedFolderPath: '',
        selectedParentFolderName: '',
        selectedFileUrl: '',
        viewMode: 'parents',
        libraryStatus: 'empty',
        libraryMessage: 'Select a document library in the Web Part properties to load its folders and files.'
      });
      return;
    }

    this.setState({ libraryStatus: 'loading', libraryMessage: `Loading ${libraryName}...` });

    try {
      const escapedLibraryName = libraryName.replace(/'/g, "''");
      const libraryEndpoint = `${this.props.siteUrl}/_api/web/lists/getByTitle('${escapedLibraryName}')?$select=RootFolder/ServerRelativeUrl&$expand=RootFolder`;
      const libraryResponse: SPHttpClientResponse = await this.props.spHttpClient.get(libraryEndpoint, SPHttpClient.configurations.v1);
      if (!libraryResponse.ok) {
        throw new Error(`Could not read library '${libraryName}' (${libraryResponse.status} ${libraryResponse.statusText}).`);
      }

      const libraryMetadata = await libraryResponse.json() as ISharePointLibraryMetadata;
      const libraryRootServerRelativeUrl = libraryMetadata.RootFolder?.ServerRelativeUrl || '';
      if (!libraryRootServerRelativeUrl) {
        throw new Error(`SharePoint did not return a root folder for '${libraryName}'.`);
      }

      const selectFields = buildListItemSelect(this.props.displayColumns.map((column) => column.internalName));
      const endpoint = `${this.props.siteUrl}/_api/web/lists/getByTitle('${escapedLibraryName}')/items?$select=${selectFields.join(',')}&$expand=Editor&$orderby=FileDirRef asc,Modified desc&$top=200`;
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
        libraryRootServerRelativeUrl,
        selectedFolderPath: navigationState?.selectedFolderPath || '',
        selectedParentFolderName: navigationState?.selectedParentFolderName || '',
        selectedFileUrl: navigationState?.selectedFileUrl || '',
        selectedFileUrls: navigationState?.selectedFileUrl ? [navigationState.selectedFileUrl] : [],
        selectedFolderForAction: '',
        selectedFolderPaths: [],
        viewMode: navigationState?.viewMode || 'parents',
        libraryStatus: documents.length > 0 ? 'loaded' : 'empty',
        libraryMessage: documents.length > 0 ? `Loaded ${documents.length} file(s) across ${folderPaths.length} folder(s).` : `${libraryName} exists, but no files were found.`
      });
    } catch (error) {
      this.setState({
        documents: [],
        folderPaths: [],
        libraryRootServerRelativeUrl: '',
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
      modifiedBy: item.Editor?.Title || this.props.userDisplayName || 'Unknown',
      fieldValues: item
    };
  }


  private _selectParentFolder(parentFolderName: string): void {
    this.setState({
      selectedParentFolderName: parentFolderName,
      selectedFolderPath: '',
      selectedFileUrl: '',
      selectedFolderForAction: '',
      selectedFileUrls: [],
      selectedFolderPaths: [],
      viewMode: 'children',
      answer: '',
      citations: [],
      suggestedActions: [],
      requestId: '',
      error: ''
    });
  }

  private _selectFolder(folderPath: string): void {
    const folderSelection = getFolderOpenSelection();

    this.setState({
      selectedFolderPath: folderPath,
      selectedFileUrl: folderSelection.selectedFileUrl,
      selectedFolderForAction: '',
      selectedFileUrls: folderSelection.selectedFileUrls,
      selectedFolderPaths: folderSelection.selectedFolderPaths,
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
    if (this.state.libraryRootServerRelativeUrl) {
      return this.state.libraryRootServerRelativeUrl;
    }

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

  private _renderFileGlyph(fileType: string): React.ReactNode {
    const iconKind = getFileIconKind(fileType);
    if (iconKind === 'pdf') {
      return (
        <svg className={styles.pdfFileGlyph} viewBox="0 0 24 24" aria-hidden="true" focusable="false">
          <path d="M5 2.5h9l5 5v14H5z" fill="#fff" stroke="#d13438" strokeWidth="1.5" strokeLinejoin="round" />
          <path d="M14 2.5v5h5" fill="#f8d7da" stroke="#d13438" strokeWidth="1.5" strokeLinejoin="round" />
          <rect x="5" y="13" width="14" height="6" rx="1" fill="#d13438" />
          <text x="12" y="17.4" textAnchor="middle" fill="#fff" fontSize="4.5" fontWeight="700" fontFamily="Arial, sans-serif">PDF</text>
        </svg>
      );
    }

    if (iconKind === 'word' || iconKind === 'excel' || iconKind === 'powerpoint') {
      const color = iconKind === 'word' ? '#185abd' : iconKind === 'excel' ? '#107c41' : '#c43e1c';
      const accent = iconKind === 'word' ? '#2b88d8' : iconKind === 'excel' ? '#21a366' : '#d24726';
      const letter = iconKind === 'word' ? 'W' : iconKind === 'excel' ? 'X' : 'P';
      const grid = iconKind === 'excel';
      return (
        <svg className={styles.officeFileGlyph} viewBox="0 0 24 24" aria-hidden="true" focusable="false">
          <path d="M7 2.5h8l4 4v15H7z" fill="#f8fbff" stroke={accent} strokeWidth="1.25" strokeLinejoin="round" />
          <path d="M15 2.5v4h4" fill="#dcecff" stroke={accent} strokeWidth="1.25" strokeLinejoin="round" />
          {grid && <><path d="M11.5 9h5M11.5 12h5M11.5 15h5M14 9v7" stroke="#21a366" strokeWidth="0.8" /></>}
          {!grid && <path d="M11 11h5.5M11 14h5.5M11 17h4" stroke={accent} strokeWidth="1.1" strokeLinecap="round" />}
          <rect x="3" y="7" width="10" height="10" rx="0.9" fill={color} />
          <text x="8" y="13.8" textAnchor="middle" fill="#fff" fontSize="6" fontWeight="700" fontFamily="Segoe UI, Arial, sans-serif">{letter}</text>
        </svg>
      );
    }

    return this._getFileGlyph(fileType);
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

  private _getRagOcrIngestEndpoint(): string {
    return getOcrIngestEndpoint(this.props.functionEndpoint);
  }

  private _getRagAnswerEndpoint(): string {
    return this.props.functionEndpoint
      .replace('/api/chat', '/api/rag/answer')
      .replace('/api/ask', '/api/rag/answer');
  }

  private _ingestLibrary(): void {
    this.setState({ confirmationDialog: 'ingest-text' });
  }

  private _ingestOcrLibrary(): void {
    this.setState({ confirmationDialog: 'ingest-ocr' });
  }

  private _getDynamicLibraryScopeRequest(): { siteUrl: string; libraryName: string } {
    return {
      siteUrl: this.props.siteUrl,
      libraryName: this.props.documentLibraryName
    };
  }

  private async _confirmLibraryIngest(): Promise<void> {
    this.setState({ isIngesting: true, ingestionMessage: `Indexing supported documents from ${this.props.documentLibraryName}...` });
    try {
      const apiClient = await this.props.aadHttpClientFactory.getClient(this.props.functionApiResource);
      const response = await apiClient.fetch(this._getRagIngestEndpoint(), AadHttpClient.configurations.v1, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(this._getDynamicLibraryScopeRequest())
      });
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

  private async _confirmOcrLibraryIngest(): Promise<void> {
    this.setState({ isOcrIngesting: true, ocrIngestionMessage: `Processing OCR-required PDFs from ${this.props.documentLibraryName}...` });
    try {
      const apiClient = await this.props.aadHttpClientFactory.getClient(this.props.functionApiResource);
      const response = await apiClient.fetch(this._getRagOcrIngestEndpoint(), AadHttpClient.configurations.v1, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(this._getDynamicLibraryScopeRequest())
      });
      const payload = await response.json() as {
        error?: string;
        detail?: string;
        ingestion?: IOcrIngestionResult;
      };
      if (!response.ok || !payload.ingestion) {
        throw new Error(payload.detail || payload.error || `OCR PDF processing returned ${response.status}.`);
      }

      const ingestion = payload.ingestion;
      this.setState({
        isOcrIngesting: false,
        ragStatus: ingestion.indexedChunks > 0 ? 'ready' : this.state.ragStatus,
        ocrIngestionMessage: formatOcrIngestionMessage(ingestion)
      });
      await this._checkRagHealth();
    } catch (error) {
      this.setState({
        isOcrIngesting: false,
        ocrIngestionMessage: (error as Error).message || 'OCR PDF processing failed.'
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
    const useLibraryRag = this.state.surfaceMode === 'document-library';
    const contextFiles = this.state.documents.filter((document) => this.state.selectedFileUrls.indexOf(document.url) >= 0);
    const contextSnippets = contextFiles.map(doc => doc.snippet);
    const activeFolderPath = selectedFile?.folderPath || this._getActiveFolderPathForAsk();
    const isCurrentFolderScope = contextFiles.length === 0
      && activeFolderPath.toLowerCase() !== this._getLibraryRootServerRelativeUrl().toLowerCase();

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
      const selectedDocument = selectedFile && !useLibraryRag ? await this._getSelectedPdfPayload(selectedFile) : undefined;
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
          selectedDocumentUrl: contextFiles.length === 1 ? contextFiles[0].url : undefined,
          selectedDocumentNames: contextFiles.map((document) => document.name),
          selectedItems: [],
          documentSnippets: contextSnippets,
          selectedDocument,
          knowledgeScope: useLibraryRag
            ? contextFiles.length > 0
              ? 'selected-file-rag'
              : isCurrentFolderScope
                ? 'current-folder-rag'
                : 'library-wide-rag'
            : 'legal-document-library'
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










