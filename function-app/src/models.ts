export interface SharePointFileContext {
  name: string;
  url?: string;
  uniqueId?: string;
  fileType?: string;
  libraryTitle?: string;
  lastModified?: string;
  snippet?: string;
}

export interface SharePointListItemContext {
  id: string | number;
  title?: string;
  listTitle?: string;
  fields?: Record<string, unknown>;
}

export type SharePointContextType = 'document-library' | 'sharepoint-list' | 'site-page' | 'current-site';

export interface SelectedDocumentContent {
  name: string;
  fileType: string;
  contentBase64: string;
}

export interface ChatRequest {
  question: string;
  scenario?: string;
  mode?: string;
  contextType?: SharePointContextType;
  siteUrl?: string;
  listId?: string;
  listTitle?: string;
  libraryName?: string;
  folderPath?: string;
  pageUrl?: string;
  pageTitle?: string;
  pageContext?: {
    webTitle?: string;
    userEmail?: string;
  };
  selectedFiles?: SharePointFileContext[];
  selectedItems?: SharePointListItemContext[];
  documentSnippets?: string[];
  selectedDocument?: SelectedDocumentContent;
  conversationId?: string;
  knowledgeScope?: string;
}

export interface Citation {
  title: string;
  url: string;
  snippet?: string;
}

export interface ChatResponse {
  answer: string;
  citations: Citation[];
  provider: string;
  requestId: string;
  status: 'success' | 'error';
  suggestedActions?: string[];
  metadata?: Record<string, unknown>;
  error?: string;
}
