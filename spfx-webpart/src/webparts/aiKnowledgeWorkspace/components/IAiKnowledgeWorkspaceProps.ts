import { AadHttpClientFactory, SPHttpClient } from '@microsoft/sp-http';
import { IDocumentLibraryDisplayColumn } from './documentLibraryColumns';

export interface IAiKnowledgeWorkspaceProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;
  siteTitle: string;
  siteUrl: string;
  pageUrl: string;
  functionEndpoint: string;
  functionApiResource: string;
  documentLibraryName: string;
  displayColumns: IDocumentLibraryDisplayColumn[];
  spHttpClient: SPHttpClient;
  aadHttpClientFactory: AadHttpClientFactory;
}
