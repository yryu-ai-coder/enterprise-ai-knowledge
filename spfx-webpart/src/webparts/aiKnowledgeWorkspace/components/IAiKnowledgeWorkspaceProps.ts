import { AadHttpClientFactory, SPHttpClient } from '@microsoft/sp-http';

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
  spHttpClient: SPHttpClient;
  aadHttpClientFactory: AadHttpClientFactory;
}
