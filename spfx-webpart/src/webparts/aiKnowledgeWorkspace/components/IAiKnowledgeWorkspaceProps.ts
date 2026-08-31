import { SPHttpClient } from '@microsoft/sp-http';

export interface IAiKnowledgeWorkspaceProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  userDisplayName: string;
  siteTitle: string;
  siteUrl: string;
  pageUrl: string;
  functionEndpoint: string;
  documentLibraryName: string;
  spHttpClient: SPHttpClient;
}
