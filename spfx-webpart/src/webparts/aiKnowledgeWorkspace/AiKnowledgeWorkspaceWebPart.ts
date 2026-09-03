import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  type IPropertyPaneDropdownOption,
  PropertyPaneDropdown,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { SPHttpClient } from '@microsoft/sp-http';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'AiKnowledgeWorkspaceWebPartStrings';
import AiKnowledgeWorkspace from './components/AiKnowledgeWorkspace';
import { IAiKnowledgeWorkspaceProps } from './components/IAiKnowledgeWorkspaceProps';

export interface IAiKnowledgeWorkspaceWebPartProps {
  description: string;
  functionEndpoint: string;
  documentLibraryName: string;
}

export default class AiKnowledgeWorkspaceWebPart extends BaseClientSideWebPart<IAiKnowledgeWorkspaceWebPartProps> {

  private static readonly _excludedDocumentLibraryTitles: Set<string> = new Set([
    'form templates',
    'site assets',
    'style library'
  ]);

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _documentLibraryOptions: IPropertyPaneDropdownOption[] = [];
  private _isLoadingDocumentLibraries: boolean = false;
  private _documentLibrariesLoaded: boolean = false;

  public render(): void {
    const element: React.ReactElement<IAiKnowledgeWorkspaceProps> = React.createElement(
      AiKnowledgeWorkspace,
      {
        description: this.properties.description || 'Reusable SharePoint AI Knowledge Workspace',
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        userDisplayName: this.context.pageContext.user.displayName,
        siteTitle: this.context.pageContext.web.title,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        pageUrl: window.location.href,
        functionEndpoint: this.properties.functionEndpoint || 'http://localhost:7072/api/chat',
        functionApiResource: 'api://d3df04c0-e580-4684-877a-0733204e7e2e',
        documentLibraryName: this.properties.documentLibraryName,
        spHttpClient: this.context.spHttpClient,
        aadHttpClientFactory: this.context.aadHttpClientFactory
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._loadDocumentLibraryOptions().catch(() => undefined);
  }

  private async _loadDocumentLibraryOptions(): Promise<void> {
    if (this._isLoadingDocumentLibraries || this._documentLibrariesLoaded) {
      return;
    }

    this._isLoadingDocumentLibraries = true;
    this.context.propertyPane.refresh();

    try {
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Title,Id&$filter=BaseTemplate eq 101 and Hidden eq false&$orderby=Title`;
      const response = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
      if (!response.ok) {
        throw new Error(`Document library lookup failed (${response.status} ${response.statusText}).`);
      }

      const payload = await response.json() as { value?: Array<{ Title: string; Id: string }> };
      this._documentLibraryOptions = (payload.value || [])
        .filter((library) => !!library.Title && !AiKnowledgeWorkspaceWebPart._excludedDocumentLibraryTitles.has(library.Title.trim().toLocaleLowerCase()))
        .map((library) => ({ key: library.Title, text: library.Title }));
      this._documentLibrariesLoaded = true;
    } finally {
      this._isLoadingDocumentLibraries = false;
      this.context.propertyPane.refresh();
    }
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Legal Document Library AI Assistant settings'
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneDropdown('documentLibraryName', {
                  label: 'Document Library',
                  options: this._documentLibraryOptions.length > 0
                    ? this._documentLibraryOptions
                    : [{ key: '', text: this._isLoadingDocumentLibraries ? 'Loading document libraries…' : 'No document libraries found' }],
                  disabled: this._isLoadingDocumentLibraries || !this._documentLibrariesLoaded
                }),
                PropertyPaneTextField('functionEndpoint', {
                  label: 'AI Function endpoint',
                  description: 'Local debug default: http://localhost:7072/api/chat'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
