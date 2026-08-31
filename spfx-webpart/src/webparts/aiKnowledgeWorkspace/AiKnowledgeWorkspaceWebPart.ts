import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
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

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

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
        documentLibraryName: this.properties.documentLibraryName || 'Litigation Documents',
        spHttpClient: this.context.spHttpClient
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
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
                PropertyPaneTextField('documentLibraryName', {
                  label: 'Document library name',
                  description: 'Default: Litigation Documents'
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
