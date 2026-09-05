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
import type { IDocumentLibraryDisplayColumn } from './components/documentLibraryColumns';
import { buildSelectedDisplayColumns, knownDocumentLibraryDisplayColumns } from './components/documentLibraryColumns';

export interface IAiKnowledgeWorkspaceWebPartProps {
  description: string;
  functionEndpoint: string;
  documentLibraryName: string;
  displayColumn1: string;
  displayColumn2: string;
  displayColumn3: string;
  displayColumn4: string;
  displayColumn5: string;
}

export default class AiKnowledgeWorkspaceWebPart extends BaseClientSideWebPart<IAiKnowledgeWorkspaceWebPartProps> {
  private static readonly _excludedDocumentLibraryTitles: Set<string> = new Set(['form templates', 'site assets', 'style library']);
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _documentLibraryOptions: IPropertyPaneDropdownOption[] = [];
  private _isLoadingDocumentLibraries: boolean = false;
  private _documentLibrariesLoaded: boolean = false;
  private _displayFieldOptions: IPropertyPaneDropdownOption[] = [];
  private _displayFieldsLoadedForLibrary: string = '';
  private _isLoadingDisplayFields: boolean = false;

  public render(): void {
    const element: React.ReactElement<IAiKnowledgeWorkspaceProps> = React.createElement(AiKnowledgeWorkspace, {
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
      displayColumns: this._getDisplayColumns(),
      spHttpClient: this.context.spHttpClient,
      aadHttpClientFactory: this.context.aadHttpClientFactory
    });
    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage()
      .then((message) => { this._environmentMessage = message; })
      .then(() => this._loadDisplayFieldOptions(this.properties.documentLibraryName));
  }

  protected onPropertyPaneConfigurationStart(): void {
    this._loadDocumentLibraryOptions().catch(() => undefined);
    this._loadDisplayFieldOptions(this.properties.documentLibraryName).catch(() => undefined);
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    if (propertyPath === 'documentLibraryName' && oldValue !== newValue) {
      this._displayFieldOptions = [];
      this._displayFieldsLoadedForLibrary = '';
      this._clearDisplayColumnSlots();
      this._loadDisplayFieldOptions(newValue as string).catch(() => undefined);
    }
  }

  private async _loadDocumentLibraryOptions(): Promise<void> {
    if (this._isLoadingDocumentLibraries || this._documentLibrariesLoaded) return;
    this._isLoadingDocumentLibraries = true;
    this.context.propertyPane.refresh();
    try {
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$select=Title,Id&$filter=BaseTemplate eq 101 and Hidden eq false&$orderby=Title`;
      const response = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
      if (!response.ok) throw new Error(`Document library lookup failed (${response.status} ${response.statusText}).`);
      const payload = await response.json() as { value?: Array<{ Title: string; Id: string }> };
      this._documentLibraryOptions = (payload.value || []).filter((library) => !!library.Title && !AiKnowledgeWorkspaceWebPart._excludedDocumentLibraryTitles.has(library.Title.trim().toLocaleLowerCase())).map((library) => ({ key: library.Title, text: library.Title }));
      this._documentLibrariesLoaded = true;
    } finally {
      this._isLoadingDocumentLibraries = false;
      this.context.propertyPane.refresh();
    }
  }

  private async _loadDisplayFieldOptions(libraryName: string): Promise<void> {
    if (!libraryName || this._isLoadingDisplayFields || this._displayFieldsLoadedForLibrary === libraryName) return;
    this._isLoadingDisplayFields = true;
    this.context.propertyPane.refresh();
    try {
      const escapedLibraryName = libraryName.replace(/'/g, "''");
      const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getByTitle('${escapedLibraryName}')/fields?$select=Title,InternalName,Hidden,ReadOnlyField,Sealed,FieldTypeKind&$filter=Hidden eq false&$orderby=Title`;
      const response = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
      if (!response.ok) throw new Error(`Column lookup failed (${response.status} ${response.statusText}).`);
      const payload = await response.json() as { value?: Array<{ Title: string; InternalName: string; ReadOnlyField?: boolean; Sealed?: boolean; FieldTypeKind?: number }> };
      const excluded: Record<string, boolean> = { Attachments: true, ContentType: true, ContentTypeId: true, Edit: true, FileLeafRef: true, FileRef: true, FileDirRef: true, FSObjType: true, LinkFilename: true, LinkTitle: true, LinkTitleNoMenu: true, _UIVersionString: true };
      const supportedFieldTypes: Record<number, boolean> = { 2: true, 3: true, 4: true, 6: true, 8: true, 9: true };
      const fields = (payload.value || []).filter((field) => !!field.Title && !!field.InternalName && !excluded[field.InternalName] && !field.InternalName.startsWith('_') && !field.Sealed && supportedFieldTypes[field.FieldTypeKind || 0]);
      const builtInFields = knownDocumentLibraryDisplayColumns.map((field) => ({ InternalName: field.internalName, Title: field.title }));
      const displayableFields = builtInFields.concat(fields.filter((field) => !builtInFields.some((builtIn) => builtIn.InternalName === field.InternalName)));
      this._displayFieldOptions = [{ key: 'FileLeafRef', text: 'Name (FileLeafRef)' }].concat(displayableFields.filter((field) => field.InternalName !== 'FileLeafRef').map((field) => ({ key: field.InternalName, text: `${field.Title} (${field.InternalName})` })));
      this._displayFieldsLoadedForLibrary = libraryName;
    } finally {
      this._isLoadingDisplayFields = false;
      this.context.propertyPane.refresh();
      this.render();
    }
  }

  private _getDisplayColumns(): IDocumentLibraryDisplayColumn[] {
    const discoveredFields = this._displayFieldOptions
      .filter((option) => option.key !== 'FileLeafRef')
      .map((option) => ({ internalName: String(option.key), title: option.text.replace(/ \([^)]*\)$/, '') }));

    return buildSelectedDisplayColumns([
      this.properties.displayColumn1,
      this.properties.displayColumn2,
      this.properties.displayColumn3,
      this.properties.displayColumn4,
      this.properties.displayColumn5
    ], discoveredFields);
  }

  private _clearDisplayColumnSlots(): void {
    this.properties.displayColumn1 = '';
    this.properties.displayColumn2 = '';
    this.properties.displayColumn3 = '';
    this.properties.displayColumn4 = '';
    this.properties.displayColumn5 = '';
  }

  private _getDisplayColumnDropdownOptions(): IPropertyPaneDropdownOption[] {
    if (this._isLoadingDisplayFields) return [{ key: '', text: 'Loading library columns…' }];
    return [{ key: '', text: '— Do not display —' } as IPropertyPaneDropdownOption, ...this._displayFieldOptions.filter((option) => option.key !== 'FileLeafRef')];
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) return this.context.sdks.microsoftTeams.teamsJs.app.getContext().then(context => {
      switch (context.app.host.name) {
        case 'Office': return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
        case 'Outlook': return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
        case 'Teams': case 'TeamsModern': return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
        default: return strings.UnknownEnvironment;
      }
    });
    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) return;
    this._isDarkTheme = !!currentTheme.isInverted;
    const { semanticColors } = currentTheme;
    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }
  }
  protected onDispose(): void { ReactDom.unmountComponentAtNode(this.domElement); }
  protected get dataVersion(): Version { return Version.parse('1.0'); }
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return { pages: [{ header: { description: this.properties.documentLibraryName || 'Document Library' }, groups: [{ groupName: '', groupFields: [
      PropertyPaneDropdown('documentLibraryName', { label: 'Document Library', options: this._documentLibraryOptions.length > 0 ? this._documentLibraryOptions : [{ key: '', text: this._isLoadingDocumentLibraries ? 'Loading document libraries…' : 'No document libraries found' }], disabled: this._isLoadingDocumentLibraries || !this._documentLibrariesLoaded }),
      PropertyPaneDropdown('displayColumn1', { label: 'Table column 1', options: this._getDisplayColumnDropdownOptions(), disabled: !this.properties.documentLibraryName || this._isLoadingDisplayFields }),
      PropertyPaneDropdown('displayColumn2', { label: 'Table column 2', options: this._getDisplayColumnDropdownOptions(), disabled: !this.properties.documentLibraryName || this._isLoadingDisplayFields }),
      PropertyPaneDropdown('displayColumn3', { label: 'Table column 3', options: this._getDisplayColumnDropdownOptions(), disabled: !this.properties.documentLibraryName || this._isLoadingDisplayFields }),
      PropertyPaneDropdown('displayColumn4', { label: 'Table column 4', options: this._getDisplayColumnDropdownOptions(), disabled: !this.properties.documentLibraryName || this._isLoadingDisplayFields }),
      PropertyPaneDropdown('displayColumn5', { label: 'Table column 5', options: this._getDisplayColumnDropdownOptions(), disabled: !this.properties.documentLibraryName || this._isLoadingDisplayFields }),
      PropertyPaneTextField('functionEndpoint', { label: 'AI Function endpoint', description: 'Local debug default: http://localhost:7072/api/chat' })
    ] }] }] };
  }
}
