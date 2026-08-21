import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneLabel
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'SubscriptionUpdatesWebPartStrings';
import SubscriptionUpdates from './components/SubscriptionUpdates';
import { ISubscriptionUpdatesProps } from './components/ISubscriptionUpdatesProps';

import { spfi, SPFx } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/search";
import "@pnp/sp/items";
import "@pnp/sp/site-users";

export interface ISubscriptionUpdatesWebPartProps {
  siteUrl: string;
  listId: string;
  subscribeFor: string;
  subscribedBy: string;
  subscribedOn: string;
  referenceId: string;
  referenceIdColumn?: string;
  customTitle: string;
  showTitle: boolean;
  descriptionText: string;
  successNotification: string;
  buttonText: string;
  showSuccessNotification: boolean;
  titleFontSize: number;
  descriptionFontSize: number;
  choicesFontSize: number;
  notificationFontSize: number;
  buttonFontSize: number;
  choicesLabel: string;
  titleBarOpacity: number;
}

export default class SubscriptionUpdatesWebPart extends BaseClientSideWebPart<ISubscriptionUpdatesWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private _sp: any;
  private _sites: IPropertyPaneDropdownOption[] = [];
  private _lists: IPropertyPaneDropdownOption[] = [];
  private _choiceColumns: IPropertyPaneDropdownOption[] = [];
  private _personColumns: IPropertyPaneDropdownOption[] = [];
  private _dateColumns: IPropertyPaneDropdownOption[] = [];
  private _allColumns: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    const siteSp = this.properties.siteUrl ? spfi(this.properties.siteUrl).using(SPFx(this.context)) : this._sp;

    const element: React.ReactElement<ISubscriptionUpdatesProps> = React.createElement(
      SubscriptionUpdates,
      {
        siteUrl: this.properties.siteUrl,
        listId: this.properties.listId,
        subscribeForColumn: this.properties.subscribeFor,
        subscribedByColumn: this.properties.subscribedBy,
        subscribedOnColumn: this.properties.subscribedOn,
        referenceIdColumn: this.properties.referenceId,
        customTitle: this.properties.customTitle,
        showTitle: this.properties.showTitle,
        descriptionText: this.properties.descriptionText,
        successNotification: this.properties.successNotification,
        buttonText: this.properties.buttonText,
        showSuccessNotification: this.properties.showSuccessNotification !== false,
        titleFontSize: this.properties.titleFontSize || 24,
        descriptionFontSize: this.properties.descriptionFontSize || 14,
        notificationFontSize: this.properties.notificationFontSize || 14,
        buttonFontSize: this.properties.buttonFontSize || 14,
        choicesFontSize: this.properties.choicesFontSize || 14,
        choicesLabel: this.properties.choicesLabel || "Subscribe For",
        titleBarOpacity: this.properties.titleBarOpacity ?? 100,
        sp: siteSp,
        userDisplayName: this.context.pageContext.user.displayName,
        userEmail: this.context.pageContext.user.email
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    this._sp = spfi().using(SPFx(this.context));
    this._environmentMessage = await this._getEnvironmentMessage();
  }

  protected onPropertyPaneConfigurationStart(): void {
    const fetchInitialData = async (): Promise<void> => {
      if (this._sites.length === 0) {
        try {
          const response: SPHttpClientResponse = await this.context.spHttpClient.get(
            `${this.context.pageContext.web.absoluteUrl}/_api/search/query?querytext='contentclass:STS_Site contentclass:STS_Web'&selectproperties='Title,Path'&rowlimit=500`,
            SPHttpClient.configurations.v1
          );

          const searchResults: any = await response.json();
          const rows = searchResults.PrimaryQueryResult.RelevantResults.Table.Rows;

          this._sites = rows.map((row: any) => {
            const title = row.Cells.find((c: any) => c.Key === "Title").Value;
            const path = row.Cells.find((c: any) => c.Key === "Path").Value;
            return { key: path, text: title };
          });
        } catch (error) {
          console.error("Error fetching sites via REST API", error);
        }
      }

      // Load lists if site is already selected
      if (this.properties.siteUrl && this._lists.length === 0) {
        await this._fetchLists(this.properties.siteUrl);
      }

      // Load columns if list is already selected
      if (this.properties.listId && this._choiceColumns.length === 0) {
        await this._fetchColumns(this.properties.siteUrl, this.properties.listId);
      }

      this.context.propertyPane.refresh();
    };

    fetchInitialData().catch(err => console.error(err));
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    const handleFieldChange = async (): Promise<void> => {
      if (propertyPath === 'siteUrl' && newValue !== oldValue) {
        this.properties.listId = '';
        this.properties.subscribeFor = '';
        this.properties.subscribedBy = '';
        this.properties.subscribedOn = '';
        this._lists = [];
        this._choiceColumns = [];
        this._personColumns = [];
        this._dateColumns = [];
        await this._fetchLists(newValue);
      } else if (propertyPath === 'listId' && newValue !== oldValue) {
        this.properties.subscribeFor = '';
        this.properties.subscribedBy = '';
        this.properties.subscribedOn = '';
        this.properties.referenceId = '';
        this._choiceColumns = [];
        this._personColumns = [];
        this._dateColumns = [];
        await this._fetchColumns(this.properties.siteUrl, newValue);
      }
      this.context.propertyPane.refresh();
    };

    handleFieldChange().catch(err => console.error(err));
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
  }

  private async _fetchLists(siteUrl: string): Promise<void> {
    try {
      const siteSp = spfi(siteUrl).using(SPFx(this.context));
      const lists = await siteSp.web.lists.select("Id", "Title").filter("Hidden eq false")();
      this._lists = lists.map(l => ({ key: l.Id, text: l.Title }));
      this.context.propertyPane.refresh();
    } catch (error) {
      console.error("Error fetching lists", error);
    }
  }

  private async _fetchColumns(siteUrl: string, listId: string): Promise<void> {
    try {
      const siteSp = spfi(siteUrl).using(SPFx(this.context));
      const fields = await siteSp.web.lists.getById(listId).fields.select("InternalName", "Title", "TypeAsString", "FieldTypeKind")();

      this._choiceColumns = fields.filter(f => f.TypeAsString === "Choice" || f.TypeAsString === "MultiChoice").map(f => ({ key: f.InternalName, text: f.Title }));
      this._personColumns = fields.filter(f => f.TypeAsString === "User").map(f => ({ key: f.InternalName, text: f.Title }));
      this._dateColumns = fields.filter(f => f.TypeAsString === "DateTime").map(f => ({ key: f.InternalName, text: f.Title }));

      this._allColumns = [
        { key: '', text: '(None - Use Person Column)' },
        ...fields.map(f => ({ key: f.InternalName, text: f.Title }))
      ];

      this.context.propertyPane.refresh();
    } catch (error) {
      console.error("Error fetching columns", error);
    }
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (this.context.sdks.microsoftTeams) {
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams':
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "About",
              groupFields: [
                PropertyPaneLabel('version', {
                  text: 'Subscription App Version : 1.0.0.1'
                })
              ]
            },
            {
              groupName: "Data Source",
              groupFields: [
                PropertyPaneDropdown('siteUrl', {
                  label: strings.SiteFieldLabel,
                  options: this._sites
                }),
                PropertyPaneDropdown('listId', {
                  label: strings.ListFieldLabel,
                  options: this._lists,
                  disabled: !this.properties.siteUrl
                }),
                PropertyPaneDropdown('subscribeFor', {
                  label: strings.SubscribeForFieldLabel,
                  options: this._choiceColumns,
                  disabled: !this.properties.listId
                }),
                PropertyPaneDropdown('subscribedBy', {
                  label: strings.SubscribedByFieldLabel,
                  options: this._personColumns,
                  disabled: !this.properties.listId
                }),
                PropertyPaneDropdown('subscribedOn', {
                  label: strings.SubscribedOnFieldLabel,
                  options: this._dateColumns,
                  disabled: !this.properties.listId
                }),
                PropertyPaneDropdown('referenceId', {
                  label: "Unique ID Column (Optional)",
                  options: this._allColumns,
                  disabled: !this.properties.listId
                })
              ]
            },
            {
              groupName: "Content & Messages",
              groupFields: [
                PropertyPaneToggle('showTitle', {
                  label: "Show Title"
                }),
                PropertyPaneTextField('customTitle', {
                  label: "WebPart Title"
                }),
                PropertyPaneTextField('descriptionText', {
                  label: "Description Text (use <userinfo> for user email)",
                  multiline: true
                }),
                PropertyPaneTextField('successNotification', {
                  label: "Success Notification (use <userinfo> for user email)",
                  multiline: true
                }),
                PropertyPaneToggle('showSuccessNotification', {
                  label: "Show Success Notification"
                }),
                PropertyPaneTextField('buttonText', {
                  label: "Button Label"
                }),
                PropertyPaneTextField('choicesLabel', {
                  label: "Choices Section Label",
                  value: this.properties.choicesLabel
                })
              ]
            },
            {
              groupName: "Visual Styling",
              groupFields: [
                PropertyPaneSlider('titleFontSize', {
                  label: "Title Font Size",
                  min: 10,
                  max: 40,
                  step: 1
                }),
                PropertyPaneSlider('titleBarOpacity', {
                  label: "Title Bar Opacity",
                  min: 0,
                  max: 100,
                  step: 1
                }),
                PropertyPaneSlider('descriptionFontSize', {
                  label: "Description Font Size",
                  min: 10,
                  max: 40,
                  step: 1
                }),
                PropertyPaneSlider('choicesFontSize', {
                  label: "Choices Label Font Size",
                  min: 10,
                  max: 40,
                  step: 1
                }),
                PropertyPaneSlider('notificationFontSize', {
                  label: "Notification Font Size",
                  min: 10,
                  max: 40,
                  step: 1
                }),
                PropertyPaneSlider('buttonFontSize', {
                  label: "Button Font Size",
                  min: 10,
                  max: 40,
                  step: 1
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
