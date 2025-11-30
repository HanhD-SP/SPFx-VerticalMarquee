import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'VerticalMarqueeWebPartStrings';
import VerticalMarquee from './components/VerticalMarquee';
import { IVerticalMarqueeProps } from './components/IVerticalMarqueeProps';

export interface IVerticalMarqueeWebPartProps {
  description: string;
  selectedList?: string;
  scrollSpeed?: string | number;
  textColor?: string;
}

export default class VerticalMarqueeWebPart extends BaseClientSideWebPart<IVerticalMarqueeWebPartProps> {
  private _lists: IPropertyPaneDropdownOption[] = [];
  private _listsDropdownDisabled: boolean = false;
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public async render(): Promise<void> {
    const element: React.ReactElement<IVerticalMarqueeProps> = React.createElement(
      VerticalMarquee,
      {
        description: this.properties.description || '',
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        selectedList: this.properties.selectedList,
        scrollSpeed: this.properties.scrollSpeed ? (typeof this.properties.scrollSpeed === 'number' ? this.properties.scrollSpeed : parseFloat(this.properties.scrollSpeed.toString()) || 1) : 1,
        textColor: this.properties.textColor || '#000000'
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    }).then(() => {
      return this._loadLists();
    });
  }

  private async _loadLists(): Promise<void> {
    try {
      const apiUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists?$filter=Hidden eq false&$select=Id,Title&$orderby=Title`;
      const response: SPHttpClientResponse = await this.context.spHttpClient.get(
        apiUrl,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        this._lists = data.value.map((list: any) => ({
          key: list.Title,
          text: list.Title
        }));
      }
    } catch (error) {
      console.error('Error loading lists:', error);
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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneDropdown('selectedList', {
                  label: strings.SelectedListFieldLabel,
                  options: this._lists,
                  disabled: this._listsDropdownDisabled
                }),
                PropertyPaneTextField('scrollSpeed', {
                  label: strings.ScrollSpeedFieldLabel,
                  description: strings.ScrollSpeedFieldDescription,
                  value: this.properties.scrollSpeed?.toString() || '1'
                }),
                PropertyPaneTextField('textColor', {
                  label: strings.TextColorFieldLabel,
                  description: strings.TextColorFieldDescription
                })
              ]
            }
          ]
        }
      ]
    };
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    if (propertyPath === 'selectedList' && newValue) {
      this.context.propertyPane.refresh();
      this.render();
    }
  }
}
