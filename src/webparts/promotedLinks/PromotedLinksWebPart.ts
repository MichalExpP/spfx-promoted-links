import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  IPropertyPaneDropdownOption
} from '@microsoft/sp-webpart-base';

import * as strings from 'PromotedLinksWebPartStrings';
import PromotedLinks from './components/PromotedLinks';
import { IPromotedLinksProps } from './components/IPromotedLinksProps';
import { IPromotedLinksWebPartProps, ISPLists, ISPList } from './IPromotedLinksWebPartProps';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Environment, EnvironmentType } from '@microsoft/sp-core-library';

export default class PromotedLinksWebPart extends BaseClientSideWebPart<IPromotedLinksWebPartProps> {

  public onInit<T>(): Promise<T> {
    this.fetchOptions()
    .then((data) => {
      this._listsInThisSite = data;
    });

    return Promise.resolve();
  }

  private _listsInThisSite: IPropertyPaneDropdownOption[] = [];

  public render(): void {
    const element: React.ReactElement<IPromotedLinksProps > = React.createElement(
      PromotedLinks,
      {
        isWorkbench: Environment.type == EnvironmentType.Local,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        listId: this.properties.listId,
        spHttpClient: this.context.spHttpClient,
        title: this.properties.title,
        tileSize: this.properties.tileSize,
        description: this.properties.description,
        displayMode: this.displayMode,
        updateProperty: (value: string) => {
          this.properties.title = value;
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
               PropertyPaneDropdown('listId', {
                  label: strings.selectedListNameFieldLabel,
                  options: this._listsInThisSite
                }),
                PropertyPaneDropdown('tileSize', {
                  label: "Select Tile Size",
                  options: [
                    { key: '50x100', text: '50x100' },
                    { key: '75x75', text: '75x75' },
                    { key: '75x150', text: '75x150' },
                    { key: '100x100', text: '100x100' },
                    { key: '100x200', text: '100x200' },
                    { key: '113x113', text: '113x113' },
                    { key: '125x125', text: '125x125' },
                    { key: '131x131', text: '131x131'},
                    { key: '142x142', text: '142x142'},
                    { key: '150x150', text: '150x150'},
                    { key: '181x181', text: '181x181'},
                    { key: '200x200', text: '200x200'},
                    { key: '293x160', text: '293x160'},
                  ],
                  selectedKey: '125x125',
                })
              ]
            }
          ]
        }
      ]
    };
  }

  private fetchLists(url: string) : Promise<ISPLists> {
    return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1).then((response: SPHttpClientResponse) => {
        if (response.ok) {
          return response.json();
        } else {
          console.log("WARNING - failed to hit URL " + url + ". Error = " + response.statusText);
          return null;
        }
      });
  }

  private fetchOptions(): Promise<IPropertyPaneDropdownOption[]> {
    var url = this.context.pageContext.web.absoluteUrl + `/_api/web/lists?$filter=BaseTemplate eq 170 and Hidden eq false`;

    return this.fetchLists(url).then((response) => {
        var options: Array<IPropertyPaneDropdownOption> = new Array<IPropertyPaneDropdownOption>();
        var lists: ISPList[] = response.value;
        lists.forEach((list: ISPList) => {
            //console.log("Found list with title = " + list.Title);
            options.push( { key: list.Id, text: list.Title });
        });

        return options;
    });
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
}
