import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './NewListCreationWpWebPart.module.scss';
import * as strings from 'NewListCreationWpWebPartStrings';

import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

export interface INewListCreationWpWebPartProps {
  description: string;
}

export default class NewListCreationWpWebPart extends BaseClientSideWebPart<INewListCreationWpWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.container}">
        <div style="margin:20px !important;">
        <h3> Creating new list dynamically.</h3><br/>

        <p>Please fill out the below details to create new list.</p> <br/>

        New List Name:<br/> <input type="text" id="newListName" /> <br/><br/>

        New List Description:<br/> <input type="text" id="newListDesc" /> <br/><br/>

        <input class="${styles.button}" type="button" id="btnCreateNewList" value="Create New List" /><br/>

        </div>
      </div>`;

    this.bindEvents();
  }

  private bindEvents(): void {
    this.domElement.querySelector('#btnCreateNewList').addEventListener('click', () => { this.createNewList(); });
  }

  private createNewList(): void {

    var newListName = document.getElementById('newListName')['value'];
    var listdesc = document.getElementById('newListDesc')['value'];

    const listURL: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists/GetByTitle('" + newListName + "')";

    this.context.spHttpClient.get(listURL, SPHttpClient.configurations.v1)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          alert("A list is already exists with same name.");
          return;
        }
        if (response.status === 404) {
          const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/lists";

          const listDefination: any = {
            "Title": newListName,
            "Description": listdesc,
            "AllowContentTypes": true,
            "BaseTemplate": 100,
            "ContentTypesEnabled": true
          };

          const spHttpClientOptions: ISPHttpClientOptions = {
            "body": JSON.stringify(listDefination)
          };

          this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
            .then((response: SPHttpClientResponse) => {
              if (response.status === 201) {
                alert("A new lsit has been created sucessfully");
                return;
              } else {
                alert("Error message : " + response.status + " - " + response.statusText);
              }
            })
        } else {
          alert("Error message : " + response.status + " - " + response.statusText);
        }
      });

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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
