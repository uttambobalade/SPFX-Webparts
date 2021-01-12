import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SiteCreationwpWebpartWebPart.module.scss';
import * as strings from 'SiteCreationwpWebpartWebPartStrings';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';

export interface ISiteCreationwpWebpartWebPartProps {
  description: string;
}

export default class SiteCreationwpWebpartWebPart extends BaseClientSideWebPart<ISiteCreationwpWebpartWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div style="margin:20px !important;" class="${styles.siteCreationwpWebpart}">

      <h1>Create a New Subsite</h1>
    <p>Please fill the below details to create a new subsite.</p><br/>

    Sub Site Title: <br/><input type='text' id='txtSubSiteTitle' /><br/>

    Sub Site URL: <br/><input type='text' id='txtSubSiteUrl' /><br/>    

    Sub Site Description: <br/><textarea id='txtSubSiteDescription' rows="5" cols="30"></textarea><br/>              
    <br/>

    <input type="button" id="btnCreateSubSite" value="Create Sub Site"/><br/>

          </div>`;

    this.bindEvents();
  }

  private bindEvents(): void {
    this.domElement.querySelector('#btnCreateSubSite').addEventListener('click', () => { this.createSubSite(); });
  }

  private createSubSite(): void {

    let subSiteTitle = document.getElementById("txtSubSiteTitle")["value"];
    let subSiteUrl = document.getElementById("txtSubSiteUrl")["value"];
    let subSiteDescription = document.getElementById("txtSubSiteDescription")["value"];

    const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/webinfos/add";

    const spHttpClientOptions: ISPHttpClientOptions = {
      body: `{
              "parameters":{
                "@odata.type": "SP.WebInfoCreationInformation",
                "Title": "${subSiteTitle}",
                "Url": "${subSiteUrl}",
                "Description": "${subSiteDescription}",
                "Language": 1033,
                "WebTemplate": "STS#0",
                "UseUniquePermissions": true
                  }
                }`
    };

    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, spHttpClientOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          alert("New Subsite has been created successfully");
        } else {
          alert("Error Message : " + response.status + " - " + response.statusText);
        }
      });


  }


  private createSubsiteold(): void {
    let siteTitle: string = document.getElementById('txtSubsiteTitle')['value'];
    let siteURL: string = document.getElementById('txtSubsiteURL')['value'];
    let siteDesc: string = document.getElementById('txtSubsiteDesc')['value'];

    const url: string = this.context.pageContext.web.absoluteUrl + "/_api/web/webinfos/add";
    const sphttpOptions: ISPHttpClientOptions = {
      body: `{
            "parameters":{
              "@odata.type":"SP.WebInfoCreationInformation",
              "Title":"${siteTitle}",
              "Url":"${siteURL}",
              "Descritpion":"${siteDesc}",
              "Language":1033,
              "WebTemplate":"STS#0",
              "UseUniquePermissions": true
            }
          }`
    };
    console.log("URL: " + url);
    console.log("body : " + JSON.stringify(sphttpOptions));

    this.context.spHttpClient.post(url, SPHttpClient.configurations.v1, sphttpOptions)
      .then((response: SPHttpClientResponse) => {
        if (response.status === 200) {
          alert("New site has been created sucessfully.");
        } else {
          alert("Error Message :" + response.status + " - " + response.statusText);
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
