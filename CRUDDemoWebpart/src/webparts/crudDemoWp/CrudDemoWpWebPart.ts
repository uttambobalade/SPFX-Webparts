import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './CrudDemoWpWebPart.module.scss';
import * as strings from 'CrudDemoWpWebPartStrings';
import { SPHttpClient, SPHttpClientResponse, ISPHttpClientOptions } from '@microsoft/sp-http';
import { IPorductListItem } from './IProductListItem';
import * as pnp from 'sp-pnp-js';

export interface ICrudDemoWpWebPartProps {
  description: string;
}

export default class CrudDemoWpWebPart extends BaseClientSideWebPart<ICrudDemoWpWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `<div> 
        <div>
        <table border='5' bgcolor='aqua'>

        <tr>
        <td>Please Enter Product ID </td>
        <td><input type='text' id='txtID' />
        <td><input type='submit' id='btnRead' value='Read All Products' />
        </td>
        </tr>

        <br/>
          <tr>
          <td>Product Title</td>
          <td><input type='text' id='txtProductTitle' />
          </tr>

          <tr>
          <td>Product Name</td>
          <td><input type='text' id='txtProductName' />
          </tr>

          <tr>
          <td>Product Vendor</td>
          <td>
          <select id="ddlProductVendor">
            <option value="Microsoft">Microsoft</option>
            <option value="Sun">Sun</option>
            <option value="Oracle">Oracle</option>
            <option value="Google">Google</option>
          </select>  
          </td>
        
          </tr>

          <tr>
          <td>Product Version</td>
          <td><input type='text' id='txtProductVersion' />
          </tr>

          <tr>
          <td>Product Description</td>
          <td><textarea rows='5' cols='40' id='txtProductDescription'> </textarea> </td>
          </tr>

          <tr>
          <td colspan='2' align='center'>
          <input type='submit'  value='Insert Item' id='btnSubmit' />
          <input type='submit'  value='Update' id='btnUpdate' />
          <input type='submit'  value='Delete' id='btnDelete' />      
          </td>
        </table>
        </div>
        <div id="divStatus"/>
        <br/>
        <h2>All List Items Are :</h2><br/>
        <div id="DivItems"/>

     </div>`;

    this._bindEvents();
    this._readAllItems();
  }

  private _readAllItems(): void {

    var html: string = `<table border=1 width=100% style="bordercollapse: collapse;"><tr><th>Title</th> <th>Name</th> <th>Vendor</th> <th>Version</th> <th>Description</th></tr>`;
    pnp.sp.web.lists.getByTitle('ProductCatalog').items.get().then((items: any[]) => {
      items.forEach(function (item) {
        html += `<tr>
        <td>${item["Title"]}</td>
        <td>${item["Productname"]}</td>
         <td>${item["Productvendor"]}</td>
        <td>${item["Productversion"]}</td>
        <td>${item["ProductDescription"]}</td>
        </tr>`;
      });
      html += "</table>";
      let allitems: Element = this.domElement.querySelector('#DivItems');
      allitems.innerHTML = html;
    }).catch(r => {
      let statusmessage: Element = this.domElement.querySelector('#divStatus');
      statusmessage.innerHTML = "An error has occured i.e.  " + r.status + " - " + r.statusText;
    });
  }

  private _bindEvents(): void {
    this.domElement.querySelector('#btnSubmit').addEventListener('click', () => { this.addListItem(); });
    this.domElement.querySelector('#btnRead').addEventListener('click', () => { this.readListItem(); });
    this.domElement.querySelector('#btnUpdate').addEventListener('click', () => { this.updateListItem(); });
    this.domElement.querySelector('#btnDelete').addEventListener('click', () => { this.deleteListItem(); });
  }

  private deleteListItem(): void {
    const id = document.getElementById("txtID")["value"];
    pnp.sp.web.lists.getByTitle("ProductCatalog").items.getById(id).delete();
    let statusmessage: Element = this.domElement.querySelector('#divStatus');
    statusmessage.innerHTML = "Product Deleted successfully !!";
  }

  private updateListItem(): void {
    const id = document.getElementById("txtID")["value"];

    var producttitle = document.getElementById("txtProductTitle")["value"];
    var productname = document.getElementById("txtProductName")["value"];
    var productversion = document.getElementById("txtProductVersion")["value"];
    var productvendor = document.getElementById("ddlProductVendor")["value"];
    var productDescription = document.getElementById("txtProductDescription")["value"];

    pnp.sp.web.lists.getByTitle("ProductCatalog").items.getById(id).update({
      Title: producttitle,
      Productvendor: productvendor,
      Productname: productname,
      Productversion: productversion,
      ProductDescription: productDescription
    }).then(r => {
      let statusmessage: Element = this.domElement.querySelector('#divStatus');
      statusmessage.innerHTML = "Product updated successfully !!";
    })
  }

  private readListItem(): void {

    const id = document.getElementById("txtID")["value"];
    pnp.sp.web.lists.getByTitle('ProductCatalog').items.getById(id).get().then((listItem: any) => {

      document.getElementById("txtProductTitle")["value"] = listItem["Title"];
      document.getElementById("ddlProductVendor")["value"] = listItem["Productvendor"];
      document.getElementById("txtProductDescription")["value"] = listItem["ProductDescription"];
      document.getElementById("txtProductName")["value"] = listItem["Productname"];
      document.getElementById("txtProductVersion")["value"] = listItem["Productversion"];

      console.log(listItem["Title"] + " - " + listItem["ProductName"] + " - " + listItem["ProductVendor"] + " - " + listItem["ProductDescription"] + " - " + listItem["ProductVersion"]);

    });
  }

  private addListItem(): void {

    var producttitle = document.getElementById("txtProductTitle")["value"];
    var productname = document.getElementById("txtProductName")["value"];
    var productversion = document.getElementById("txtProductVersion")["value"];
    var productvendor = document.getElementById("ddlProductVendor")["value"];
    var productDescription = document.getElementById("txtProductDescription")["value"];

    pnp.sp.web.lists.getByTitle('ProductCatalog').items.add({
      Title: producttitle,
      Productvendor: productvendor,
      Productname: productname,
      Productversion: productversion,
      ProductDescription: productDescription
    }).then(r => {
      let statusmessage: Element = this.domElement.querySelector('#divStatus');
      statusmessage.innerHTML = "List Item has been created successfully.";
      this.Reset();
      this._readAllItems();
    }).catch(r => {
      let statusmessage: Element = this.domElement.querySelector('#divStatus');
      statusmessage.innerHTML = "An error has occured i.e.  " + r.status + " - " + r.statusText;
    });

  }

  private Reset(): void {
    document.getElementById("txtProductTitle")["value"] = '';
    document.getElementById("ddlProductVendor")["value"] = 'Microsoft';
    document.getElementById("txtProductDescription")["value"] = '';
    document.getElementById("txtProductVersion")["value"] = '';
    document.getElementById("txtProductName")["value"] = '';
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
