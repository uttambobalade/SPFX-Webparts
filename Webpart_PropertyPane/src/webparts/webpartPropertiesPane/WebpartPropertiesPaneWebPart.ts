import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './WebpartPropertiesPaneWebPart.module.scss';
import * as strings from 'WebpartPropertiesPaneWebPartStrings';

export interface IWebpartPropertiesPaneWebPartProps {

  productname: string;
  productdescription: string;
  productcost: number;
  quantity: number;
  billamount: number;
  discount: number;
  netbillamount: number;

}

export default class WebpartPropertiesPaneWebPart extends BaseClientSideWebPart<IWebpartPropertiesPaneWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.webpartPropertiesPane}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.title}">SPFX Webpart Proeprty Pane</span>
            
              <br/>
              <table>
                <tr>
                  <td>Product Name :</td>
                  <td>${this.properties.productname}</td>
                </tr>
                 <tr>
                  <td>Product Description :</td>
                  <td>${this.properties.productdescription}</td>
                </tr>
                <tr>
                  <td>Product Cost :</td>
                  <td>${this.properties.productcost}</td>
                </tr>
                <tr>
                  <td>Product Quantity :</td>
                  <td>${this.properties.quantity}</td>
                </tr>
                <tr>
                  <td>Bill Amount :</td>
                  <td>${this.properties.billamount = this.properties.productcost * this.properties.quantity}</td>
                </tr>
                <tr>
                  <td>Discount :</td>
                  <td>${this.properties.discount = this.properties.billamount * 10 / 100}</td>
                </tr>
                <tr>
                  <td>Net Bill Amount :</td>
                  <td>${this.properties.netbillamount = this.properties.billamount - this.properties.discount}</td>
                </tr>
              </table>
            </div>
          </div>
        </div>
      </div>`;
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: strings.PropertyPaneDescription },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('productname', {
                  label: "Product Name",
                  multiline: false,
                  resizable: false,
                  placeholder: "Please enter Product Name", "description": "Text field"
                }),
                PropertyPaneTextField('productdescription', {
                  label: "Product Description",
                  multiline: true,
                  resizable: false,
                  placeholder: "Please enter Product Description", "description": "Text field"
                }),

                PropertyPaneTextField('productcost', {
                  label: "Product Cost",
                  multiline: false,
                  resizable: false,
                  placeholder: "Please enter Product Cost", "description": "Number field"
                }),

                PropertyPaneTextField('quantity', {
                  label: "Quantity",
                  multiline: false,
                  resizable: false,
                  placeholder: "Please enter Quantity", "description": "Number field"
                })

              ]
            }
          ]
        }
      ]
    }
  }

  // protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
  //   return {
  //     pages: [
  //       {
  //         header: {
  //           description: strings.PropertyPaneDescription
  //         },
  //         groups: [
  //           {
  //             groupName: strings.BasicGroupName,
  //             groupFields: [
  //               PropertyPaneTextField('description', {
  //                 label: strings.DescriptionFieldLabel
  //               })
  //             ]
  //           }
  //         ]
  //       }
  //     ]
  //   };
  // }
}
