import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider
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
  iscertified: boolean;
  rating: number;

}

export default class WebpartPropertiesPaneWebPart extends BaseClientSideWebPart<IWebpartPropertiesPaneWebPartProps> {


  protected onInit(): Promise<void> {

    return new Promise<void>((resolve, _reject) => {
      this.properties.productname = "key board";
      this.properties.productdescription = "Wireless Kay board";
      this.properties.productcost = 50;
      this.properties.quantity = 10;
      this.properties.iscertified = false;
      this.properties.rating = 1;
      resolve(undefined);
    });
  }

  protected get disableReactivePropertyChanges(): boolean {
    return true;
  }

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
                  <td>Product Name(text) :</td>
                  <td>${this.properties.productname}</td>
                </tr>
                 <tr>
                  <td>Product Description(text) :</td>
                  <td>${this.properties.productdescription}</td>
                </tr>
                <tr>
                  <td>Product Cost(text) :</td>
                  <td>${this.properties.productcost}</td>
                </tr>
                <tr>
                  <td>Product Quantity(text) :</td>
                  <td>${this.properties.quantity}</td>
                </tr>
                <tr>
                  <td>Bill Amount(text) :</td>
                  <td>${this.properties.billamount = this.properties.productcost * this.properties.quantity}</td>
                </tr>
                <tr>
                  <td>Discount(text) :</td>
                  <td>${this.properties.discount = this.properties.billamount * 10 / 100}</td>
                </tr>
                <tr>
                  <td>Net Bill Amount(text) :</td>
                  <td>${this.properties.netbillamount = this.properties.billamount - this.properties.discount}</td>
                </tr>
                <tr>
                  <td>Is Certified?(toggle) :</td>
                  <td>${this.properties.iscertified}</td>
                </tr>
                <tr>
                  <td>Rating(slider) :</td>
                  <td>${this.properties.rating}</td>
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
                }),

                PropertyPaneToggle('iscertified', {
                  key: "iscertified",
                  label: "Is Centified?",
                  onText: "ISI Certified !",
                  offText: "Not Certified"
                }),
                PropertyPaneSlider('rating', {
                  label: "Rate the Product",
                  min: 1,
                  max: 10,
                  step: 1,
                  value: 1,
                  showValue: true
                })

              ]
            }
          ]
        }
      ]
    }
  }


}
