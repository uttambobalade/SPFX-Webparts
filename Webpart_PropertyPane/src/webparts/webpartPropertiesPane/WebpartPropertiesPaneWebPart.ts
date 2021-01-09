import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneSlider,
  PropertyPaneChoiceGroup,
  PropertyPaneDropdown,
  PropertyPaneCheckbox,
  PropertyPaneLink
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
  processor: string;
  invoicetype: string;
  color: string;
  windowslicense: boolean;
  ecommercesite: URL;
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
      this.properties.processor = "I7";
      this.properties.invoicetype = "Paper";
      this.properties.color = "black";
      this.properties.windowslicense = false;

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
                  <td>Invoice Type(choicegroupwithimage) :</td>
                  <td>${this.properties.invoicetype}</td>
                </tr>
                <tr>
                  <td>Is Certified?(toggle) :</td>
                  <td>${this.properties.iscertified}</td>
                </tr>
                <tr>
                  <td>Rating(slider) :</td>
                  <td>${this.properties.rating}</td>
                </tr>

                <tr>
                  <td>Processor Type(choicegroup) :</td>
                  <td>${this.properties.processor}</td>
                </tr>
                <tr>
                  <td>Color(dropdown) :</td>
                  <td>${this.properties.color}</td>
                </tr>
                <tr>
                  <td>Required Windows License?(checkbox) :</td>
                  <td>${this.properties.windowslicense}</td>
                </tr>
                <tr>
                  <td>Buy from Best E-Commerce Site(link) :</td>
                  <td>${this.properties.ecommercesite}</td>
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
                  label: 'Product Name',
                  multiline: false,
                  resizable: false,
                  placeholder: 'Please enter Product Name', 'description': 'Text field'
                }),
                PropertyPaneTextField('productdescription', {
                  label: 'Product Description',
                  multiline: true,
                  resizable: false,
                  placeholder: 'Please enter Product Description', 'description': 'Text field'
                })
              ]
            },
            {
              groupName: 'Billing Info',
              groupFields: [
                PropertyPaneTextField('productcost', {
                  label: 'Product Cost',
                  multiline: false,
                  resizable: false,
                  placeholder: 'Please enter Product Cost', 'description': 'Number field'
                }),

                PropertyPaneTextField('quantity', {
                  label: 'Quantity',
                  multiline: false,
                  resizable: false,
                  placeholder: 'Please enter Quantity', 'description': 'Number field'
                }),

                PropertyPaneChoiceGroup('invoicetype', {
                  label: 'Invoice Type',
                  options: [
                    {
                      key: 'Paper Print', text: 'Paper Print',
                      imageSrc: 'file:///C:/Users/uttam/Pictures/Sample/MS%20office%20icons/print.jpg',
                      imageSize: { width: 32, height: 32 },
                      selectedImageSrc: 'file:///C:/Users/uttam/Pictures/Sample/MS%20office%20icons/print.jpg'
                    },
                    {
                      key: 'Email', text: 'Email',
                      imageSrc: 'file:///C:/Users/uttam/Pictures/Sample/MS%20office%20icons/email.jpg',
                      imageSize: { width: 32, height: 32 },
                      selectedImageSrc: 'file:///C:/Users/uttam/Pictures/Sample/MS%20office%20icons/email.jpg'
                    },
                    {
                      key: 'PDF', text: 'PDF',
                      imageSrc: 'file:///C:/Users/uttam/Pictures/Sample/MS%20office%20icons/pdf.jpg',
                      imageSize: { width: 32, height: 32 },
                      selectedImageSrc: 'file:///C:/Users/uttam/Pictures/Sample/MS%20office%20icons/pdf.jpg'
                    },
                    {
                      key: 'Word', text: 'Word',
                      imageSrc: 'file:///C:/Users/uttam/Pictures/Sample/MS%20office%20icons/word.jpg',
                      imageSize: { width: 32, height: 32 },
                      selectedImageSrc: 'file:///C:/Users/uttam/Pictures/Sample/MS%20office%20icons/word.jpg'

                    }
                  ]
                })

              ]
            },
            {
              groupName: 'Product Configuration',
              groupFields: [
                PropertyPaneToggle('iscertified', {
                  key: 'iscertified',
                  label: 'Is Centified?',
                  onText: 'ISI Certified !',
                  offText: 'Not Certified'
                }),
                PropertyPaneSlider('rating', {
                  label: 'Rate the Product',
                  min: 1,
                  max: 10,
                  step: 1,
                  value: 1,
                  showValue: true
                }),

                PropertyPaneChoiceGroup('processor', {
                  label: 'Processor Type',
                  options: [
                    { key: 'I5', text: 'Intel I5' },
                    { key: 'I7', text: 'Intel I7', checked: true },
                    { key: 'I9', text: 'Intel I9' },
                    { key: 'AMD', text: 'AMD 5000' }
                  ]
                }),

                PropertyPaneDropdown('color', {
                  label: 'Processor Type',
                  options: [
                    { key: 'black', text: 'black' },
                    { key: 'white', text: 'white' },
                    { key: 'gold', text: 'gold' }
                  ],
                  selectedKey: 'black'
                }),

                PropertyPaneCheckbox('windowslicense', {
                  text: 'Do you require Windows License?',
                  checked: false,
                  disabled: false
                }),

                PropertyPaneLink('ecommercesite', {
                  href: 'https://www.amazon.com/ref=nav_logo',
                  text: 'Buy this order from the best seller',
                  target: '_blank',
                  //optionally set the below property to show the page/link in pop up.
                  popupWindowProps: {
                    height: 500,
                    width: 500,
                    positionWindowPosition: 2,
                    title: 'Amazon'
                  }
                })
              ]
            }
          ],
          displayGroupsAsAccordion: true
        }
      ]
    }
  }


}
