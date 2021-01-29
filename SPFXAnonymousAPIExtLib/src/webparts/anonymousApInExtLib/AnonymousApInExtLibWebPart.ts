import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'AnonymousApInExtLibWebPartStrings';
import AnonymousApInExtLib from './components/AnonymousApInExtLib';
import { IAnonymousApInExtLibProps } from './components/IAnonymousApInExtLibProps';

export interface IAnonymousApInExtLibWebPartProps {
  description: string;
  siteURL: string;
  userId: string;
}

export default class AnonymousApInExtLibWebPart extends BaseClientSideWebPart<IAnonymousApInExtLibWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IAnonymousApInExtLibProps> = React.createElement(
      AnonymousApInExtLib,
      {
        description: this.properties.description,
        siteURL: this.properties.siteURL,
        userId: this.properties.userId,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('siteURL', {
                  label: "Enter API URL"
                }),
                PropertyPaneTextField('userId', {
                  label: "Enter User ID"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
