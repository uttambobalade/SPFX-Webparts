import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SpListDispReactDemoWebPartStrings';
import SpListDispReactDemo from './components/SpListDispReactDemo';
import { ISpListDispReactDemoProps } from './components/ISpListDispReactDemoProps';

export interface ISpListDispReactDemoWebPartProps {
  description: string;
}

export default class SpListDispReactDemoWebPart extends BaseClientSideWebPart<ISpListDispReactDemoWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpListDispReactDemoProps> = React.createElement(
      SpListDispReactDemo,
      {
        description: this.properties.description,
        siteURL: this.context.pageContext.web.absoluteUrl
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
