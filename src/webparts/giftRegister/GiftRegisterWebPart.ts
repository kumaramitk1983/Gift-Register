import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version,UrlQueryParameterCollection } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';

import * as strings from 'GiftRegisterWebPartStrings';
import GiftRegister from './components/GiftRegister';
import { IGiftRegisterProps } from './components/IGiftRegisterProps';

export interface IGiftRegisterWebPartProps {
  description: string;
}
import "core-js/modules/es6.promise";
import "core-js/modules/es6.array.iterator.js";
import "core-js/modules/es6.array.from.js";
import "whatwg-fetch";
import "es6-map/implement";

export default class GiftRegisterWebPart extends BaseClientSideWebPart<IGiftRegisterWebPartProps> {

  public render(): void {
    var queryParameters = new UrlQueryParameterCollection(window.location.href);
    const element: React.ReactElement<IGiftRegisterProps > = React.createElement(
      GiftRegister,
      {
        description: this.properties.description,
        context : this.context,
        itemID : queryParameters.getValue('item'),
        viewMode : queryParameters.getValue('view'),
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
