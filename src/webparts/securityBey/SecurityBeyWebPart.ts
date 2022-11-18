import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'SecurityBeyWebPartStrings';
import SecurityBey from './components/SecurityBey';
import { ISecurityBeyProps } from './components/ISecurityBeyProps';
import { sp } from '@pnp/sp';

export interface ISecurityBeyWebPartProps {
  description: string;
}

export default class SecurityBeyWebPart extends BaseClientSideWebPart<ISecurityBeyWebPartProps> {
  public onInit() {
    sp.setup({ spfxContext: this.context });
    return super.onInit();
  }
  public render(): void {
    const element: React.ReactElement<ISecurityBeyProps> = React.createElement(
      SecurityBey,
      {
        //description: this.properties.description
        url: this.context.pageContext.web.absoluteUrl,
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
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
