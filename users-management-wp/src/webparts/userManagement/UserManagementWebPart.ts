import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'UserManagementWebPartStrings';
import UserManagement from './components/UserManagement';
import { IUserManagementProps } from './components/IUserManagementProps';
import { SPHttpClient } from '@pnp/sp/presets/all';

export interface IUserManagementWebPartProps {
  spHttpClient: SPHttpClient;
  versionName: string;
}

export default class UserManagementWebPart extends BaseClientSideWebPart<IUserManagementWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IUserManagementProps> = React.createElement(
      UserManagement,
      {
        spHttpClient: this.context.spHttpClient,
        httpClient: this.context.httpClient,
        webpartContext: this.context,
        versionName: this.properties.versionName,
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
                PropertyPaneTextField('versionName', {
                  label: strings.versionName
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
