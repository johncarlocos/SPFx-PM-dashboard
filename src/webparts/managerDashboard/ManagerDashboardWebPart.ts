import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import ManagerDashboard from './components/ManagerDashboard';
import { IManagerDashboardProps } from './components/IManagerDashboardProps';

export interface IManagerDashboardWebPartProps {
  description: string;
}

export default class ManagerDashboardWebPart extends BaseClientSideWebPart<IManagerDashboardWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IManagerDashboardProps> = React.createElement(
      ManagerDashboard,
      {
        siteUrl: this.context.pageContext.web.absoluteUrl,
        userDisplayName: this.context.pageContext.user.displayName
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
            description: 'Manager Dashboard Settings'
          },
          groups: [
            {
              groupName: 'Basic',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Description'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
