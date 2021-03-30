import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'OutlookEventsWebPartStrings';
import OutlookEvents from './components/OutlookEvents';
import { IOutlookEventsProps } from './components/IOutlookEventsProps';

export interface IOutlookEventsWebPartProps {
  description: string;
}

export default class OutlookEventsWebPart extends BaseClientSideWebPart<IOutlookEventsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IOutlookEventsProps> = React.createElement(
      OutlookEvents,
      {
        description: this.properties.description,
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
