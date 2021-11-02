import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import * as strings from 'HelloWorldAdaptiveCardExtensionStrings';
import { IHelloWorldAdaptiveCardExtensionProps } from './HelloWorldAdaptiveCardExtension';

export class HelloWorldPropertyPane {
  public getPropertyPaneConfiguration(properties: IHelloWorldAdaptiveCardExtensionProps): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: 'Card Settings'
          },
          groups: [
            {
              groupName: "Card Settings",
              groupFields: [
                PropertyPaneDropdown('filterBySenderEmail', {
                  label: 'Filter by Sender',
                  options: properties.senderList
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
