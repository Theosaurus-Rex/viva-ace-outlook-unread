import { IPropertyPaneConfiguration, PropertyPaneTextField, PropertyPaneDropdown } from '@microsoft/sp-property-pane';
import * as strings from 'HelloWorldAdaptiveCardExtensionStrings';
import { IHelloWorldAdaptiveCardExtensionProps } from './HelloWorldAdaptiveCardExtension';

export class HelloWorldPropertyPane {
  public getPropertyPaneConfiguration(properties: IHelloWorldAdaptiveCardExtensionProps): IPropertyPaneConfiguration {
    let senderListOptions = []
    properties.senderList.forEach(sender =>{
      senderListOptions.push(
        {key: sender.senderEmailAddress, text: sender.senderName}
      )
    })
    console.log("SENDER OPTIONS PANE", senderListOptions)
    let filteredSenderListEmails = []
   let filteredSenderListOptions = []
   senderListOptions.forEach(sender => {
    if (!filteredSenderListEmails.includes(sender.key)){
      filteredSenderListEmails.push(sender.key);
      filteredSenderListOptions.push(sender)
    }
   })
   console.log("FILTERED", filteredSenderListEmails)
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
                PropertyPaneDropdown('senderEmail', {
                  label: 'Filter by Sender',
                  options: filteredSenderListOptions
                  
                }),
              ]
            }
          ]
        }
      ]
    };
  }
}
