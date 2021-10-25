import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { HelloWorldPropertyPane } from './HelloWorldPropertyPane';
import { MSGraphClientFactory } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { GraphError } from "@microsoft/microsoft-graph-client/lib/src";
export interface IHelloWorldAdaptiveCardExtensionProps {
  title: string;
  description: string;
  iconProperty: string;
  unreadCount: number;
}

export interface IHelloWorldAdaptiveCardExtensionState {
  description: string;
  unreadCount: number | null;
  emails: {id: string, subject: string, sender: string}[];
}

const CARD_VIEW_REGISTRY_ID: string = 'HelloWorld_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'HelloWorld_QUICK_VIEW';

export default class HelloWorldAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IHelloWorldAdaptiveCardExtensionProps,
  IHelloWorldAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: HelloWorldPropertyPane | undefined;

  public async onInit(): Promise<void> {
    this.state = {
      description: this.properties.description,
      unreadCount: 0,
      emails: []
    };

    await this.getUnreadCount();
    await this.getEmailDetails();

    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());

    return Promise.resolve();
  }

  private async getUnreadCount() {
    const graphClient = await this.context.msGraphClientFactory.getClient();
    try {
      await graphClient
        .api('/me/messages')
        .version('v1.0')
        .filter('isRead ne true&$count=true&$top=999')
        .get((error: GraphError, response: any, rawResponse?: any): void => {

            this.setState({unreadCount: response.value.length});
            console.log("getUnreadCount RESPONSE", response);
          });
      
    } catch (error) {
      console.log(error);
    } 
  }

  private async getEmailDetails() {
    const graphClient = await this.context.msGraphClientFactory.getClient();
    try {
      await graphClient
        .api('/me/messages')
        .version('v1.0')
        .filter('isRead ne true&$count=true&$top=999')
        .get((error: GraphError, response: any, rawResponse?: any): void => {
            response.value.forEach(email => {
              this.state.emails.push(
                {
                  id: email.id,
                  subject: email.subject,
                  sender: email.sender.emailAddress.name
                }
              );
            });
          });
          console.log(this.state.emails);
      
    } catch (error) {
      console.log(error);
    } 
  }


  public get title(): string {
    return this.properties.title;
  }

  protected get iconProperty(): string {
    return this.properties.iconProperty || require('./assets/SharePointLogo.svg');
  }

  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'HelloWorld-property-pane'*/
      './HelloWorldPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.HelloWorldPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane!.getPropertyPaneConfiguration();
  }
}
