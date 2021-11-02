import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { HelloWorldPropertyPane } from './HelloWorldPropertyPane';
import { GraphError } from "@microsoft/microsoft-graph-client/lib/src";
import * as strings from 'HelloWorldAdaptiveCardExtensionStrings';
export interface IHelloWorldAdaptiveCardExtensionProps {
  title: string;
  description: string;
  iconProperty: string;
  unreadCount: number;
  senderList: {key: string, text: string}[];
  filterBySenderEmail: string;
}

export interface IHelloWorldAdaptiveCardExtensionState {
  description: string;
  unreadCount: number | null;
  emails: {webLink: string, subject: string, sender: string, senderEmail: string}[];
  senderList: {key: string, text: string}[];
  filterBySenderEmail: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'HelloWorld_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'HelloWorld_QUICK_VIEW';

export default class HelloWorldAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IHelloWorldAdaptiveCardExtensionProps,
  IHelloWorldAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: HelloWorldPropertyPane;

  public async componentDidUpdate(){
    console.log("UPDATING");
    await this.getUnreadCount();
    await this.getEmailDetails();
    await this.getSenderList();

    return Promise.resolve();
  }

  public async onInit(): Promise<void> {
    this.state = {
      description: this.properties.description,
      unreadCount: this.properties.unreadCount || 0,
      emails: [],
      senderList: [],
      filterBySenderEmail: this.properties.filterBySenderEmail
    };
    await this.getUnreadCount();
    await this.getEmailDetails();
    await this.getSenderList();

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
        .filter(`(from/emailAddress/address) eq '${this.properties.filterBySenderEmail}'&$isRead ne true&$count=true&$top=999`)
        .get((error: GraphError, response: any, rawResponse?: any): void => {
          console.log("This.properties.filterBySenderEmail", this.properties.filterBySenderEmail)
            this.setState({unreadCount: response.value.length});
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
        .filter(`(from/emailAddress/address) eq '${this.properties.filterBySenderEmail}'&$isRead ne true&$count=true&$top=999`)
        .get((error: GraphError, response: any, rawResponse?: any): void => {
          console.log("getEmailDetails response value:", response.value)
            response.value.forEach(email => {
              let emails = [...this.state.emails];
              emails.push({
                webLink: email.webLink,
                subject: email.subject,
                sender: email.sender.emailAddress.name,
                senderEmail: email.sender.emailAddress.address
              });
              this.setState({emails});
            });
            console.log('this.state.emails after spread syntax:', this.state.emails);
          });
    } catch (error) {
      console.log(error);
    } 
  }

  public async getSenderList() {
    const graphClient = await this.context.msGraphClientFactory.getClient();
    try {
      await graphClient
        .api('/me/messages')
        .version('v1.0')
        .filter(`isRead ne true&$count=true&$top=999`)
        .get((error: GraphError, response: any, rawResponse?: any): void => {
          console.log("getEmailDetails response value:", response.value)
          response.value.forEach(email => {
            let senderList = [...this.state.senderList];
            senderList.push(
              {
                key: email.sender.emailAddress.address,
                text: email.sender.emailAddress.name 
              }
            );
            this.setState({senderList});
            this.properties.senderList = this.state.senderList;
          });
          console.log('this.state.senderList after spread:', this.state.senderList);
          console.log('this.props.senderList', this.properties.senderList);
        });    
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
    return this._deferredPropertyPane!.getPropertyPaneConfiguration(this.properties);
  }
}
