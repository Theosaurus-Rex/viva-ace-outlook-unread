import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { HelloWorldPropertyPane } from './HelloWorldPropertyPane';
import { GraphError } from "@microsoft/microsoft-graph-client/lib/src";
export interface IHelloWorldAdaptiveCardExtensionProps {
  title: string;
  description: string;
  iconProperty: string;
  unreadCount: number;
  filterBySenderEmail: string;
  filterBySubject: string;
  appliedFilters: string;
}

export interface IHelloWorldAdaptiveCardExtensionState {
  description: string;
  unreadCount: number | null;
  emails: {webLink: string, subject: string, sender: string, senderEmail: string}[];
  filterBySenderEmail: string;
  filterBySubject: string;
  appliedFilters: string;
}

const CARD_VIEW_REGISTRY_ID: string = 'HelloWorld_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'HelloWorld_QUICK_VIEW';

export default class HelloWorldAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IHelloWorldAdaptiveCardExtensionProps,
  IHelloWorldAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: HelloWorldPropertyPane;

  public async onPropertyPaneFieldChanged() {
    this.setState({appliedFilters: 'isRead ne true&$count=true&$top=999'});
    this.applySenderFilter();
    this.applySubjectFilter();
    await this.getUnreadCount();
    await this.getEmailDetails();

    return Promise.resolve();
    
  }

  public async onInit(): Promise<void> {
    this.state = {
      description: this.properties.description,
      unreadCount: this.properties.unreadCount || 0,
      emails: [],
      filterBySenderEmail: this.properties.filterBySenderEmail,
      filterBySubject: this.properties.filterBySubject,
      appliedFilters: this.properties.appliedFilters
    };
    this.setState({appliedFilters: 'isRead ne true&$count=true&$top=999'});
    this.applySenderFilter();
    this.applySubjectFilter();
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
        .api('/me/mailfolders/inbox/messages')
        .version('v1.0')
        .filter(this.state.appliedFilters)
        .get((error: GraphError, response: any, rawResponse?: any): void => {
            this.setState({unreadCount: response.value.length});
          });
    } catch (error) {
      console.log(error);
    } 
  }

  private applySenderFilter() {
    
    if (this.properties.filterBySenderEmail){
      let prevFilterString = this.state.appliedFilters;
      this.setState({appliedFilters: `(from/emailAddress/address) eq '${this.properties.filterBySenderEmail}'` + ' AND ' + prevFilterString });
    } 
    console.log('applySenderFilter', this.state.appliedFilters);
  }


  private applySubjectFilter() {
    if (this.properties.filterBySubject){
      let prevFilterString = this.state.appliedFilters;
      this.setState({appliedFilters: `contains(subject, '${this.properties.filterBySubject}')` + ' AND ' + prevFilterString });
    }
    console.log('applySubjectFilter', this.state.appliedFilters);
  }

  private async getEmailDetails() {
    const graphClient = await this.context.msGraphClientFactory.getClient();
    try {
      await graphClient
        .api('/me/messages')
        .version('v1.0')
        .filter(this.state.appliedFilters)
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
