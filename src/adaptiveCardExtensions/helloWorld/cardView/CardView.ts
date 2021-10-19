import {
  BaseBasicCardView,
  IBasicCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'HelloWorldAdaptiveCardExtensionStrings';
import { IHelloWorldAdaptiveCardExtensionProps, IHelloWorldAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../HelloWorldAdaptiveCardExtension';

export class CardView extends BaseBasicCardView<IHelloWorldAdaptiveCardExtensionProps, IHelloWorldAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return [
      {
        title: 'View Messages',
        action: {
          type: 'ExternalLink',
          parameters: {
            target: "http://outlook.office.com"
          }
        }
      }
    ];
  }

  public get data(): IBasicCardParameters {
    return {
      primaryText: `You have ${this.state.unreadCount} unread messages`,
      unreadCount: this.state.unreadCount
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'ExternalLink',
      parameters: {
        target: 'https://outlook.office.com'
      }
    };
  }
}
