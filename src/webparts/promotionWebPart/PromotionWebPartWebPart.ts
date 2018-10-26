import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneDropdown
} from '@microsoft/sp-webpart-base';

import * as strings from 'PromotionWebPartWebPartStrings';
import PromotionWebPart from './components/PromotionWebPart';
import { IPromotionWebPartProps } from './components/IPromotionWebPartProps';

export interface IPromotionWebPartWebPartProps {
  helloMessage: string;
  promotionMessage: string;
  userName: string;
  backgroundColor: string;
  expandCollapseDefaultValue: string;
}

export default class PromotionWebPartWebPart extends BaseClientSideWebPart<IPromotionWebPartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IPromotionWebPartProps > = React.createElement(
      PromotionWebPart,
      {
        helloMessage: this.properties.helloMessage,
        promotionMessage: this.properties.promotionMessage,
        userName: this.context.pageContext.user.displayName,
        backgroundColor: this.properties.backgroundColor,
        expandCollapseDefaultValue: this.properties.expandCollapseDefaultValue
      }
    );

    ReactDom.render(element, this.domElement);
  }

  private validateHelloMessage(value: string): string {
    if (value === null ||
      value.trim().length === 0) {
      return "Provide a hello message with {username} token.";
    }

    if (value.indexOf("{username}") === - 1) {
      return "Hello message should contain {username} token.";
    }

    return "";
  }

  private validatePromotionMessage(value: string): string {
    if (value === null ||
      value.trim().length === 0) {
      return "Provide a promotion message.";
    }

    return "";
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
              groupName: strings.FirstGroupName,
              groupFields: [
                PropertyPaneTextField('helloMessage', {
                  label: strings.HelloFieldLabel,
                  onGetErrorMessage: this.validateHelloMessage.bind(this)
                }),
                PropertyPaneTextField('promotionMessage', {
                  label: strings.PromotionFieldLabel,
                  onGetErrorMessage: this.validatePromotionMessage.bind(this)
                }),
              ]
            },
            {
              groupName: strings.SecondGroupName,
              groupFields: [
                PropertyPaneDropdown("backgroundColor", {
                  label: strings.BackGroundFieldLabel,
                  options: [
                    { key: "black", text: strings.BlackColor },
                    { key: "blue", text: strings.BlueColor }
                  ],
                }),
                PropertyPaneDropdown("expandCollapseDefaultValue", {
                  label: strings.ExpandCollapseDefaultValueFieldLabel,
                  options: [
                    { key: "expand", text: strings.ExpandOption },
                    { key: "collapse", text: strings.CollapseOption }
                  ],
                })
              ]
            }
          ]
        }
      ]    
    };
  }
}
