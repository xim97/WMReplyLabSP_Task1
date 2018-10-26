import * as React from 'react';
import styles from './PromotionWebPart.module.scss';
import { IPromotionWebPartProps } from './IPromotionWebPartProps';
import { escape } from '@microsoft/sp-lodash-subset';
import { IPromotionWebPartState } from './IPromotionWebPartState';
import * as strings from "PromotionWebPartWebPartStrings";

export default class PromotionWebPart extends React.Component<IPromotionWebPartProps, IPromotionWebPartState> {

  constructor(props) {
    super(props);    
    if (JSON.parse(localStorage.getItem("isExpanded")) === null) {
      localStorage.setItem("isExpanded", 
        JSON.stringify(this.props.expandCollapseDefaultValue  === "expand" ? true : false));
    }
    this.state = {
      isExpanded: JSON.parse(localStorage.getItem("isExpanded"))
    }
    this.handleClickExpandCollapseButton = this.handleClickExpandCollapseButton.bind(this);
  }

  private handleClickExpandCollapseButton(): void {
    let isExpanded: boolean = JSON.parse(localStorage.getItem("isExpanded"));
    localStorage.setItem("isExpanded", JSON.stringify(!isExpanded));
    this.setState({ isExpanded: !isExpanded });
  }

  public render(): React.ReactElement<IPromotionWebPartProps> {
    let helloMessage = this.props.helloMessage;
    helloMessage = helloMessage.replace("{username}", this.props.userName);

    return (
      <div className={styles.promotionWebPart}>
        <div style={{ backgroundColor: this.props.backgroundColor }} className={styles.container}>
          <div className={styles.row}>
            <div className={styles.column}>
              <p>{escape(helloMessage)}</p>
              {
                this.state.isExpanded && <p>{escape(this.props.promotionMessage)}</p>
              }
              <input
                onClick={this.handleClickExpandCollapseButton}
                className={styles.button}
                type="button"
                value={this.state.isExpanded ? strings.CollapseButton : strings.ExpandButton}
              />
            </div>
          </div>
        </div>
      </div>
    );
  }
}
