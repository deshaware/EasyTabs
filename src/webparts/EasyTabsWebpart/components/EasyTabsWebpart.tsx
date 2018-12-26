import * as React from "react";
import styles from "./EasyTabsWebpart.module.scss";
import { IEasyTabsWebpartProps } from "./IEasyTabsWebpartProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { Pivot, PivotItem } from "office-ui-fabric-react/lib/Pivot";
import { Label } from "office-ui-fabric-react/lib/Label";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

import DetailsListDocuments from "./DetailsListDocuments";

export interface IEasyTabsWebpartState {
  tabCount: number;
  error: string;
}

export default class EasyTabsWebpart extends React.Component<
  IEasyTabsWebpartProps,
  IEasyTabsWebpartState
> {
  constructor(props: IEasyTabsWebpartProps, state: IEasyTabsWebpartState) {
    super();

    this.state = { tabCount: 3, error: "" };
  }
  public componentWillReceiveProps(
    prevProps: IEasyTabsWebpartProps,
    prevState: IEasyTabsWebpartState
  ): void {
    if (this.props.numberOfItems !== prevProps.numberOfItems) {
      this.setState({ tabCount: this.props.numberOfItems + 1 });
    }
  }

  public render(): React.ReactElement<IEasyTabsWebpartProps> {
    return (
      <div className={styles.easyTabsWebpart}>
        {/* <div className={ styles.container }>
          <div className={ styles.row }>
            <div className={ styles.column }>
              <span className={ styles.title }>Welcome to SharePoint!</span>
              <p className={ styles.subTitle }>Customize SharePoint experiences using Web Parts.</p>
              <p className={ styles.description }>{escape(this.props.description)}</p>
              <a href="https://aka.ms/spfx" className={ styles.button }>
                <span className={ styles.label }>Learn more</span>
              </a>
            </div>
          </div>
        </div> */}
        <Pivot>
          <PivotItem
            // headerText="My Files"
            linkText="Tab 1"
            // headerButtonProps={{
            //   'data-order': 1,
            //   'data-title': 'My Files Title'
            // }}
            itemIcon="Globe"
          >
            <Label className={styles.label}>
              {this.state.tabCount}

              <DetailsListDocuments
                spHttpClient={this.props.spHttpClient}
                siteUrl={this.props.siteUrl}
                listTitle="doc_test"
              />
            </Label>
          </PivotItem>
          <PivotItem linkText="Recent">
            <Label>
              Pivot #2 {this.props.numberOfItems}
              <p className="ms-font-l">{escape(this.props.tabName)}</p>
              <p className="ms-font-l">{escape(this.props.item)}</p>
              <DetailsListDocuments
                spHttpClient={this.props.spHttpClient}
                siteUrl={this.props.siteUrl}
                listTitle="Documents"
              />
            </Label>
          </PivotItem>
          <PivotItem linkText="Shared with me">
            <Label>Pivot #3 {this.props.siteUrl}</Label>
          </PivotItem>
        </Pivot>
      </div>
    );
  }
}
