import * as React from "react";
import styles from "./EasyTabsWebpart.module.scss";
import { IEasyTabsWebpartProps } from "./IEasyTabsWebpartProps";
import { escape } from "@microsoft/sp-lodash-subset";
import { Pivot, PivotItem } from "office-ui-fabric-react/lib/Pivot";
import { Label } from "office-ui-fabric-react/lib/Label";

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";

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
          <PivotItem linkText="Home" itemIcon="Globe">
            <Label className={styles.label}>
              {this.state.tabCount}
              Label 1{this.getDocs()}
            </Label>
          </PivotItem>
          <PivotItem linkText="Recent">
            <Label>
              Pivot #2 {this.props.numberOfItems}
              <p className="ms-font-l">{escape(this.props.listName)}</p>
              <p className="ms-font-l">{escape(this.props.item)}</p>
            </Label>
          </PivotItem>
          <PivotItem linkText="Shared with me">
            <Label>Pivot #3 {this.props.siteUrl}</Label>
          </PivotItem>
        </Pivot>
      </div>
    );
  }
  public getDocs(): any {
    // console.log("Calling services");
    // this._getLibraries()
    //   .then(res => {
    //     // console.log(res);
    //     res.forEach(val => {
    //       console.log(val);
    //     });
    //   })
    //   .catch(err => {
    //     console.log(err);
    //   });
  }

  private _getLibraries(): Promise<any> {
    return new Promise<any>(
      (resolve: (Title: any) => void, reject: (error: any) => void): void => {
        this.props.spHttpClient
          .get(
            this.props.siteUrl +
              `/_api/Web/Lists?$filter=BaseTemplate eq 101 and Title ne 'Site Assets' and Title ne 'Style Library'`,
            SPHttpClient.configurations.v1,
            {
              headers: {
                Accept: "application/json;odata=nometadata",
                "odata-version": ""
              }
            }
          )
          .then(
            (response: SPHttpClientResponse): any => {
              return response.json();
            },
            (error: any): void => {
              reject(error);
            }
          )
          .then(
            (response: { value: { Title: string }[] }): void => {
              if (!response.value) {
                resolve(null);
              } else {
                resolve(response.value);
              }
            }
          );
      }
    );
  }
}
