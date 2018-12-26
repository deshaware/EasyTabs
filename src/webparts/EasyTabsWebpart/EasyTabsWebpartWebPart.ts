import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-webpart-base';

import * as strings from 'EasyTabsWebpartWebPartStrings';
import EasyTabsWebpart from './components/EasyTabsWebpart';
import { IEasyTabsWebpartProps } from './components/IEasyTabsWebpartProps';

import { SPHttpClient, SPHttpClientResponse } from "@microsoft/sp-http";
import { PropertyPaneAsyncDropdown } from '../../controls/PropertyPaneAsyncDropdown/PropertyPaneAsyncDropdown';
import { IDropdownOption } from 'office-ui-fabric-react/lib/components/Dropdown';
import { update, get } from '@microsoft/sp-lodash-subset';

export interface IEasyTabsWebpartWebPartProps {
  tabName: string;
  order: string;
  numberOfItems: number;
  style: string;
  item: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
}



export default class EasyTabsWebpartWebPart extends BaseClientSideWebPart<IEasyTabsWebpartWebPartProps> {
  private itemsDropDown: PropertyPaneAsyncDropdown;
  
  public render(): void {
    const element: React.ReactElement<IEasyTabsWebpartProps > = React.createElement(
      EasyTabsWebpart,
      {
        tabName: this.properties.tabName,
        order: this.properties.order,
        numberOfItems: this.properties.numberOfItems,
        style: this.properties.style,
        item: this.properties.item,
        spHttpClient: this.context.spHttpClient,
        siteUrl: this.context.pageContext.web.absoluteUrl
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    this.itemsDropDown = new PropertyPaneAsyncDropdown('item', {
      label: strings.ItemFieldLabel,
      loadOptions: this.loadItems.bind(this),
      onPropertyChange: this.onListItemChange.bind(this),
      selectedKey: this.properties.item,
    
      // should be disabled if no list has been selected
      disabled: !this.properties.tabName
    });  
      return {
        pages:[
          {
            header:{
              description: "This is description"
            },
            groups :[
              { 
                groupName:"Group Name",
                groupFields: [
                  // PropertyPaneTextField('tabName',{//props
                  //   label:'LableS',               
                  // }),
                  PropertyPaneSlider('numberOfItems',{
                    label:"Number of tabs",
                    min:3,
                    max:8,
                    step:1
                  }),
                  new PropertyPaneAsyncDropdown('tabName', {
                    label: "Configure Tabs Here",
                    loadOptions: this.loadLists.bind(this),
                    onPropertyChange: this.onListChange.bind(this),
                    selectedKey: this.properties.tabName
                  }),
                  this.itemsDropDown                
                ]
              }
            ]
          }
        ]
      };
  }
  private loadLists(): Promise<IDropdownOption[]> {    
    return new Promise<IDropdownOption[]>((resolve: (options: IDropdownOption[]) => void
    , reject: (error: any) => void) => {      
        setTimeout(() => {
          resolve([
            {
              key: "tab1",
              text: "Tab 1"
            },
            {
              key: "tab2",
              text: "Tab 2"
            }
          ]);
        }, 2000);
      }
    );
}       
     

  private onListChange(propertyPath: string, newValue: any): void {
  //  const oldValue: any = get(this.properties, propertyPath);
  //  // store new value in web part properties
  //  update(this.properties, propertyPath, (): any => { return newValue; });
  //  // refresh web part
  //  this.render();
  const oldValue: any = get(this.properties, propertyPath);
   // store new value in web part properties
   update(this.properties, propertyPath, (): any => { return newValue; });
   // reset selected item
   this.properties.item = undefined;
   // store new value in web part properties
   update(this.properties, 'item', (): any => { return this.properties.item; });
   // refresh web part
   this.render();
   // reset selected values in item dropdown
   this.itemsDropDown.properties.selectedKey = this.properties.item;
   // allow to load items
   this.itemsDropDown.properties.disabled = false;
   // load items and re-render items dropdown
   this.itemsDropDown.render();

 }

 private loadItems(): Promise<IDropdownOption[]> {
  if (!this.properties.tabName) {
    // resolve to empty options since no list has been selected
    return Promise.resolve();
  }

  const wp: EasyTabsWebpartWebPart = this;

  return new Promise<IDropdownOption[]>((resolve: (options: IDropdownOption[]) => void, reject: (error: any) => void) => {    
     let dropdownLib = [];
     this._getLibraries()
       .then(response => {
         let obj: IDropdownOption[];
         response.forEach(element => {
           let values: IDropdownOption = {
             key: element.Id,
             text: element.Title
           };          
           dropdownLib.push(values);
         });
         const items = {
           tab1: dropdownLib,
           tab2: dropdownLib
         };
         resolve(items[wp.properties.tabName]);
       })
       .catch(err => {
         console.log("err is "+err);
         reject(err);
       });
  });
}
private onListItemChange(propertyPath: string, newValue: any): void {
  const oldValue: any = get(this.properties, propertyPath);
  // store new value in web part properties
  update(this.properties, propertyPath, (): any => { return newValue; });
  // refresh web part
  this.render();
}

private _getLibraries(): Promise<any> {
  return new Promise<any>(
    (resolve: (Title: any) => void, reject: (error: any) => void): void => {
      this.context.spHttpClient
        .get(
          this.context.pageContext.web.absoluteUrl +
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
              console.log("inside then2");
              resolve(response.value);
            }
          }
        );
    }
  );
}

}
