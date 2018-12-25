import { SPHttpClient } from "@microsoft/sp-http";
export interface IEasyTabsWebpartProps {
  listName: string;
  order: string;
  numberOfItems: number;
  style: string;
  item: string;
  spHttpClient: SPHttpClient;
  siteUrl: string;
}
