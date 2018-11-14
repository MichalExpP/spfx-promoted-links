import { IPromotedLinksWebPartProps } from '../IPromotedLinksWebPartProps';
import { SPHttpClient } from '@microsoft/sp-http';
import { DisplayMode } from '@microsoft/sp-core-library';

export interface IPromotedLinksProps extends IPromotedLinksWebPartProps {
  isWorkbench: boolean;
  siteUrl: string;
  spHttpClient: SPHttpClient;
  title: string;
  displayMode: DisplayMode;
  updateProperty: (value: string) => void;
}

export interface IPromotedLinkDataItem {
  Title: string;
  ImageUrl: string;
  Description: string;
  LaunchBehavior: string;
  LinkUrl: string;
}

