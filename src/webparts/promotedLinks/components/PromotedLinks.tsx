import * as React from 'react';
import styles from './PromotedLinks.module.scss';
import { IPromotedLinksProps, IPromotedLinkDataItem } from './IPromotedLinksProps';
import PromotedLinkItem, { IPromotedLinkItemProps }  from './PromotedLinkItem';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient } from '@microsoft/sp-http';
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";

export interface IPromotedLinksState {
  listData: IPromotedLinkDataItem[];
}

export default class PromotedLinks extends React.Component<IPromotedLinksProps, IPromotedLinksState> {

  constructor(props: IPromotedLinksProps, state: IPromotedLinksState) {
    super(props);

    this.state = { listData: [] };
  }

  public render(): React.ReactElement<IPromotedLinksProps> {

    //To add new value, modify the tileSize PropertyPaneDropdown in the PromotedLinksWebPart.ts file
    //Then add new scss in the PromotedLinks.module.scss
    //Then modify the fi=unction below
    function getTileSizeClass(size){
      switch (size) {
        case '50x100': return styles.promotedLinks50x100;
        case '75x75': return styles.promotedLinks75x75;
        case '75x150': return styles.promotedLinks75x150;
        case '100x100': return styles.promotedLinks100x100;
        case '100x200': return styles.promotedLinks100x200;
        case '113x113': return styles.promotedLinks113x113;
        case '125x125': return styles.promotedLinks125x125;
        case '131x131': return styles.promotedLinks131x131;
        case '142x142': return styles.promotedLinks142x142;
        case '150x150': return styles.promotedLinks150x150;
        case '181x181': return styles.promotedLinks181x181;
        case '200x200': return styles.promotedLinks200x200;
        case '293x160': return styles.promotedLinks293x160;
        default : return styles.promotedLinks125x125;
      }
    }

    return (
      <div className={getTileSizeClass(this.props.tileSize)}>
        <WebPartTitle displayMode={this.props.displayMode} title={this.props.title} updateProperty={this.props.updateProperty} />
        <div className={styles.container}>
          {
            this.state.listData.map((item: IPromotedLinkDataItem) => {
              return <PromotedLinkItem
                title={item.Title}
                description={item.Description}
                imageUrl={item.ImageUrl}
                launchBehavior={item.LaunchBehavior}
                href={item.LinkUrl} />;
            })
          }
          <div style={{clear:'both'}}></div>
        </div>
      </div>
    );
  }

  public componentDidMount(): void {
    this.loadData();
  }

  private loadData(): void {
    if (this.props.isWorkbench) {
      // get mock data in Workbench
      this.setState({
        listData: [
          {
            Title: "Test Item",
            Description: "Test description",
            ImageUrl: "https://media-cdn.tripadvisor.com/media/photo-s/04/a8/17/f5/el-arco.jpg",
            LaunchBehavior: "_blank",
            LinkUrl: "http://www.google.com"
          },
          {
            Title: "Test Item with a Long Title",
            Description: "Test description",
            ImageUrl: "https://pgcpsmess.files.wordpress.com/2014/04/330277-red-fox-kelly-lyon-760x506.jpg",
            LaunchBehavior: "_blank",
            LinkUrl: "http://www.google.com"
          },
          {
            Title: "Test Item",
            Description: "Test description",
            ImageUrl: "https://s-media-cache-ak0.pinimg.com/736x/d6/d4/d7/d6d4d7224687ca3de4a160f5264b5b99.jpg",
            LaunchBehavior: "_self",
            LinkUrl: "Test item with a long description for display."
          }
        ]
      });
    } else {
      // get data from SharePoint
      this.props.spHttpClient.get(`${this.props.siteUrl}/_api/Web/Lists(guid'${this.props.listId}')/Items?$orderby=TileOrder,Title`, SPHttpClient.configurations.v1)
      .then(response => {
        return response.json();
      })
      .then((items: any) => {
        const listItems: IPromotedLinkDataItem[] = [];
        let backgroundImageLocationObject: any = {Description: "", Url: ""};
        let LaunchBehaviorTarget;
        for (let i: number = 0; i < items.value.length; i++) {
          //Check for empty BackgroundImageLocation
          if(items.value[i].BackgroundImageLocation === null){
            backgroundImageLocationObject = {Description: "", Url: ""};
          }else{
            backgroundImageLocationObject = items.value[i].BackgroundImageLocation.Url;
          }
          //Check for chose LaunchBehavior
          if(items.value[i].LaunchBehavior === "In page navigation"){
            LaunchBehaviorTarget = "_self";
          }else if(items.value[i].LaunchBehavior === "New tab"){
            LaunchBehaviorTarget = "_blank";
          }else{LaunchBehaviorTarget = "_blank";}

          listItems.push({
            Title: items.value[i].Title,
            Description: items.value[i].Description,
            ImageUrl: backgroundImageLocationObject,
            LaunchBehavior: LaunchBehaviorTarget,
            LinkUrl: items.value[i].LinkLocation.Url
          });
        }
        this.setState({ listData: listItems });
      }, (err: any) => {
        console.log(err);
      });
    }
  }

  public componentDidUpdate(prevProps: IPromotedLinksProps, prevState: IPromotedLinksState, prevContext: any) {
    if (prevProps.tileSize != this.props.tileSize || prevProps.listId != this.props.listId) {
        this.loadData();
    }
  }
}
