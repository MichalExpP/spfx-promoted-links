export interface IPromotedLinksWebPartProps {
    listId: string;
    tileSize:string;
    description: string;
    title: string;
  }
  
  export interface ISPList {
    Title: string;
    Id: string;
  }
  
  export interface ISPLists {
    value: ISPList[];
  }