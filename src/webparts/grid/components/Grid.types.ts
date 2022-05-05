import { IDropdownOption } from "office-ui-fabric-react";

export interface IGridProps {
  siteurl: string;

}
export interface IGridState {
  CardsData: IGridItem[];
  currentPage: any;
  hideDialog: boolean;
  CardsDataPerPage: any;
  file: string;
  type: string;
  CardImagelink: string;
  cardTitle: string;
  downloadPdf: string;
  downloadPPT: string;
  Id: string;
  Categoryitems: IDropdownOption[];
}
export interface ISPDocuments {
  value: ISPDocument[];
}

export interface ISPDocument {
  Title: string;
  Id: string;
  Url: string;
  Name: string;

}
export interface IGridItem {
  thumbnail: string;
  title: string;
  name: string;
  profileImageSrc: string;
  location: string;
  activity: string;
}
