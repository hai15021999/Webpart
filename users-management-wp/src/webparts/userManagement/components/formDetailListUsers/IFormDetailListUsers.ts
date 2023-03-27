import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IPersona, IPersonaProps } from "office-ui-fabric-react";

export interface IFormDetalListUsersProps {
    webPartContext?: WebPartContext;
    listAllUsers: IFormUser[];
    onSelectUserClickCallBack: any;
    isDataLoaded: boolean;
    isLoadedListUser: boolean;
}

export interface IFormDetalListUsersState {
    selectingUsers: IFormUser[];
    detailListItem: IDetailListItem[],
    isDataLoaded: boolean;
}

export interface IDetailListItem {
    checkBox: any,
    user: IPersonaProps;
    isChecked: boolean;
}

export interface IFormUser {
    UserName?: string;
    Id?: number;
    Email?: string;
    DisplayName?: string;
    PictureURL?:string;
    Jobtitle?: string
  }