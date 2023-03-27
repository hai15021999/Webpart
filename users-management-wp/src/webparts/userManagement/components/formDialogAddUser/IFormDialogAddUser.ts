import { WebPartContext } from "@microsoft/sp-webpart-base";
import { Web } from "sp-pnp-js";
import formservice from "../../services/formservice";

export interface IFormDialogAddUserProps {
    onButtonOKClickCallback: any;
    onCloseDialogCallback: any;
    hideDialog: boolean;
    titleDialog: string;
    isShowPeoplePicker: boolean;
    webPartContext?: WebPartContext;
    web: Web;
    groupName: string;
}

export interface IFormDialogAddUserState {
    user?: IFormUser;
    hideDialog: boolean;
    formservices: formservice;
}

export interface IFormUser {
    UserName?: string;
    Id?: number;
    Email?: string;
    DisplayName?: string;
    PictureURL?:string;
    Jobtitle?: string
  }