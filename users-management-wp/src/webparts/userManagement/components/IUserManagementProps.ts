import { DialogType, MessageBarType } from "office-ui-fabric-react";
import formservice from "../services/formservice";
import {
  WebPartContext
} from '@microsoft/sp-webpart-base';
import { SPHttpClient, HttpClient } from "@microsoft/sp-http";
import { Web } from "sp-pnp-js";

export interface IUserManagementProps {
  webpartContext: WebPartContext;
  spHttpClient: SPHttpClient;
  httpClient: HttpClient;
  versionName: string;
}

export interface IUserManagementState {
  formLoadingDialog?: IFormDialog;
  formMessageBarObject?: IFormMessageBar;
  siteUrl: string;
  groupList: any;
  selectedGroup: any;
  peopleList: IFormUser[];
  selectedUsers: any;
  showDialogAddUser: boolean;
  isDisableButtonLoadUser: boolean;
  isDisableButtonSubmit: boolean;
  formservices: formservice;
  isSiteConnected: boolean;
  web: Web;
  groupListFilter: any;
  textFilter: string;
  isLoadedListUser: boolean;
  stateGroups: any;
  user: any;
}

export interface IFormDialog {
  isShow: boolean;
  title?: string;
  subtext?: string;
  isBlocking?: boolean;
  type?: DialogType;
  onDismiss?: any;
}

export interface IFormMessageBar {
  isShow: boolean;
  text?: string;
  type?: MessageBarType;
  stack?: string;
  isMultiline?: boolean;
  overflowButtonAriaLabel?: string;
}

export interface IFormUser {
  UserName?: string;
  Id?: number;
  Email?: string;
  DisplayName?: string;
  PictureURL?:string;
  Jobtitle?: string
}