import { DefaultButton, Dialog, DialogFooter, DialogType, Label, values } from "office-ui-fabric-react";
import * as React from "react";
import styles from "./FormDialogAddUser.module.scss";
import { IFormDialogAddUserProps, IFormDialogAddUserState } from './IFormDialogAddUser';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import formservice from "../../services/formservice";

export class FormDialogAddUser extends React.Component<IFormDialogAddUserProps, IFormDialogAddUserState> {
    constructor(props: IFormDialogAddUserProps) {
        super(props);
        this.state = {
            user: null,
            hideDialog: this.props.hideDialog,
            formservices: new formservice(),
        }
    }

    _closeDialog = () => {
        this.setState({
            hideDialog: true
        });
        this.props.onCloseDialogCallback();
    }

    eventButonOKClick = () => {
        let add = this.state.formservices.addUserToGroup(this.props.web, this.props.groupName, this.state.user);
        let status: boolean = false;
        Promise.all([add]).then(value => {
            status = value[0];
            this._closeDialog();
            if (typeof this.props.onButtonOKClickCallback !== "undefined" && this.props.onButtonOKClickCallback !== null) {
                this.props.onButtonOKClickCallback(status);
            }
        })

    }

    _getPeoplePickerItems = (items: any[]) => {
        if (items.length > 0) {
            let user = {
                UserName: items[0].loginName,
                DisplayName: items[0].text,
                Email: items[0].secondaryText
            };
            this.setState({
                user: user
            });
        } else {
            let _user = null
            this.setState({
                user: _user,
            })
        }
    }

    openDialog = () => {
        this.setState({
            hideDialog: false,
        })
    }

    public render(): React.ReactElement<IFormDialogAddUserProps> {
        let { titleDialog, isShowPeoplePicker } = this.props;
        return (
            <Dialog
                className={styles['app-dialog-wrap']}
                hidden={this.state.hideDialog}
                onDismiss={this._closeDialog}
                dialogContentProps={{
                    type: DialogType.close,
                    title: titleDialog,
                    closeButtonAriaLabel: "Close"
                }}
                modalProps={{
                    isBlocking: false,
                }}
            >
                <div className={styles['app-dialog-content']}>
                    {
                        isShowPeoplePicker &&
                        <div className={styles['app-field-wrap']}>
                            <div className={styles["app-field-label"]}>
                                <Label className={styles["app-text-bold"]} required>Người được thêm vào: </Label>
                            </div>
                            <div className={styles["app-field-control"]}>
                                <PeoplePicker
                                    context={this.props.webPartContext}
                                    // titleText="People Picker"
                                    placeholder="Chọn tài khoản hoặc email...."
                                    personSelectionLimit={1}
                                    showtooltip={true}
                                    onChange={this._getPeoplePickerItems}
                                    showHiddenInUI={false}
                                    principalTypes={[PrincipalType.User]}
                                    ensureUser={true}
                                    resolveDelay={1000} />
                            </div>
                        </div>
                    }

                </div>
                <DialogFooter>
                    <DefaultButton onClick={this.eventButonOKClick} text="OK" disabled={this.state.user === null} />
                    <DefaultButton onClick={this._closeDialog} text="Cancel" />
                </DialogFooter>
            </Dialog>
        );
    }
}