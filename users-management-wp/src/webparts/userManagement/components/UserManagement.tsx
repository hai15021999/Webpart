import * as React from 'react';
import styles from './UserManagement.module.scss';
import { IFormUser, IUserManagementProps, IUserManagementState } from './IUserManagementProps';
import { ChoiceGroup, DefaultButton, DetailsList, DetailsListLayoutMode, Dialog, DialogType, Dropdown, IPersonaProps, ISelectableDroppableTextProps, Label, MessageBar, MessageBarType, PrimaryButton, SearchBox, SelectionMode, Shimmer, Spinner, SpinnerSize, TextField, values } from 'office-ui-fabric-react';
import formservice from '../services/formservice';
import { FormDialogAddUser } from './formDialogAddUser/FormDialogAddUser'
import { Web } from 'sp-pnp-js';
import { FormDetalListUsers } from './formDetailListUsers/FormDetailListUsers'
import { FieldTextRenderer } from '@pnp/spfx-controls-react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';

const urlPlaceHolder = 'Example: https://appvity.sharepoint.com/sites/SiteName';

export default class UserManagement extends React.Component<IUserManagementProps, IUserManagementState> {
  constructor(props: IUserManagementProps) {
    super(props);
    this.state = {
      formMessageBarObject: {
        isShow: false,
        text: "",
        stack: "",
        type: MessageBarType.info
      },
      formLoadingDialog: {
        isShow: false,
        title: "",
        isBlocking: true,
        onDismiss: null
      },
      siteUrl: "",
      groupList: [],
      selectedGroup: null,
      peopleList: [],
      selectedUsers: [],
      showDialogAddUser: false,
      isDisableButtonSubmit: true,
      isDisableButtonLoadUser: true,
      formservices: new formservice(),
      isSiteConnected: false,
      web: null,
      groupListFilter: [],
      textFilter: '',
      isLoadedListUser: false,
      stateGroups: [],
      user: null,
    }
  }

  _onFilterChanged = (newValue: string) => {
    let _groupList = this.state.groupList;
    if (newValue === "") {
      this.setState({
        textFilter: '',
        groupListFilter: _groupList,
      })
    } else {
      let _groupListFilter = this.state.groupListFilter;
      _groupListFilter = _groupList.filter(value => {
        return value.text.toLowerCase().indexOf(newValue.toLowerCase()) > -1;
      })
      this.setState({
        textFilter: newValue,
        groupListFilter: _groupListFilter,
        isLoadedListUser: false,
      })
    }
  }

  onRenderOption = (props: ISelectableDroppableTextProps<any>): JSX.Element => {
    return (
      <div>
        <div className={`${styles["app-filter-wrap"]}`}>
          <div className={`${styles["app-filter-label"]}`}>
            <Label>
              Tìm kiếm:
          </Label>
          </div>
          <div className={`${styles["app-filter-control"]}`}>
            <SearchBox
              onChanged={this._onFilterChanged}
              value={this.state.textFilter}
              onClear={() => {
                this.setState({
                  textFilter: '',
                })
              }}
            />
          </div>

        </div>

        <DetailsList
          isHeaderVisible={false}
          selectionMode={SelectionMode.none}
          items={props.options}
          columns={
            [{
              key: 'choice',
              name: "",
              fieldName: "",
              minWidth: 40,
              onRender: (item) => {
                const onChoice = (item) => {
                  this.setState({
                    selectedGroup: item,
                  })
                }
                const option = {
                  key: item.key,
                  text: ""
                };

                return <ChoiceGroup options={[option]} selectedKey={this.state.selectedGroup !== null ? this.state.selectedGroup.key : null}
                  onChange={
                    () => {
                      onChoice(item);
                    }
                  }
                ></ChoiceGroup>;
              }
            },
            {
              key: 'text',
              minWidth: 400,
              name: '---- Select Group ----',
              fieldName: 'text',
              onRender: (item) => {
                return <FieldTextRenderer text={item.text} />;
              }
            }]
          }
        />
      </div>
    )
  }

  onShowMessageBar = (messageBarText: string, messageBarType?: MessageBarType, timeClose?: number) => {
    this.setState({
      formMessageBarObject: {
        isShow: true,
        text: messageBarText,
        type: typeof messageBarType !== "undefined" ? messageBarType : MessageBarType.info
      }
    });
    if (typeof timeClose === "undefined") {
      timeClose = 3000;
    }
    setTimeout(() => {
      this.setState({
        formMessageBarObject: {
          isShow: false
        }
      });
    }, timeClose);
  }

  onChangeSiteUrl = (newValue: string) => {
    let siteUrl = this.state.siteUrl;
    siteUrl = newValue;
    this.setState({
      siteUrl: siteUrl,
    })
  }

  cleanGrid = () => {
    let FormDetailListUsers: any = this.refs.FormDetailListUsers;
    FormDetailListUsers.cleanDetailListItem();
  }

  onChangeSelectedGroup = (option) => {
    let selectedGroup = this.state.selectedGroup;
    selectedGroup = option;
    this.setState({
      selectedGroup: selectedGroup,
      isLoadedListUser: false,
    })
  }

  onAddUserClick = () => {
    let FormDialogAddUser: any = this.refs.FormDialogAddUser;
    FormDialogAddUser.openDialog();
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

  reloadListUser = () => {
    this.cleanGrid();
    let _selectedGroup = this.state.selectedGroup;
    if (_selectedGroup !== null) {
      // this.onShowLoadingDialog(`Hệ thống đang lấy thông tin của các user trong group ${_selectedGroup.text}. Vui lòng chờ trong vài giây...`);
      let { web } = this.state;
      web.siteGroups.getByName(_selectedGroup.text).users.get().then(values => {
        let users = values.map(item => {
          return {
            UserName: item.LoginName,
            Id: item.Id,
            Email: item.Email,
            DisplayName: item.Title,
          }
        });
        this.setState({
          peopleList: users
        })
        // this.onCloseLoadingDialog();
        let FormDetailListUsers: any = this.refs.FormDetailListUsers;
        FormDetailListUsers.loadDetailListItem();
      }).catch(err => {
        this.setState({
          peopleList: [],
          isDisableButtonSubmit: true,
        })
        this.onShowMessageBar(`Đã sảy ra lỗi khi lấy thông tin user trong group ${_selectedGroup.text}.`, MessageBarType.error, 5000);
        this.onCloseLoadingDialog();
      })
    } else {
      window.alert('Vui lòng chọn group để hiển thị những user có trong group đó.')
    }
  }

  onDeleteUserClick = () => {
    let { peopleList } = this.state;
    let _selectedUser: IPersonaProps[] = this.state.selectedUsers;
    if (_selectedUser.length < 1) {
      window.alert('Vui lòng chọn user để xóa khỏi group.')
    } else {
      let message = "Hệ thống sẽ tự động xóa user khỏi Group " + this.state.selectedGroup.text + ". Bạn có muốn tiếp tục? ";
      let confirm = window.confirm(message);
      if (confirm) {
        this.onShowLoadingDialog("Hệ thống đang xóa user khỏi group. Vui lòng chờ trong vài giây...");
        let userLoginName = [];
        _selectedUser.forEach(value => {
          peopleList.forEach(person => {
            if (person.Email === value.secondaryText) {
              userLoginName.push(person.UserName);
            }
          })
        });
        let remove = this.state.formservices.removeUserFromGroup(this.state.web, this.state.selectedGroup.text, userLoginName)
        Promise.all([remove]).then(value => {
          let FormDetailListUsers: any = this.refs.FormDetailListUsers;
          FormDetailListUsers.onRemoveUserClickCallBack();
          this.setState({
            selectedUsers: [],
          })
          if (value) {
            this.onCloseLoadingDialog();
            this.reloadListUser();
            this.onShowMessageBar('Đã hoàn tất việc xóa user khỏi group', MessageBarType.success);
          } else {
            this.onCloseLoadingDialog();
            this.reloadListUser();
            this.onShowMessageBar('Đã sảy ra lỗi khi xóa user khỏi group', MessageBarType.error);
          }
        })
      }
    }
  }

  reloadGrid = () => {
    let FormDetailListUsers: any = this.refs.FormDetailListUsers;
    FormDetailListUsers.loadDetailListItem();
  }

  onSelectUserClick = (selectedUser) => {
    this.setState({
      selectedUsers: selectedUser,
    })
  }

  onShowLoadingDialog = (text: string, timeout?: number) => {
    this.setState({
      formLoadingDialog: {
        isShow: true,
        title: text
      }
    });
    if (timeout !== undefined) {
      setTimeout(() => {
        this.onCloseLoadingDialog();
      }, timeout)
    }
  }

  onCloseLoadingDialog = () => {
    this.setState({
      formLoadingDialog: {
        isShow: false,
      }
    });
  }

  onButtonConnectClick = () => {
    let url = this.state.siteUrl;

    if (url !== "") {
      this.onShowLoadingDialog("Hệ thống đang kết nối vào site. Vui lòng chờ trong vài giây...");
      let web = new Web(url);
      web.siteGroups.get().then(values => {
        let groups = values.filter(item => {
          return item.LoginName.indexOf('SharingLinks.') < 0
        });
        groups = groups.map(item => {
          return { choice: '', key: item.Id, text: item.LoginName }
        })
        this.setState({
          groupList: groups,
          groupListFilter: groups,
          isSiteConnected: true,
          web: web,
          isLoadedListUser: false,
        })
        this.onCloseLoadingDialog();
        this.onShowMessageBar('Đã kết nối tới site thành công.', MessageBarType.success, 5000);
      }).catch(err => {
        this.setState({
          groupList: [],
          groupListFilter: [],
          isSiteConnected: false,
          web: null,
          selectedGroup: null,
          isLoadedListUser: false,
        })
        this.onCloseLoadingDialog();
        this.onShowMessageBar('Không thể kết nối tới site.', MessageBarType.error, 5000);
      })
    } else {
      window.alert('Vui lòng nhập url của site để connect.')
    }
  }

  onButtonLoadUserClick = () => {
    let _selectedGroup = this.state.selectedGroup;
    if (_selectedGroup !== null) {
      this.onShowLoadingDialog(`Hệ thống đang lấy thông tin của các user trong group ${_selectedGroup.text}. Vui lòng chờ trong vài giây...`);
      let { web } = this.state;
      web.siteGroups.getByName(_selectedGroup.text).users.get().then(values => {
        let users = values.map(item => {
          return {
            UserName: item.LoginName,
            Id: item.Id,
            Email: item.Email,
            DisplayName: item.Title,
          }
        });
        this.setState({
          peopleList: users,
          isDisableButtonSubmit: false,
          isLoadedListUser: true,
        })
        this.onCloseLoadingDialog();
        let FormDetailListUsers: any = this.refs.FormDetailListUsers;
        FormDetailListUsers.loadDetailListItem();
      }).catch(err => {
        this.setState({
          peopleList: [],
          isDisableButtonSubmit: true,
          isLoadedListUser: false,
        })
        this.onShowMessageBar(`Đã sảy ra lỗi khi lấy thông tin user trong group ${_selectedGroup.text}.`, MessageBarType.error, 5000);
        this.onCloseLoadingDialog();
      })
    } else {
      window.alert('Vui lòng chọn group để hiển thị những user có trong group đó.')
    }
  }

  onButtonSearchUserClick = () => {
    if (this.state.user === null) {
      window.alert('Vui lòng chọn user để tìm kiếm.')
    } else {
      let { stateGroups, user } = this.state;
      let filter = stateGroups.filter(item => {
        if (item.user.length > 0) {
          for (let i = 0; i < item.user.length; i++) {
            if(item.user[i].Email == user.Email){
              return true;
            }
          }
          return false;
        }
      });
      debugger;
    }
  }

  onCloseDialogCallback = () => {
    this.setState({
      showDialogAddUser: false,
    })
  }

  onButtonOKClickCallback = (status: boolean) => {
    if (status) {
      this.onButtonLoadUserClick();
      this.onShowMessageBar('Đã hoàn tất việc thêm user vào group', MessageBarType.success);
    } else {
      this.onButtonLoadUserClick();
      this.onShowMessageBar('Đã sảy ra lỗi khi thêm user vào group', MessageBarType.error);
    }
  }

  //#region render lại từng form cho từng version
  renderV1Form = (): React.ReactElement<IUserManagementProps> => {
    return (
      <div className={`${styles["userManagement"]}`}>
        {this.state.formLoadingDialog.isShow &&
          <Dialog
            hidden={!this.state.formLoadingDialog.isShow}
            dialogContentProps={{
              type: DialogType.normal
            }}
            modalProps={{
              isBlocking: true,
              containerClassName: 'ms-dialogMainOverride'
            }}>
            <Spinner size={SpinnerSize.large} label={this.state.formLoadingDialog.title} ariaLive='assertive' />
          </Dialog>
        }

        {
          this.state.formMessageBarObject.isShow
          &&
          <MessageBar hidden={!this.state.formMessageBarObject.isShow}
            messageBarType={this.state.formMessageBarObject.type}
            isMultiline={this.state.formMessageBarObject.isMultiline}
            onDismiss={() => {
              this.setState({
                formMessageBarObject: {
                  isShow: false
                }
              })
            }}
            dismissButtonAriaLabel='Close'>
            {this.state.formMessageBarObject.text}
            {this.state.formMessageBarObject.stack != "" &&
              <p>
                {this.state.formMessageBarObject.stack}
              </p>
            }
          </MessageBar>
        }

        <div className={`${styles["app-field-wrap"]}`}>
          <div className={`${styles["app-field-label"]}`}>
            <Label required>Site Url</Label>
          </div>
          <div className={`${styles["app-field-control"]}`}>
            <TextField value={this.state.siteUrl} onChanged={this.onChangeSiteUrl} placeholder={urlPlaceHolder} />
            <DefaultButton text="Connect" onClick={this.onButtonConnectClick} />
          </div>
        </div>
        <div className={`${styles["app-field-wrap"]}`}>
          <div className={`${styles["app-field-label"]}`}>
            <Label required>Group</Label>
          </div>
          <div className={`${styles["app-field-control"]}`}>
            <Dropdown
              options={this.state.groupListFilter}
              selectedKey={this.state.selectedGroup !== null ? this.state.selectedGroup.key : null}
              onChanged={this.onChangeSelectedGroup}
              onRenderList={this.onRenderOption}
            />
            <DefaultButton text="Load User" disabled={!this.state.isSiteConnected} onClick={this.onButtonLoadUserClick} />
          </div>
        </div>
        <div className={`${styles["app-field-wrap"]}`}>
          <FormDetalListUsers
            ref='FormDetailListUsers'
            isLoadedListUser={this.state.isLoadedListUser}
            webPartContext={this.props.webpartContext}
            listAllUsers={this.state.peopleList}
            onSelectUserClickCallBack={this.onSelectUserClick}
            isDataLoaded={false}
          />
        </div>
        <div className={`${styles["app-field-wrap"]}`}>
          <div className={`${styles["app-button-submit"]}`}>
            <PrimaryButton text="Thêm User" disabled={this.state.isDisableButtonSubmit} onClick={this.onAddUserClick} className={`${styles["app-button"]}`} />
            <DefaultButton text="Xóa User" disabled={this.state.isDisableButtonSubmit || (this.state.selectedUsers.length < 1)} onClick={this.onDeleteUserClick} className={`${styles["app-button"]}`} />
            {/* <DefaultButton text="Test check" disabled={this.state.isDisableButtonSubmit} onClick={this.onTest} className={`${styles["app-button"]}`} /> */}
          </div>
        </div>
        <FormDialogAddUser
          ref='FormDialogAddUser'
          isShowPeoplePicker={true}
          web={this.state.web}
          groupName={this.state.selectedGroup !== null ? this.state.selectedGroup.text : null}
          hideDialog={!this.state.showDialogAddUser}
          titleDialog='Thêm mới user'
          onCloseDialogCallback={this.onCloseDialogCallback}
          onButtonOKClickCallback={this.onButtonOKClickCallback}
          webPartContext={this.props.webpartContext}
        />
      </div>
    );
  }

  renderV2Form = (): React.ReactElement<IUserManagementProps> => {
    return (
      <div className={`${styles["userManagement"]}`}>
        {this.state.formLoadingDialog.isShow &&
          <Dialog
            hidden={!this.state.formLoadingDialog.isShow}
            dialogContentProps={{
              type: DialogType.normal
            }}
            modalProps={{
              isBlocking: true,
              containerClassName: 'ms-dialogMainOverride'
            }}>
            <Spinner size={SpinnerSize.large} label={this.state.formLoadingDialog.title} ariaLive='assertive' />
          </Dialog>
        }

        {
          this.state.formMessageBarObject.isShow
          &&
          <MessageBar hidden={!this.state.formMessageBarObject.isShow}
            messageBarType={this.state.formMessageBarObject.type}
            isMultiline={this.state.formMessageBarObject.isMultiline}
            onDismiss={() => {
              this.setState({
                formMessageBarObject: {
                  isShow: false
                }
              })
            }}
            dismissButtonAriaLabel='Close'>
            {this.state.formMessageBarObject.text}
            {this.state.formMessageBarObject.stack != "" &&
              <p>
                {this.state.formMessageBarObject.stack}
              </p>
            }
          </MessageBar>
        }

        <div className={`${styles["app-field-wrap"]}`}>
          <div className={`${styles["app-field-label"]}`}>
            <Label required>Site Url</Label>
          </div>
          <div className={`${styles["app-field-control"]}`}>
            <TextField value={this.state.siteUrl} onChanged={this.onChangeSiteUrl} placeholder={urlPlaceHolder} />
            <DefaultButton text="Connect" onClick={this.getAllGroupInSite} />
          </div>
        </div>
        {/* render lại people picker chổ này */}
        <div className={styles['app-field-wrap']}>
          <div className={styles["app-field-label"]}>
            <Label className={styles["app-text-bold"]} required>Người được thêm vào: </Label>
          </div>
          <div className={styles["app-field-control"]}>
            <PeoplePicker
              context={this.props.webpartContext}
              placeholder="Chọn tài khoản hoặc email...."
              personSelectionLimit={1}
              showtooltip={true}
              onChange={this._getPeoplePickerItems}
              showHiddenInUI={false}
              principalTypes={[PrincipalType.User]}
              ensureUser={true}
              resolveDelay={1000} />
            <DefaultButton text="Tìm kiếm" iconProps={{ iconName: 'Search' }} disabled={!this.state.isSiteConnected || this.state.user === null} onClick={this.onButtonSearchUserClick} />
          </div>
        </div>
        <div className={`${styles["app-field-wrap"]}`}>
          <div className={`${styles["app-button-submit"]}`}>
            <PrimaryButton text="Thêm User" disabled={this.state.isDisableButtonSubmit} onClick={this.onAddUserClick} className={`${styles["app-button"]}`} />
            <DefaultButton text="Xóa User" disabled={this.state.isDisableButtonSubmit || (this.state.selectedUsers.length < 1)} onClick={this.onDeleteUserClick} className={`${styles["app-button"]}`} />
          </div>
        </div>
        <FormDialogAddUser
          ref='FormDialogAddUser'
          isShowPeoplePicker={true}
          web={this.state.web}
          groupName={this.state.selectedGroup !== null ? this.state.selectedGroup.text : null}
          hideDialog={!this.state.showDialogAddUser}
          titleDialog='Thêm mới user'
          onCloseDialogCallback={this.onCloseDialogCallback}
          onButtonOKClickCallback={this.onButtonOKClickCallback}
          webPartContext={this.props.webpartContext}
        />
      </div>
    )
  }

  //#endregion


  public render(): React.ReactElement<IUserManagementProps> {
    switch (this.props.versionName) {
      case 'V1': //version update user when connect to site and select a group
        return (
          this.renderV1Form()
        );
      case 'V2': //version viags yêu cầu nhập user => load những group trong site mà user đó là member
        return (
          this.renderV2Form()
        )
      default: return (<Label>Please config version name of webpart config</Label>)
    }
  }
  //#region update lại logic của tool cho mấy thằng Viags
  // componentDidMount = () => {
  //   this.getAllGroupInSite();
  // }
  getAllGroupInSite = () => {
    let url = this.state.siteUrl;

    if (url !== "") {
      this.onShowLoadingDialog("Hệ thống đang kết nối vào site. Vui lòng chờ trong vài giây...");
      let web = new Web(url);
      web.siteGroups.get().then(async result => {
        let groups = result.filter(item => {
          return item.LoginName.indexOf('SharingLinks.') < 0
        });
        let stateGroups = []
        for (let data of groups) {
          await web.siteGroups.getByName(data.LoginName).users.get().then(values => {
            let users = values.map(item => {
              return {
                UserName: item.LoginName,
                Id: item.Id,
                Email: item.Email,
                DisplayName: item.Title,
              }
            });
            stateGroups.push({
              choice: '', key: data.Id, text: data.LoginName, user: users
            })
          });
        }
        this.setState({
          isSiteConnected: true,
          web: web,
          stateGroups: stateGroups,
        })
        this.onCloseLoadingDialog();
        this.onShowMessageBar('Đã kết nối tới site thành công.', MessageBarType.success, 5000);
      }).catch(err => {
        this.setState({
          stateGroups: [],
          web: null,
          isSiteConnected: false,
        })
        this.onCloseLoadingDialog();
        this.onShowMessageBar('Không thể kết nối tới site.', MessageBarType.error, 5000);
      })
    } else {
      window.alert('Vui lòng nhập url của site để connect.')
    }
  }
  //#endregion
}

