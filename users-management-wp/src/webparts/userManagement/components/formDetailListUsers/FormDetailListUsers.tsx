import { FieldTextRenderer } from "@pnp/spfx-controls-react";
import { Checkbox, DetailsList, IColumn, IPersonaProps, ISelection, Label, Persona, PersonaSize, SelectionMode, Shimmer, ShimmerElementsGroup, ShimmerElementType as ElemType } from "office-ui-fabric-react";
import * as React from "react";
import { IFormDialogAddUserProps } from "../formDialogAddUser/IFormDialogAddUser";
import styles from "../UserManagement.module.scss";
import { IDetailListItem, IFormDetalListUsersProps, IFormDetalListUsersState } from "./IFormDetailListUsers";

const _columns: IColumn[] = [
    {
        key: 'checkBox',
        fieldName: 'checkBox',
        name: '',
        minWidth: 50,
    },
    {
        key: 'user',
        fieldName: 'user',
        name: '',
        minWidth: 500,
    }
]

export class FormDetalListUsers extends React.Component<IFormDetalListUsersProps, IFormDetalListUsersState> {
    constructor(props: IFormDetalListUsersProps) {
        super(props);
        this.state = {
            selectingUsers: [],
            detailListItem: [],
            isDataLoaded: false,
        }
    }


    getSelectedUser = () => {

    }


    componentDidMount = () => {
        this.loadDetailListItem();
    }

    onRemoveUserClickCallBack = () => {
        this.setState({
            selectingUsers: [],
        })
    }

    cleanDetailListItem = () => {
        this.setState({
            detailListItem: [],
        })
    }

    loadDetailListItem = () => {
        this.setState({
            isDataLoaded: false,
        })
        let { listAllUsers } = this.props;
        let listItems: IDetailListItem[] = [];
        let item: IPersonaProps[] = listAllUsers.map(user => {
            return {
                imageUrl: null,
                imageInitials: '?',
                primaryText: user.DisplayName,
                onRenderPrimaryText: () => {
                    let displayUserName: any = user.DisplayName;
                    displayUserName = <div className={styles["app-approvalcontrol-user"]}><div className={`${styles['app-approvalcontrol-userdisplayname']}`}>{displayUserName}</div></div>
                    return displayUserName;
                },
                secondaryText: user.Email,
                onRenderSecondaryText: () => {
                    return <div className={styles["app-approvalcontrol-email"]}>{user.Email}</div>;
                }
            }
        });
        item.forEach(element => {
            listItems.push({
                checkBox: '',
                user: element,
                isChecked: false,
            })
        });
        this.setState({
            detailListItem: listItems,
            isDataLoaded: true,
        })

    }

    onRenderItemColumn = (item, index, column) => {
        switch (column.fieldName) {
            case 'checkBox':
                return (
                    <Checkbox
                        defaultChecked={item.isChecked}
                        onChange={() => {
                            let checked = item.isChecked;
                            let items = this.state.detailListItem;
                            let _selectingUser = this.state.selectingUsers;
                            items.forEach(value => {
                                if (value === item) {
                                    value.isChecked = !checked;
                                }
                            });
                            let temp = _selectingUser.map(value => {
                                return value == item.user;
                            })
                            if (temp.length < 1 || temp[0] === false) {
                                if (!checked) {
                                    _selectingUser.push(item.user);
                                }
                            } else {
                                if (checked) {
                                    _selectingUser = _selectingUser.filter(value => {
                                        return value != item.user;
                                    });
                                }
                            }
                            this.setState({
                                detailListItem: items,
                                selectingUsers: _selectingUser,
                            })
                            this.props.onSelectUserClickCallBack(_selectingUser);
                        }}
                    />
                )
            case 'user':
                return (
                    <Persona
                        size={PersonaSize.size24}
                        {...item.user}
                    />
                )
            default: return '';
        }
    }

    public render(): React.ReactElement<IFormDetalListUsersProps> {
        return (
            <Shimmer isDataLoaded={this.state.isDataLoaded} width={'100%'} customElementsGroup={this.getFormShimmerElements()}          >
                <div className={`${styles["app-field-wrap"]}`}>
                    <div className={`${styles["app-field-label"]}`}>
                        <Label>List Users</Label>
                    </div>
                    {this.props.isLoadedListUser ?
                        this.props.listAllUsers.length > 0 ?
                            <div className={`${styles["app-field-control"]}`}>
                                <DetailsList
                                    isHeaderVisible={false}
                                    columns={_columns}
                                    className={`${styles['app-grid']}`}
                                    items={this.state.detailListItem}
                                    selectionMode={SelectionMode.none}
                                    onRenderItemColumn={this.onRenderItemColumn}
                                />
                            </div> : <Label>Không tìm thấy user trong group này.</Label>
                        : ''
                    }
                </div>
            </Shimmer>

        )
    }


    private getFormShimmerElements = (): JSX.Element => {
        return (
            <div>
                <ShimmerElementsGroup
                    width={'100%'}
                    shimmerElements={[
                        { type: ElemType.gap, height: 10, width: '100%' }]}
                />
                <ShimmerElementsGroup
                    width={'100%'}
                    shimmerElements={[
                        { type: ElemType.line, height: 20, width: '20%' },
                        { type: ElemType.gap, width: '5%' },
                        { type: ElemType.line, height: 20, width: '25%' },
                        { type: ElemType.gap, width: '5%' },
                        { type: ElemType.line, height: 20, width: '20%' },
                        { type: ElemType.gap, width: '5%' },
                        { type: ElemType.line, height: 20, width: '15%' },
                        { type: ElemType.gap, width: '5%' },
                        { type: ElemType.line, height: 20, width: '20%' },
                        { type: ElemType.gap, width: '100%' },
                    ]}
                />
                <ShimmerElementsGroup
                    width={'100%'}
                    shimmerElements={[
                        { type: ElemType.gap, height: 10, width: '100%' }]}
                />
                <div style={{ width: '100%' }}>
                    <div>
                        <ShimmerElementsGroup
                            width={'100%'}
                            shimmerElements={[
                                { type: ElemType.gap, height: 5, width: '100%' }]}
                        />
                        <ShimmerElementsGroup
                            width={'100%'}
                            shimmerElements={[
                                { type: ElemType.line, height: 10, width: '10%' },
                                { type: ElemType.gap, height: 10, width: '60%' },
                                { type: ElemType.line, height: 10, width: '10%' },
                                { type: ElemType.gap, height: 10, width: '20%' },
                            ]}
                        />
                        <ShimmerElementsGroup
                            width={'100%'}
                            shimmerElements={[
                                { type: ElemType.gap, height: 5, width: '100%' }]}
                        />
                        <ShimmerElementsGroup
                            width={'100%'}
                            shimmerElements={[
                                { type: ElemType.line, height: 10, width: '20%' },
                                { type: ElemType.gap, height: 10, width: '50%' },
                                { type: ElemType.line, height: 10, width: '10%' },
                                { type: ElemType.gap, height: 10, width: '20%' },
                            ]}
                        />
                    </div>
                    <div>
                        <ShimmerElementsGroup
                            width={'100%'}
                            shimmerElements={[
                                { type: ElemType.gap, height: 5, width: '100%' }]}
                        />
                        <ShimmerElementsGroup
                            width={'100%'}
                            shimmerElements={[
                                { type: ElemType.line, height: 10, width: '10%' },
                                { type: ElemType.gap, height: 10, width: '60%' },
                                { type: ElemType.line, height: 10, width: '8%' },
                                { type: ElemType.gap, height: 10, width: '22%' },
                            ]}
                        />
                        <ShimmerElementsGroup
                            width={'100%'}
                            shimmerElements={[
                                { type: ElemType.gap, height: 5, width: '100%' }]}
                        />
                        <ShimmerElementsGroup
                            width={'100%'}
                            shimmerElements={[
                                { type: ElemType.line, height: 10, width: '20%' },
                                { type: ElemType.gap, height: 10, width: '50%' },
                                { type: ElemType.line, height: 10, width: '15%' },
                                { type: ElemType.gap, height: 10, width: '15%' },
                            ]}
                        />
                    </div>
                    <div>
                        <ShimmerElementsGroup
                            width={'100%'}
                            shimmerElements={[
                                { type: ElemType.gap, height: 5, width: '100%' }]}
                        />
                        <ShimmerElementsGroup
                            width={'100%'}
                            shimmerElements={[
                                { type: ElemType.line, height: 10, width: '10%' },
                                { type: ElemType.gap, height: 10, width: '60%' },
                                { type: ElemType.line, height: 10, width: '10%' },
                                { type: ElemType.gap, height: 10, width: '20%' },
                            ]}
                        />
                        <ShimmerElementsGroup
                            width={'100%'}
                            shimmerElements={[
                                { type: ElemType.gap, height: 5, width: '100%' }]}
                        />
                        <ShimmerElementsGroup
                            width={'100%'}
                            shimmerElements={[
                                { type: ElemType.line, height: 10, width: '30%' },
                                { type: ElemType.gap, height: 10, width: '40%' },
                                { type: ElemType.line, height: 10, width: '15%' },
                                { type: ElemType.gap, height: 10, width: '15%' },
                            ]}
                        />
                    </div>
                </div>

            </div>
        );
    }
}