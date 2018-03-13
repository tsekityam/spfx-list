import * as React from 'react';
import styles from './SpList.module.scss';
import { ISpListProps } from './ISpListProps';
import { ISpListState } from './ISpListState';
import { ISpField } from '../../../interfaces/ISpField';
import { ISpItem } from '../../../interfaces/ISpItem';
var moment = require('moment');

import { sp } from "@pnp/sp";
import {
  DetailsList,
  DetailsListLayoutMode,
  Selection,
  IColumn,
  SelectionMode
} from 'office-ui-fabric-react/lib/DetailsList';
import { MarqueeSelection } from 'office-ui-fabric-react/lib/MarqueeSelection';
import { CommandBar } from 'office-ui-fabric-react/lib/CommandBar';
import { IContextualMenuItem } from 'office-ui-fabric-react/lib/ContextualMenu';
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';

import SpForm from './SpForm';

export default class SpList extends React.Component<ISpListProps, ISpListState> {
  private _selection: Selection = new Selection({
    onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() })
  });
  private _formComponents: any[] = [];

  constructor(props: any) {
    super(props);
    this.state = {
      fields: [],
      items: [],
      selectionDetails: this._getSelectionDetails(),
      hideDeleteDialog: true,
      showEditPanel: false,
      formItem: undefined,
      editFormErrors: {}
    };
  }

  public render(): React.ReactElement<ISpListProps> {
    return (
      <div>
        {this.getDetailList()}
        {this.getDeleteDialog()}
        <SpForm
          fields={this.state.fields}
          showEditPanel={this.state.showEditPanel}
          onDismiss={this._onCloseEditPanel}
          formItem={this.state.formItem}
          list={this.props.list}
          onSaved={this._onSaved}
        />
      </div>
    );
  }

  private getDetailList() {
    return <div>
      <CommandBar
        isSearchBoxVisible={false}
        items={this._getCommandBarItems()}
        farItems={this._getCommandBarFarItems()}
      />
      <MarqueeSelection selection={this._selection}>
        <DetailsList
          items={this._getItems()}
          columns={this._getColumns()}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          selection={this._selection}
          selectionPreservedOnEmptyClick={true}
          selectionMode={SelectionMode.multiple}
          ariaLabelForSelectionColumn='Toggle selection'
          ariaLabelForSelectAllCheckbox='Toggle selection for all items'
          onItemInvoked={this._onShowEditPanel}
        />
      </MarqueeSelection>
    </div>;
  }

  private getDeleteDialog() {
    return <Dialog
      hidden={this.state.hideDeleteDialog}
      onDismiss={this._closeDeleteDialog}
      dialogContentProps={{
        type: DialogType.normal,
        title: `${this._getDeleteDialogTitle()}`,
        subText: 'Are you sure you want to deleted selected item(s)?'
      }}
      modalProps={{
        isBlocking: false,
      }}
    >
      <DialogFooter>
        <DefaultButton onClick={this._deleteSelectedItems} text='Delete' />
        <PrimaryButton onClick={this._closeDeleteDialog} text='Cancel' />
      </DialogFooter>
    </Dialog>;
  }

  public componentDidMount() {
    this._updateListFields(this.props);
    this._updateListItems(this.props);
  }

  public componentWillReceiveProps(nextProps) {
    this._updateListFields(nextProps);
    this._updateListItems(nextProps);
  }

  private _getColumns(): IColumn[] {
    var columns: IColumn[] = [];

    this.state.fields.map((item: ISpField, index: number) => {
      columns.push({
        key: item.Id,
        name: item.Title,
        fieldName: item.InternalName,
        minWidth: 100,
        maxWidth: 200,
        isResizable: true,
        ariaLabel: item.Description
      });
    });

    return columns;
  }

  private _getItems() {
    return this.state.items;
  }

  private _getCommandBarItems(): IContextualMenuItem[] {
    var items: IContextualMenuItem[] = [];

    if (this._selection.getSelectedCount() === 0) {
      items.push({
        key: 'newItem',
        name: 'New',
        icon: 'Add',
        onClick: this._onShowEditPanel,
      });
    }

    if (this._selection.getSelectedCount() === 1) {
      items.push({
        key: 'editItem',
        name: 'Edit',
        icon: 'Edit',
        onClick: this._onShowEditPanel,
      });
    }

    if (this._selection.getSelectedCount() > 0) {
      items.push({
        key: 'deletItem',
        name: 'Delete',
        icon: 'Delete',
        onClick: this._showDeleteDialog,
      });
    }

    return items;
  }

  private _getCommandBarFarItems(): IContextualMenuItem[] {
    var items: IContextualMenuItem[] = [];

    if (this._selection.getSelectedCount() > 0) {
      items.push({
        key: 'cancelSelection',
        name: `${this._getSelectionDetails()}`,
        icon: 'Cancel',
        onClick: this._cancelSelection,
      });
    }

    return items;
  }

  private _updateListFields(props): void {
    if (props.list === "") {
      return;
    }

    sp.web.lists.getById(props.list)
      .fields.filter("Hidden eq false and ReadOnlyField eq false and Group eq 'Custom Columns'")
      .get().then((response: ISpField[]) => {
        console.log(response);
        this.setState({
          fields: response
        });
      });
  }

  private _updateListItems(props): void {
    if (props.list === "") {
      return;
    }

    // get all fields that is visible and editable in a list
    sp.web.lists.getById(props.list)
      .items
      .get().then((response) => {
        console.log(response);
        this.setState({
          items: response
        });
      });
  }

  @autobind
  private _showDeleteDialog() {
    this.setState({ hideDeleteDialog: false });
  }

  @autobind
  private _closeDeleteDialog() {
    this.setState({ hideDeleteDialog: true });
  }

  @autobind
  private _cancelSelection() {
    this._selection.setAllSelected(false);
  }

  private _getDeleteDialogTitle(): string {
    switch (this._selection.getSelectedCount()) {
      case 1:
        return `Delete ${(this._selection.getSelection()[0] as ISpItem).Title}?`;
      default:
        return "Delete?";
    }
  }

  @autobind
  private _deleteSelectedItems() {
    let list = sp.web.lists.getById(this.props.list);

    let batch = sp.web.createBatch();

    this._selection.getSelection().map((item, index) => {
      list.items.getById((item as ISpItem).Id).inBatch(batch).delete().then(_ => { });
    });

    batch.execute().then(d => {
      this._updateListItems(this.props);
    });

    this._closeDeleteDialog();
  }

  private _getSelectionDetails(): string {
    return `${this._selection.getSelectedCount()} items selected`;
  }

  @autobind
  private _onCloseEditPanel(): void {
    this.setState({ showEditPanel: false });
  }

  @autobind
  private _onSaved(): void {
    this._updateListItems(this.props);
    this.setState({
      formItem: {}
    });
  }

  @autobind
  private _onShowEditPanel(): void {
    this._formComponents.length = 0;

    if (this._selection.getSelectedCount() === 1) {
      this.setState({
        formItem: (this._selection.getSelection()[0] as ISpItem)
      }, () => {
        this.setState({
          editFormErrors: {},
          showEditPanel: true
        });
      });
    } else {
      this.setState({
        formItem: undefined
      }, () => {
        this.setState({
          editFormErrors: {},
          showEditPanel: true
        });
      });
    }
  }
}
