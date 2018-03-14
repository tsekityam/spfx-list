import * as React from 'react';
import styles from './SpGrid.module.scss';
import { ISpGridProps } from './ISpGridProps';
import { ISpGridState } from './ISpGridState';
import { ISpField } from '../../../interfaces/ISpField';
import { ISpItem } from '../../../interfaces/ISpItem';
var moment = require('moment');

import { sp, ItemAddResult } from "@pnp/sp";
import { autobind } from 'office-ui-fabric-react/lib/Utilities';
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
import { Dialog, DialogType, DialogFooter } from 'office-ui-fabric-react/lib/Dialog';
import { PrimaryButton, DefaultButton } from 'office-ui-fabric-react/lib/Button';

import SpForm from './SpForm';

export default class SpGrid extends React.Component<ISpGridProps, ISpGridState> {
  private _selection: Selection = new Selection({
    onSelectionChanged: () => this.setState({ selectionDetails: this._getSelectionDetails() })
  });

  constructor(props: any) {
    super(props);
    this.state = {
      hideDeleteDialog: true,
      selectionDetails: this._getSelectionDetails(),
      showEditPanel: false,
      formItem: undefined,
    };
  }

  public render(): React.ReactElement<ISpGridProps> {
    return (
      <div>
        {this._getDetailList()}
        {this._getDeleteDialog()}
        {this._getEditPanel()}
      </div>
    );
  }

  public componentWillReceiveProps() {
  }

  private _getDetailList() {
    return <div>
      <CommandBar
        isSearchBoxVisible={false}
        items={this._getCommandBarItems()}
        farItems={this._getCommandBarFarItems()}
      />
      <MarqueeSelection selection={this._selection}>
        <DetailsList
          items={this.props.items}
          columns={this._getColumns()}
          layoutMode={DetailsListLayoutMode.fixedColumns}
          selection={this._selection}
          selectionPreservedOnEmptyClick={true}
          selectionMode={SelectionMode.multiple}
          ariaLabelForSelectionColumn='Toggle selection'
          ariaLabelForSelectAllCheckbox='Toggle selection for all items'
          onItemInvoked={this._showEditingPanel}
        />
      </MarqueeSelection>
    </div>;
  }

  private _getDeleteDialog() {
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

  private _getEditPanel() {
    return (
      <SpForm
          fields={this.props.fields}
          showEditPanel={this.state.showEditPanel}
          onDismiss={this._onCloseEditPanel}
          item={this.state.formItem}
          onSave={this.props.onSave}
          onSaved={this._onSaved}
        />
    );
  }

  /*
   * Grid related
   */

  private _getCommandBarItems(): IContextualMenuItem[] {
    var items: IContextualMenuItem[] = [];

    if (this._selection.getSelectedCount() === 0) {
      items.push({
        key: 'newItem',
        name: 'New',
        icon: 'Add',
        onClick: this._addItem,
      });
    }

    if (this._selection.getSelectedCount() === 1) {
      items.push({
        key: 'editItem',
        name: 'Edit',
        icon: 'Edit',
        onClick: this._editItem,
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
  
  private _getColumns(): IColumn[] {
    var columns: IColumn[] = [];

    this.props.fields.map((item: ISpField, index: number) => {
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

  private _getSelectionDetails(): string {
    return `${this._selection.getSelectedCount()} items selected`;
  }

  @autobind
  private _addItem() {
    this._showEditingPanel();
  }

  @autobind
  private _editItem() {
    if (this._selection.getSelectedCount() === 1) {
      this._showEditingPanel(this._selection.getSelection()[0] as ISpItem);
    }
  }

  @autobind
  private _showDeleteDialog() {
    this.setState({ hideDeleteDialog: false });
  }

  @autobind
  private _cancelSelection() {
    this._selection.setAllSelected(false);
  }

  /*
   *  Delete Dialog related
   */

  private _getDeleteDialogTitle(): string {
    switch (this._selection.getSelectedCount()) {
      case 1:
        return `Delete ${(this._selection.getSelection()[0] as ISpItem).Title}?`;
      default:
        return "Delete?";
    }
  }

  @autobind
  private _closeDeleteDialog() {
    this.setState({ hideDeleteDialog: true });
  }

  @autobind
  private _deleteSelectedItems() {
    var selectedItems: ISpItem[] = [];

    this._selection.getSelection().map((item, index) => {
      selectedItems.push(item as ISpItem);
    });

    this.props.onDeleteSelectedItems(selectedItems).then(() => {
      this._closeDeleteDialog();
    });
  }

  /*
   *  Edit Panel related
   */

  @autobind
  private _showEditingPanel(selectedItem?: ISpItem): void {
    if (selectedItem) {
      this.setState({
        formItem: (selectedItem)
      }, () => {
        this.setState({
          showEditPanel: true
        });
      });
    } else {
      this.setState({
        formItem: undefined
      }, () => {
        this.setState({
          showEditPanel: true
        });
      });
    }
  }

  @autobind
  private _onCloseEditPanel(): void {
    this.setState({
      showEditPanel: false,
      formItem: {}
    });
  }

  @autobind
  private _onSaved(): void {
    this.props.onRefreshItems();
  }
}
