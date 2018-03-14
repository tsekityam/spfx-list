import * as React from 'react';
import styles from './SpList.module.scss';
import { ISpListProps } from './ISpListProps';
import { ISpListState } from './ISpListState';
import { ISpField } from '../../../interfaces/ISpField';
import { ISpItem } from '../../../interfaces/ISpItem';
var moment = require('moment');

import { sp, ItemAddResult } from "@pnp/sp";
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
import SpGrid from './SpGrid';

export default class SpList extends React.Component<ISpListProps, ISpListState> {
  constructor(props: any) {
    super(props);
    this.state = {
      fields: [],
      items: [],
      showEditPanel: false,
      formItem: undefined,
    };
  }

  public render(): React.ReactElement<ISpListProps> {
    return (
      <div>
        <SpGrid
          items={this.state.items}
          fields={this.state.fields}
          onItemInvoked={() => { }}
          onDeleteSelectedItems={this.onDeleteSelectedItems}
          onShowEditingPanel={this._onShowEditingPanel}
        />
        <SpForm
          fields={this.state.fields}
          showEditPanel={this.state.showEditPanel}
          onDismiss={this._onCloseEditPanel}
          item={this.state.formItem}
          onSave={this._onSaveItemForm}
          onSaved={this._onSaved}
        />
      </div>
    );
  }

  public componentDidMount() {
    this._updateListFields(this.props);
    this._updateListItems(this.props);
  }

  public componentWillReceiveProps(nextProps) {
    this._updateListFields(nextProps);
    this._updateListItems(nextProps);
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
  private _onShowEditingPanel(selectedItem?: ISpItem): void {
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
  private onDeleteSelectedItems(selectedItems: ISpItem[]) {
    let list = sp.web.lists.getById(this.props.list);

    let batch = sp.web.createBatch();

    selectedItems.map((item, index) => {
      list.items.getById(item.Id).inBatch(batch).delete().then(_ => { });
    });

    return batch.execute().then(d => {
      this._updateListItems(this.props);
    });
  }

  @autobind
  private _onCloseEditPanel(): void {
    this.setState({
      showEditPanel: false,
      formItem: {}
    });
  }

  @autobind
  private _onSaveItemForm(formItem: ISpItem, oldFormItem: ISpItem): Promise<ItemAddResult> {
    if (oldFormItem === undefined) {
      // add an item to the list
      return sp.web.lists.getById(this.props.list).items.add(formItem);
    } else {
      // update item in the list
      return sp.web.lists.getById(this.props.list).items.getById(oldFormItem.Id).update(formItem);
    }
  }

  @autobind
  private _onSaved(): void {
    this._updateListItems(this.props);
  }
}
