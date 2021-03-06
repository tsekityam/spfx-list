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

import SpGrid from './SpGrid';

export default class SpList extends React.Component<ISpListProps, ISpListState> {
  constructor(props: any) {
    super(props);
    this.state = {
      fields: [],
      items: [],
    };
  }

  public render(): React.ReactElement<ISpListProps> {
    return (
      <div>
        <SpGrid
          items={this.state.items}
          fields={this.state.fields}
          onDeleteSelectedItems={this._deleteItems}
          onRefreshItems={this._onRefreshItems}
          onSave={this._onSaveItemForm}
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
  private _deleteItems(items: ISpItem[]) {
    let list = sp.web.lists.getById(this.props.list);

    let batch = sp.web.createBatch();

    items.map((item, index) => {
      list.items.getById(item.Id).inBatch(batch).delete().then(_ => { });
    });

    return batch.execute().then(d => {
      this._updateListItems(this.props);
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
  private _onRefreshItems(): void {
    this._updateListItems(this.props);
  }
}
