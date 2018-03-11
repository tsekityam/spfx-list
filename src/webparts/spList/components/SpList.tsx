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
import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import { ChoiceGroup } from 'office-ui-fabric-react/lib/ChoiceGroup';
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { DatePicker, IDatePickerStrings, IDatePicker } from 'office-ui-fabric-react/lib/DatePicker';
import { ITextField } from 'office-ui-fabric-react/lib/components/TextField';
import { stringIsNullOrEmpty } from '@pnp/common';

const DayPickerStrings: IDatePickerStrings = {
  months: [
    'January',
    'February',
    'March',
    'April',
    'May',
    'June',
    'July',
    'August',
    'September',
    'October',
    'November',
    'December'
  ],

  shortMonths: [
    'Jan',
    'Feb',
    'Mar',
    'Apr',
    'May',
    'Jun',
    'Jul',
    'Aug',
    'Sep',
    'Oct',
    'Nov',
    'Dec'
  ],

  days: [
    'Sunday',
    'Monday',
    'Tuesday',
    'Wednesday',
    'Thursday',
    'Friday',
    'Saturday'
  ],

  shortDays: [
    'S',
    'M',
    'T',
    'W',
    'T',
    'F',
    'S'
  ],

  goToToday: 'Go to today',

  isRequiredErrorMessage: 'Cannot be empty'
};

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
      formItem: {},
      editFormErrors: {}
    };
  }

  public render(): React.ReactElement<ISpListProps> {
    return (
      <div>
        {this.getDetailList()}
        {this.getDeleteDialog()}
        {this.getEditPanel()}
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

  private getEditPanel() {
    var headerText = "";

    if (this._selection.getSelectedCount() === 0) {
      headerText = "New Item";
    } else {
      headerText = `Edit ${(this._selection.getSelection()[0] as ISpItem).Title}`;
    }

    var components: JSX.Element[] = [];
    this.state.fields.map((field: ISpField, index: number) => {
      components.push(
        this._getComponentByField(field)
      );
    });

    return <Panel
      isOpen={this.state.showEditPanel}
      type={PanelType.smallFixedFar}
      headerText={headerText}
      onRenderFooterContent={this._onRenderFooterContent}
    >
      {components}
    </Panel>;
  }

  public componentDidMount() {
    this._updateListFields(this.props);
    this._updateListItems(this.props);
  }

  public componentWillReceiveProps(nextProps) {
    this._updateListFields(nextProps);
    this._updateListItems(nextProps);
  }

  private _getComponentByField(field: ISpField): JSX.Element {
    switch (field.TypeAsString) {
      case "Currency":
        return (
          <TextField
            componentRef={(component: ITextField) => { this._formComponents.push(component); }}
            label={field.Title}
            type='number'
            required={field.Required}
            onGetErrorMessage={(value) => { return this._validate(value, field); }}
            errorMessage={this.state.editFormErrors[field.InternalName]}
            onChanged={(value) => { return this._onValueChanged(value, field); }}
            validateOnFocusOut={true}
            validateOnLoad={false}
            value={this.state.formItem[field.InternalName]}
          />
        );
      case "DateTime":
        return (
          <DatePicker
            componentRef={(component: IDatePicker) => { this._formComponents.push(component); }}
            label={field.Title}
            isRequired={field.Required}
            minDate={moment().toDate()}
            value={this._getDateOfField(field)}
            onSelectDate={(date) => { return this._onValueChanged(date, field); }}
            strings={DayPickerStrings}
          />
        );
      case "Note":
        return (
          <TextField
            componentRef={(component: ITextField) => { this._formComponents.push(component); }}
            label={field.Title}
            required={field.Required}
            multiline
            rows={4}
            onChanged={(val) => { return this._onValueChanged(val, field); }}
            onGetErrorMessage={(value) => { return this._validate(value, field); }}
            errorMessage={this.state.editFormErrors[field.InternalName]}
            validateOnFocusOut={true}
            validateOnLoad={false}
            value={this.state.formItem[field.InternalName]}
          />
        );
      case "Text":
        return (
          <TextField
            componentRef={(component: ITextField) => { this._formComponents.push(component); }}
            label={field.Title}
            required={field.Required}
            onChanged={(value) => { return this._onValueChanged(value, field); }}
            onGetErrorMessage={(value) => { return this._validate(value, field); }}
            errorMessage={this.state.editFormErrors[field.InternalName]}
            validateOnFocusOut={true}
            validateOnLoad={false}
            value={this.state.formItem[field.InternalName]}
          />
        );
      default:
        return (
          <TextField
            label={field.Title}
            disabled={true}
            placeholder={`${field.TypeAsString} is not supported yet`}
          />
        );
    }
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
  private _onRenderFooterContent(): JSX.Element {
    return (
      <div>
        <PrimaryButton
          onClick={this._onSaveEditForm}
        >
          Save
        </PrimaryButton>
        <DefaultButton
          onClick={this._onCloseEditPanel}
        >
          Cancel
        </DefaultButton>
      </div>
    );
  }

  @autobind
  private _onSaveEditForm(): void {

    var canSave: boolean = true;
    var editFormErrors: {} = {};
    this.state.fields.map((field, index) => {
      var error = this._validate(this.state.formItem[field.InternalName], field);
      editFormErrors[field.InternalName] = error;
      canSave = canSave && stringIsNullOrEmpty(error);
    });

    this.setState({ editFormErrors: editFormErrors });

    var formItem = {};
    console.log(this.state.formItem);

    this.state.fields.map((field, index) => {
      formItem[field.InternalName] = this.state.formItem[field.InternalName];
    });

    if (canSave) {
      if (this._selection.getSelectedCount() === 0) {
        // add an item to the list
        sp.web.lists.getById(this.props.list).items.add(formItem).then((iar: ItemAddResult) => {
          this._updateListItems(this.props);
          this.setState({
            formItem: {}
          });
        }).catch((error: any) => {
          console.log(error);
        });
      } else {
        // update item in the list
        sp.web.lists.getById(this.props.list).items.getById(this.state.formItem.Id).update(formItem).then((iar: ItemAddResult) => {
          this._updateListItems(this.props);
          this.setState({
            formItem: {}
          });
        }).catch((error: any) => {
          console.log(error);
        });
      }
    }

    this.setState({ showEditPanel: false });
  }

  @autobind
  private _onCloseEditPanel(): void {
    this.setState({ showEditPanel: false });
  }

  @autobind
  private _onShowEditPanel(): void {
    this._formComponents.length = 0;

    if (this._selection.getSelectedCount() === 1) {
      this.setState({
        formItem: (this._selection.getSelection()[0] as ISpItem)
      });
    }

    this.setState({
      editFormErrors: {},
      showEditPanel: true
    });
  }

  private _onValueChanged(value: any, field: ISpField) {
    var formItem = this.state.formItem;
    formItem[field.InternalName] = value;
    this.setState({
      formItem: formItem
    }, () => { console.log(this.state.formItem); });
  }

  private _validate(value: string, field: ISpField): string {
    if (field.Required && stringIsNullOrEmpty(value)) {
      return "Cannot be empty";
    }

    switch (field.TypeAsString) {
      case "Currency": {
        if (Number(value) < 0) {
          return "Cannot be smaller then 0";
        }
        return "";
      }
    }

    return "";
  }

  private _getDateOfField(field: ISpField) {
    var value: any = this.state.formItem[field.InternalName];

    if (typeof value === "string") {
      return moment(value).toDate();
    }

    return value;
  }
}
