import { ISpField } from "../../../interfaces/ISpField";
import { ISpItem } from "../../../interfaces/ISpItem";

export interface ISpGridProps {
  fields: ISpField[];
  items: ISpItem[];
  onItemInvoked: () => void;
  onDeleteSelectedItems: (selectedItems: ISpItem[]) => Promise<void>;
  onShowEditingPanel: (selectedItem?: ISpItem) => void;
}
