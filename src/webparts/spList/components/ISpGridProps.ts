import { ISpField } from "../../../interfaces/ISpField";
import { ISpItem } from "../../../interfaces/ISpItem";

import { ItemAddResult } from "@pnp/sp";

export interface ISpGridProps {
  fields: ISpField[];
  items: ISpItem[];
  onDeleteSelectedItems: (selectedItems: ISpItem[]) => Promise<void>;
  onRefreshItems: () => void;
  onSave: (item: ISpItem, oldItem: ISpItem) => Promise<ItemAddResult>;
}
