import { ISpField } from "../../../interfaces/ISpField";
import { ISpItem } from "../../../interfaces/ISpItem";
import { ItemAddResult } from "@pnp/sp";

export interface ISpFormProps {
  fields: ISpField[];
  showEditPanel: boolean;
  item?: ISpItem;
  onDismiss: () => void;
  onSave: (item: ISpItem, oldItem: ISpItem) => Promise<ItemAddResult>;
  onSaved?: () => void;
}
