import { ISpField } from "../../../interfaces/ISpField";
import { ISpItem } from "../../../interfaces/ISpItem";
import { ItemAddResult } from "@pnp/sp";

export interface ISpFormProps {
  fields: ISpField[];
  showEditPanel: boolean;
  formItem?: ISpItem;
  onDismiss: () => void;
  onSave: (formItem: ISpItem, oldFformItem: ISpItem) => Promise<ItemAddResult>;
  onSaved?: () => void;
}
