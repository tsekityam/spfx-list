import { ISpField } from "../../../interfaces/ISpField";
import { ISpItem } from "../../../interfaces/ISpItem";

export interface ISpFormProps {
  fields: ISpField[];
  showEditPanel: boolean;
  formItem?: ISpItem;
  onDismiss: () => void;
  onSaved?: () => void;
  list: string;
}
