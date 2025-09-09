import { IInputs } from "./generated/ManifestTypes";

export interface ILookUpProps {
  context: ComponentFramework.Context<IInputs>;
  fieldLogicName: string;
  currentEntityName: string;
  lookupRecordId: string;
  lookupEntityName: string;
  searchParameter: string;
  description: string; 
  cardAppearance: "filled" | "subtle" | "outline" | "filled-alternative";
  onConfirm: (selectedOption?: ComponentFramework.LookupValue[]) => void;
  onDismiss: () => void;
}