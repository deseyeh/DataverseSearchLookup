/* eslint-disable @typescript-eslint/no-redundant-type-constituents */
/* eslint-disable @typescript-eslint/no-unnecessary-type-assertion */
/* eslint-disable @typescript-eslint/no-unsafe-call */
/* eslint-disable @typescript-eslint/no-unsafe-member-access */
/* eslint-disable @typescript-eslint/no-unsafe-assignment */
import { IInputs, IOutputs } from "./generated/ManifestTypes";
 
import * as React from "react";
import LookupRender from "./Components/LookupRender";
import { ILookUpProps } from "./types";
import { AuthoringMode } from "./Components/AuthoringMode";

export class DataverseSearchLookup implements ComponentFramework.ReactControl<IInputs, IOutputs> {
    private notifyOutputChanged: () => void;
    private selectedValue: ComponentFramework.LookupValue[] | undefined;

    /**
     * Empty constructor.
     */
    constructor() {
        // Empty
    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     */
    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary
    ): void {
        this.notifyOutputChanged = notifyOutputChanged;
        const { lookupField, searchParameter, } = context.parameters;
        const fieldLogicName = lookupField?.attributes?.LogicalName ?? "";
        
    }

    private  handleHideControl = (fieldLogicName: string) => {
                    const formContext = Xrm.Page;
                    const fieldControl = formContext.getControl(
                      fieldLogicName
                    ) as Xrm.Controls.StandardControl | null;
                    // Hide or show the field
                    fieldControl?.setVisible(false);
    };

    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     * @returns ReactElement root react element for the control
     */
    public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
        const { lookupField, searchParameter, description, appearance } = context.parameters;
        
        console.log("isAuthoringMode", context.mode.isAuthoringMode as boolean | undefined);
        // Check if we are in authoring mode
        const isAuthoringMode = context.mode.isAuthoringMode as boolean | undefined;
        if (isAuthoringMode) {
            return React.createElement(AuthoringMode, { description: description.raw ?? "", cardAppearance: appearance.raw ?? "filled" });
        }

        const targetEntities = lookupField?.attributes?.Targets;
        const fieldLogicName = lookupField?.attributes?.LogicalName ?? "";
        const selectedValue = lookupField?.raw[0]?.id ?? "";
        const searchParam = searchParameter?.raw ?? "";
        //if no search parameter or there is lookup value, hide the control
        const shouldHide = (selectedValue !== "" || searchParam === "");
        if (shouldHide) {
        this.handleHideControl(fieldLogicName as string);
        }

         const props: ILookUpProps = { 
            context: context,
            fieldLogicName: fieldLogicName,
            currentEntityName: lookupField?.attributes?.EntityLogicalName ?? "",
            lookupRecordId: selectedValue,
            lookupEntityName: lookupField?.raw[0]?.name ?? targetEntities?.[0] ?? "",
            searchParameter: searchParam,
            description: description.raw ?? "",
            cardAppearance: appearance.raw ?? "subtle",
            onConfirm: (selectedOption?: ComponentFramework.LookupValue[]) => { 
                if (selectedOption) {
                    this.selectedValue = selectedOption;
                    this.notifyOutputChanged();
                }
            },
            onDismiss: () => {
                if (fieldLogicName) {
                     this.handleHideControl(fieldLogicName as string);
                    console.log("Dismiss clicked");
                    this.notifyOutputChanged();
                }
            }
        };
        return React.createElement(
            LookupRender, props
        );
    }

    /**
     * It is called by the framework prior to a control receiving new data.
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as "bound" or "output"
     */
    public getOutputs(): IOutputs {
        return {
            lookupField: this.selectedValue
         };
    }

    /**
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}
