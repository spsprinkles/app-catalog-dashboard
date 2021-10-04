import { ContextInfo, Components, Helper, IconTypes, $REST, List, SPTypes, Types, jQuery  } from "gd-sprest-bs";
import Strings from "../strings";
import Toast from "./toast";

// Filter Properties
export interface IFilterProps {
    el: HTMLElement;
    onChange: (value: string) => void;
    onRefresh: () => void;
}

/**
 * Filter
 */
export class Filter {
    private _props: IFilterProps = null;

    // Constructor
    constructor(props: IFilterProps) {
        // Save the parameters
        this._props = props;

        // Render the filter
        this.render();
    }

    // Render the filter
    private render() {
        // Render a card
        Components.Card({
            el: this._props.el,
            body: [
                {
                    onRender: (el) => {
                        // Create a div to wrap the new item button in, giving it a white background
                        let elDivNew = document.createElement("div");
                        elDivNew.classList.add("bg-white");
                        elDivNew.classList.add("d-inline-flex");
                        elDivNew.classList.add("ml-auto");
                        elDivNew.classList.add("rounded");
                        elDivNew.style.marginRight = "10px";
                        //el.appendChild(elDivNew);
                        // Define the new item button
                        Components.Button({
                            el: elDivNew,
                            //iconType: IconTypes.FilePlus, //CloudUpload Plus Upload Box CodeSqaure
                            iconSize: 20,
                            title: "Add/Update an app/solution package",
                            text: " Add/Update App ",
                            type: Components.ButtonTypes.OutlineSecondary,
                            onClick: () => {
                                Helper.SP.ModalDialog.showModalDialog({
                                    title: "Upload App/Solution",
                                    url: "AppUpload.aspx",
                                    autoSize: true, //must use for post upload opening of EditForm
                                    //width: 515,
                                    //height: 210,
                                    dialogReturnValueCallback: (result, args) => {
                                        //if (result == SPTypes.ModalDialogResult.OK) {
                                        //Always run, even if modal was canceled (could be after app was uploaded)
                                            if (args != "NotInitialUpload")
                                                Toast.showMsg("/_layouts/images/centraladmin_upgradeandmigration_upgradeandpatchmanagement_32x32.png", "App Loading In Progress", "In under a minute you can test your app");

                                            // Refresh the table
                                            this._props.onRefresh();
                                        //}
                                    }
                                });
                            }
                        });

                        let elDivRefresh = document.createElement("div");
                        elDivRefresh.classList.add("bg-white");
                        elDivRefresh.classList.add("d-inline-flex");
                        elDivRefresh.classList.add("ml-auto");
                        elDivRefresh.classList.add("rounded");
                        elDivRefresh.style.marginRight = "25px";
                        // Define the new item button
                        Components.Button({
                            el: elDivRefresh,
                            //id: "btnRefresh",
                            iconType: IconTypes.ArrowRepeat, //CloudUpload Plus Upload Box CodeSqaure
                            iconSize: 20,
                            className: "filterTopBtn",
                            title: "Refresh table",
                            type: Components.ButtonTypes.OutlineSecondary,

                            onClick: () => {
                                this._props.onRefresh();
                            }
                        });

                        // The buttons above are added to the DOM later
                        //----------------------------------------------

                        // Render "choices" as switches
                        Components.CheckboxGroup({
                            el,
                            isInline: true,
                            type: Components.CheckboxGroupTypes.Switch,
                            items: [
                                { label: "Show apps from others" },
                                { label: "Show all apps for review" }
                            ],
                            onChange: (item: Components.ICheckboxGroupItem) => {
                                // Call the change event
                                /*if (item) {
                                    switch (item.label)  {
                                        case "Show apps from others":
                                            this._props.onChange("");
                                            break;
                                        case "Show all apps for review":
                                            this._props.onChange("DevAppStatus eq 'In Review'");
                                            break;
                                    }
                                }
                                else //Default
                                    this._props.onChange("Owners eq '" + ContextInfo.userId + "'");
                                */
                                this._props.onChange(item ? item.label : "");
                            }
                        });

                        jQuery(el.querySelector(".col-12")).prepend(elDivRefresh);
                        jQuery(el.querySelector(".col-12")).prepend(elDivNew);
                    }
                }
            ]
        });
    }
}