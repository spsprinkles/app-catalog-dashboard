import { LoadingDialog } from "dattatable";
import { Components, ContextInfo } from "gd-sprest-bs";
import { appIndicator } from "gd-sprest-bs/build/icons/svgs/appIndicator";
import { chatSquareDots } from "gd-sprest-bs/build/icons/svgs/chatSquareDots";
import { columnsGap } from "gd-sprest-bs/build/icons/svgs/columnsGap";
import { pencilSquare } from "gd-sprest-bs/build/icons/svgs/pencilSquare";
import { trash } from "gd-sprest-bs/build/icons/svgs/trash";
import { AppForms } from "./itemForms";
import * as jQuery from "jquery";
import { DataSource, IAppItem } from "./ds";
import Strings from "./strings";


/**
 * Main Application
 */
export class App {
    private _el: HTMLElement = null;
    private _forms: AppForms = null;

    // Constructor
    constructor(el: HTMLElement) {
        // Set the global variables
        this._el = el;
        this._forms = new AppForms();

        // Render the template
        this._el.classList.add("bs");
        this._el.innerHTML = `
            <div id="app-dashboard" class="row">
                <div id="app-info" class="col-8"></div>
                <div id="app-actions" class="col-4"></div>
            </div>
        `;

        // Render the dashboard
        this.render();
    }

    // Refreshes the dashboard
    private refresh() {
        // Show a loading dialog
        LoadingDialog.setHeader("Refreshing the Data");
        LoadingDialog.setBody("This will close after the data is loaded.");

        // Load the events
        DataSource.init().then(() => {
            // Clear the element
            while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

            // Render the dashboard
            this.render();

            // Hide the dialog
            LoadingDialog.hide();
        });
    }

    // Renders the dashboard
    private render() {
        // Determine if the user can edit the item
        let canEdit = (DataSource.DocSetItem["AuthorId"] == ContextInfo.userId ||
            (DataSource.DocSetItem["OwnersId"] ? DataSource.DocSetItem["OwnersId"].results.indexOf(ContextInfo.userId) != -1 : false));

        // Render the properties
        Components.ListForm.renderDisplayForm({
            info: DataSource.DocSetInfo,
            el: this._el.querySelector("#app-info"),
            includeFields: [
                "FileLeafRef",
                "DevAppStatus",
                "Owners",
                "SharePointAppCategory",
                "AppProductID",
                "AppVersion",
                "IsAppPackageEnabled"
            ]
        });

        // Render the actions
        Components.ButtonGroup({
            el: this._el.querySelector("#app-actions"),
            isVertical: true,
            buttons: [
                {
                    text: "Edit Properties",
                    iconSize: 20,
                    iconType: pencilSquare,
                    isSmall: true,
                    type: Components.ButtonTypes.OutlineSecondary,
                    onClick: () => {
                        // Display the edit form
                        this._forms.edit(DataSource.DocSetItem.Id, () => {
                            // Refresh the table
                            this.refresh();
                        });
                    }
                },
                {
                    text: "Submit for Review",
                    iconSize: 20,
                    iconType: appIndicator,
                    isDisabled: canEdit && DataSource.DocSetItem.DevAppStatus != "In Review" ? false : true,
                    isSmall: true,
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                        // Display the submit form
                        this._forms.submit(DataSource.DocSetItem, () => {
                            // Refresh the table
                            this.refresh();
                        });
                    }
                },
                {
                    text: "Review this App",
                    iconSize: 20,
                    iconType: chatSquareDots,
                    isDisabled: DataSource.DocSetItem.DevAppStatus != "In Review" ? true : !canEdit,
                    isSmall: true,
                    type: Components.ButtonTypes.OutlinePrimary,
                    onClick: () => {
                        // Display the review form
                        this._forms.review(DataSource.DocSetItem, () => {
                            // Refresh the table
                            this.refresh();
                        });
                    }
                },
                {
                    text: "Delete App/Solution",
                    iconSize: 20,
                    iconType: trash,
                    isDisabled: !canEdit,
                    isSmall: true,
                    type: Components.ButtonTypes.OutlineDanger,
                    onClick: () => {
                        // Display the delete form
                        this._forms.delete(DataSource.DocSetItem, () => {
                            // Refresh the table
                            this.refresh();
                        });
                    }
                },
                {
                    text: "Deploy",
                    iconSize: 20,
                    //iconType: trash,
                    //isDisabled: !canEdit,
                    isSmall: true,
                    type: Components.ButtonTypes.OutlineWarning,
                    onClick: () => {
                        // Deploy the app
                        this._forms.deploy(DataSource.DocSetItem, () => {
                            // Refresh the table
                            this.refresh();
                        });
                    }
                },
                {
                    text: "Retract",
                    iconSize: 20,
                    //iconType: trash,
                    //isDisabled: !canEdit,
                    isSmall: true,
                    type: Components.ButtonTypes.OutlineDanger,
                    onClick: () => {
                        // Retract the app
                        this._forms.retract(DataSource.DocSetItem, () => {
                            // Refresh the table
                            this.refresh();
                        });
                    }
                }
            ]
        });
    }
}