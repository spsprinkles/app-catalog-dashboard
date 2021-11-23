import { LoadingDialog } from "dattatable";
import { Components, Types } from "gd-sprest-bs";
import { appIndicator } from "gd-sprest-bs/build/icons/svgs/appIndicator";
import { chatSquareDots } from "gd-sprest-bs/build/icons/svgs/chatSquareDots";
import { pencilSquare } from "gd-sprest-bs/build/icons/svgs/pencilSquare";
import { trash } from "gd-sprest-bs/build/icons/svgs/trash";
import * as Common from "./common";
import { AppForms } from "./itemForms";
import { DataSource } from "./ds";
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
                <div id="app-nav" class="col-12"></div>
                <div id="app-info" class="col-8"></div>
                <div id="app-actions" class="col-4"></div>
            </div>
        `.trim();

        // Render the actions
        this.renderActions();

        // Render the navigation
        this.renderNav();

        // Render the info
        this.renderInfo();
    }

    // Refreshes the dashboard
    private refresh() {
        // Show a loading dialog
        LoadingDialog.setHeader("Refreshing the Data");
        LoadingDialog.setBody("This will close after the data is loaded.");

        // Load the events
        DataSource.init().then(() => {
            // Clear the information
            let elInfo = this._el.querySelector("#app-info");
            while (elInfo.firstChild) { elInfo.removeChild(elInfo.firstChild); }

            // Render the info
            this.renderInfo();

            // Hide the dialog
            LoadingDialog.hide();
        });
    }

    // Renders the dashboard
    private renderActions() {
        let canEdit = Common.canEdit(DataSource.DocSetItem);

        // Render a card
        Components.Card({
            el: this._el.querySelector("#app-actions"),
            body: [{
                title: "Actions",
                onRender: el => {
                    let tooltips: Components.ITooltipProps[] = [];

                    // Edit Properties
                    tooltips.push({
                        content: "Displays the edit form to update the app properties.",
                        btnProps: {
                            text: "Edit Properties",
                            iconSize: 20,
                            iconType: pencilSquare,
                            isSmall: true,
                            type: Components.ButtonTypes.OutlineSecondary,
                            onClick: () => {
                                // Display the edit form
                                this._forms.edit(DataSource.DocSetItem.Id, () => {
                                    // Refresh the page
                                    window.location.reload();
                                });
                            }
                        }
                    });

                    // See if this is an owner
                    if (DataSource.DocSetItem.DevAppStatus == "Draft" && Common.isOwner(DataSource.DocSetItem)) {
                        // Submit
                        tooltips.push({
                            content: "Submit the app for review",
                            btnProps: {
                                text: "Submit for Review",
                                iconSize: 20,
                                iconType: appIndicator,
                                isDisabled: !canEdit,
                                isSmall: true,
                                type: Components.ButtonTypes.OutlinePrimary,
                                onClick: () => {
                                    // Display the submit form
                                    this._forms.submit(DataSource.DocSetItem, () => {
                                        // Refresh the page
                                        window.location.reload();
                                    });
                                }
                            }
                        });
                    }

                    // See if this app is in review and ensure this was not submitted by the user
                    if ((DataSource.DocSetItem.DevAppStatus == "Submitted for Review" || DataSource.DocSetItem.DevAppStatus == "In Review")
                        && !Common.isSubmitter(DataSource.DocSetItem)) {
                        // Review button
                        tooltips.push({
                            content: "Start/Continue an assessment of the app.",
                            btnProps: {
                                text: "Review",
                                iconSize: 20,
                                iconType: chatSquareDots,
                                isDisabled: !canEdit,
                                isSmall: true,
                                type: Components.ButtonTypes.OutlinePrimary,
                                onClick: () => {
                                    // Display the review form
                                    this._forms.review(DataSource.DocSetItem, () => {
                                        // Refresh the page
                                        window.location.reload();
                                    });
                                }
                            }
                        });
                    }

                    // Ensure the user is an approver
                    if (DataSource.IsApprover) {
                        // See if the app is not in the catalog
                        let app = DataSource.getAppById(DataSource.DocSetItem.AppProductID);
                        if (app == null) {
                            // Delete
                            tooltips.push({
                                content: "Deletes the app.",
                                btnProps: {
                                    text: "Delete App/Solution",
                                    iconSize: 20,
                                    iconType: trash,
                                    isDisabled: !canEdit,
                                    isSmall: true,
                                    type: Components.ButtonTypes.OutlineDanger,
                                    onClick: () => {
                                        // Display the delete form
                                        this._forms.delete(DataSource.DocSetItem, () => {
                                            // Redirect to the dashboard
                                            window.open(Strings.DashboardUrl, "_self");
                                        });
                                    }
                                }
                            })
                        } else {
                            // See if the app is deployed
                            if (app.Deployed) {
                                // Deploy
                                tooltips.push({
                                    content: "Deploys the solution to the tenant app catalog.",
                                    btnProps: {
                                        text: "Deploy",
                                        iconSize: 20,
                                        //iconType: trash,
                                        isSmall: true,
                                        type: Components.ButtonTypes.OutlineWarning,
                                        onClick: () => {
                                            // Deploy the app
                                            this._forms.deploy(DataSource.DocSetItem, () => {
                                                // Refresh the page
                                                window.location.reload();
                                            });
                                        }
                                    }
                                });
                            } else {
                                // Retract
                                tooltips.push({
                                    content: "Retracts the solution from the tenant app catalog.",
                                    btnProps: {
                                        text: "Retract",
                                        iconSize: 20,
                                        //iconType: trash,
                                        isSmall: true,
                                        type: Components.ButtonTypes.OutlineDanger,
                                        onClick: () => {
                                            // Retract the app
                                            this._forms.retract(DataSource.DocSetItem, () => {
                                                // Refresh the page
                                                window.location.reload();
                                            });
                                        }
                                    }
                                });
                            }
                        }
                    }

                    // Render the actions
                    Components.TooltipGroup({ el, isVertical: true, tooltips });
                }
            }]
        });
    }

    // Renders the information
    private renderInfo() {
        // Render a card
        Components.CardGroup({
            el: this._el.querySelector("#app-info"),
            cards: [
                {
                    body: [{
                        title: "App Information",
                        onRender: el => {
                            // Render the properties
                            Components.ListForm.renderDisplayForm({
                                info: DataSource.DocSetInfo,
                                el,
                                includeFields: [
                                    "FileLeafRef",
                                    "DevAppStatus",
                                    "Owners"
                                ]
                            });
                        }
                    }]
                },
                {
                    body: [{
                        onRender: el => {
                            // Render the properties
                            Components.ListForm.renderDisplayForm({
                                info: DataSource.DocSetInfo,
                                el,
                                includeFields: [
                                    "SharePointAppCategory",
                                    "AppProductID",
                                    "AppVersion",
                                    "IsAppPackageEnabled"
                                ]
                            });
                        }
                    }]
                }
            ]
        });
    }

    // Renders the navigation
    private renderNav() {
        // Generate the left side items
        let items: Components.INavbarItem[] = [
            {
                className: "btn-outline-light",
                isButton: true,
                text: "To Dashboard",
                onClick: () => {
                    // Redirect to the dashboard
                    window.open(Strings.DashboardUrl, "_self");
                }
            }
        ];

        // See if the help url exists
        if (DataSource.Configuration.helpPageUrl) {
            // Add the item
            items.push({
                className: "btn-outline-light ms-2",
                isButton: true,
                text: "Help",
                onClick: () => {
                    // Display in a new tab
                    window.open(DataSource.Configuration.helpPageUrl, "_blank");
                }
            });
        }

        // Generate the right side items
        let itemsEnd: Components.INavbarItem[] = null;
        if (DataSource.Templates && DataSource.Templates.length > 0) {
            // Clear the items
            itemsEnd = [{
                className: "btn-outline-light",
                text: "Templates",
                isButton: true,
                items: []
            }];

            // Parse the templates
            for (let i = 0; i < DataSource.Templates.length; i++) {
                let template = DataSource.Templates[i];

                // Add the item
                itemsEnd[0].items.push({
                    data: template,
                    text: template.Name,
                    onClick: item => { this.uploadTemplate(item.data); }
                });
            }
        }

        // Render the navigation
        Components.Navbar({
            el: this._el.querySelector("#app-nav"),
            brand: "App",
            type: Components.NavbarTypes.Primary,
            items,
            itemsEnd
        });
    }

    // Method to upload a template
    private uploadTemplate(file: Types.SP.File) {
        // Ensure the file data exists
        if (file == null) { return; }

        // Show a loading dialog
        LoadingDialog.setHeader("Copying Template");
        LoadingDialog.setBody("Copying the template to this folder. This dialog will close after it completes.");
        LoadingDialog.show();

        // Get the file
        file.content().execute(data => {
            // Upload the file
            DataSource.DocSetItem.Folder().Files().add(file.Name, true, data).execute(
                // Success
                () => {
                    // Close the dialog
                    LoadingDialog.hide();

                    // Refresh the page
                    window.location.reload();
                },

                // Error
                () => {
                    // TODO
                }
            );
        });
    }
}