import { LoadingDialog } from "dattatable";
import { Components, Types } from "gd-sprest-bs";
import { appIndicator } from "gd-sprest-bs/build/icons/svgs/appIndicator";
import { caretRightFill } from "gd-sprest-bs/build/icons/svgs/caretRightFill";
import { chatSquareDots } from "gd-sprest-bs/build/icons/svgs/chatSquareDots";
import { fileEarmarkMedical } from "gd-sprest-bs/build/icons/svgs/fileEarmarkMedical";
import { pencilSquare } from "gd-sprest-bs/build/icons/svgs/pencilSquare";
import { questionLg } from "gd-sprest-bs/build/icons/svgs/questionLg";
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
                <div id="app-info" class="col-12"></div>
            </div>
        `.trim();

        // Render the navigation
        this.renderNav();

        // Render the info
        this.renderInfo();
        
        // Render the actions
        this.renderActions();
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
            el: this._el.querySelector("#app-info .card-group"),
            body: [{
                title: "Actions",
                onRender: el => {
                    let tooltips: Components.ITooltipProps[] = [];

                    // See if the user can edit
                    if (canEdit) {
                        // Edit Properties
                        tooltips.push({
                            content: "Displays the edit form to update the app properties.",
                            btnProps: {
                                text: "Edit Properties",
                                iconClassName: "me-1",
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
                    }

                    // See if this is an owner
                    if ((DataSource.DocSetItem.DevAppStatus == "Draft" || DataSource.DocSetItem.DevAppStatus == "Requires Attention") && Common.canEdit(DataSource.DocSetItem)) {
                        // Submit
                        tooltips.push({
                            content: "Submit the app for review",
                            btnProps: {
                                text: "Submit for Review",
                                iconClassName: "me-1",
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
                    if (DataSource.DocSetItem.DevAppStatus.indexOf("Review") > 0 &&
                        (!Common.isSubmitter(DataSource.DocSetItem) || DataSource.IsApprover)) {
                        // Review button
                        tooltips.push(
                            {
                                content: "Start/Continue an assessment of the app.",
                                btnProps: {
                                    text: "Review",
                                    iconClassName: "me-1",
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
                            },
                            {
                                content: "Sends the request back to the developer(s).",
                                btnProps: {
                                    text: "Reject",
                                    iconClassName: "me-1",
                                    iconSize: 20,
                                    iconType: chatSquareDots,
                                    isSmall: true,
                                    type: Components.ButtonTypes.OutlineDanger,
                                    onClick: () => {
                                        // Display the reject form
                                        this._forms.reject(DataSource.DocSetItem, () => {
                                            // Refresh the page
                                            window.location.reload();
                                        });
                                    }
                                }
                            }
                        );
                    }

                    // Ensure we are not in a draft state, and see if the site app catalog exists
                    if (DataSource.DocSetItem.DevAppStatus != "Draft" && DataSource.SiteCollectionAppCatalogExists) {
                        // See if the app is not in the site app catalog
                        let app = DataSource.getSiteCollectionAppById(DataSource.DocSetItem.AppProductID);

                        // See if the app is deployed
                        if (app && app.Deployed) {
                            // Retract
                            tooltips.push({
                                content: "Retracts the solution from the site app catalog.",
                                btnProps: {
                                    text: "Retract from Site Catalog",
                                    iconClassName: "me-1",
                                    iconSize: 20,
                                    //iconType: trash,
                                    isSmall: true,
                                    type: Components.ButtonTypes.OutlineDanger,
                                    onClick: () => {
                                        // Retract the app
                                        this._forms.retract(DataSource.DocSetItem, false, () => {
                                            // Refresh the page
                                            window.location.reload();
                                        });
                                    }
                                }
                            });
                        } else {
                            // Deploy
                            tooltips.push({
                                content: "Deploys the solution to the site app catalog.",
                                btnProps: {
                                    text: "Deploy to Site Catalog",
                                    iconClassName: "me-1",
                                    iconSize: 20,
                                    //iconType: trash,
                                    isSmall: true,
                                    type: Components.ButtonTypes.OutlineWarning,
                                    onClick: () => {
                                        // Deploy the app
                                        this._forms.deploy(DataSource.DocSetItem, false, () => {
                                            // Refresh the page
                                            window.location.reload();
                                        });
                                    }
                                }
                            });
                        }
                    }

                    // Ensure the user is an approver
                    if (DataSource.IsApprover) {
                        // See if the app is not in the tenant app catalog
                        let app = DataSource.getTenantAppById(DataSource.DocSetItem.AppProductID);

                        // See if the app hasn't been deployed
                        if (app == null) {
                            // Delete
                            tooltips.push({
                                content: "Deletes the app.",
                                btnProps: {
                                    text: "Delete App/Solution",
                                    iconClassName: "me-1",
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
                            });
                        }

                        // Assessment button
                        tooltips.push({
                            content: "View the last assessment of the app.",
                            btnProps: {
                                text: "Assessment",
                                iconClassName: "me-1",
                                iconSize: 20,
                                iconType: chatSquareDots,
                                isSmall: true,
                                type: Components.ButtonTypes.OutlinePrimary,
                                onClick: () => {
                                    // Display the last assessment
                                    this._forms.lastAssessment(DataSource.DocSetItem);
                                }
                            }
                        });

                        // See if we are requesting approval
                        if (DataSource.DocSetItem.DevAppStatus == "Requesting Approval") {
                            // Retract
                            tooltips.push({
                                content: "Approves the application for deployment.",
                                btnProps: {
                                    text: "Approve",
                                    iconClassName: "me-1",
                                    iconSize: 20,
                                    //iconType: trash,
                                    isSmall: true,
                                    type: Components.ButtonTypes.OutlineSuccess,
                                    onClick: () => {
                                        // Approve the app
                                        this._forms.approve(DataSource.DocSetItem, () => {
                                            // Refresh the page
                                            window.location.reload();
                                        });
                                    }
                                }
                            });
                        }
                        // Else, see if the app is approved
                        else if (DataSource.DocSetItem.DevAppStatus == "Approved") {
                            // See if the app is deployed
                            if (app && app.Deployed) {
                                // Retract
                                tooltips.push({
                                    content: "Retracts the solution from the tenant app catalog.",
                                    btnProps: {
                                        text: "Retract from Tenant",
                                        iconClassName: "me-1",
                                        iconSize: 20,
                                        //iconType: trash,
                                        isSmall: true,
                                        type: Components.ButtonTypes.OutlineDanger,
                                        onClick: () => {
                                            // Retract the app
                                            this._forms.retract(DataSource.DocSetItem, true, () => {
                                                // Refresh the page
                                                window.location.reload();
                                            });
                                        }
                                    }
                                });
                            } else {
                                // Deploy
                                tooltips.push({
                                    content: "Deploys the solution to the tenant app catalog.",
                                    btnProps: {
                                        text: "Deploy to Tenant",
                                        iconClassName: "me-1",
                                        iconSize: 20,
                                        //iconType: trash,
                                        isSmall: true,
                                        type: Components.ButtonTypes.OutlineWarning,
                                        onClick: () => {
                                            // Deploy the app
                                            this._forms.deploy(DataSource.DocSetItem, true, () => {
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
                    let ttg = Components.TooltipGroup({
                        el,
                        className: "w-50",
                        isVertical: true,
                        tooltipOptions: {
                            maxWidth: "200px"
                        },
                        tooltipPlacement: Components.TooltipPlacements.Right,
                        tooltips
                    });
                    // Fix weird alignment issue
                    (ttg.el.querySelector("button:first-child") as HTMLButtonElement).style.marginLeft = "-1px";
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
                        title: "App Details",
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
        // Generate the nav items
        let itemsEnd: Components.INavbarItem[] = null;
        if (DataSource.Templates && DataSource.Templates.length > 0) {
            // Clear the items
            itemsEnd = [{
                className: "btn-outline-light ms-2 ps-1 pt-1",
                iconClassName: "me-1",
                iconSize: 24,
                iconType: fileEarmarkMedical,
                isButton: true,
                text: "Templates",
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

        // See if the help url exists
        if (DataSource.Configuration.helpPageUrl) {
            // Add the item
            itemsEnd.push({
                className: "btn-outline-light ms-2 ps-1 pt-1",
                iconSize: 24,
                iconType: questionLg,
                isButton: true,
                text: "Help",
                onClick: () => {
                    // Display in a new tab
                    window.open(DataSource.Configuration.helpPageUrl, "_blank");
                }
            });
        }

        // Render a breadcrumb for the nav brand
        let crumb = Components.Breadcrumb({
            el: this._el,
            items: [
                { text: "App Dashboard", href: Strings.DashboardUrl },
                { text: DataSource.DocSetItem.Title, href: "#", isActive: true }
            ]
        });
        
        // Update the breadcrumb divider to use a bootstrap icon
        let caret = caretRightFill(20, 20).outerHTML.replaceAll("\"", "'").replaceAll("<", "%3C").replaceAll(">", "%3E").replaceAll("\n", "").replaceAll("  ", " ").replace("currentColor", "%23fff");
        crumb.el.setAttribute("style", "--bs-breadcrumb-divider: url(\"data:image/svg+xml," + caret + "\");");
        
        // Enable the link back to the app dashboard
        crumb.el.querySelector(".breadcrumb-item a").classList.add("pe-auto");

        // Render the navigation
        let nav = Components.Navbar({
            el: this._el.querySelector("#app-nav"),
            brand: crumb.el,
            className: "bg-sharepoint rounded-top",
            type: Components.NavbarTypes.Primary,
            itemsEnd
        });
        
        // Adjust the nav alignment
        nav.el.querySelector("nav div.container-fluid").classList.add("pe-2");
        nav.el.querySelector("nav div.container-fluid").classList.add("ps-3");

        // Disable the link on the root nav brand
        nav.el.querySelector("nav div.container-fluid a.navbar-brand").classList.add("pe-none");
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