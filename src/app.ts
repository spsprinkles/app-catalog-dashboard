import { Documents, LoadingDialog } from "dattatable";
import { Components, Types } from "gd-sprest-bs";
import { appIndicator } from "gd-sprest-bs/build/icons/svgs/appIndicator";
import { caretRightFill } from "gd-sprest-bs/build/icons/svgs/caretRightFill";
import { chatSquareDots } from "gd-sprest-bs/build/icons/svgs/chatSquareDots";
import { folderSymlink } from "gd-sprest-bs/build/icons/svgs/folderSymlink";
import { layoutTextWindow } from "gd-sprest-bs/build/icons/svgs/layoutTextWindow";
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
                <div id="app-docs" class="col-12 d-none"></div>
            </div>
        `.trim();

        // Render the navigation
        this.renderNav();

        // Render the info
        this.renderInfo();

        // Render the actions
        this.renderActions();

        // Render the documents
        this.renderDocuments();
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

    // Renders the documents
    private renderDocuments() {
        // Render the documents
        new Documents({
            el: this._el.querySelector("#app-docs"),
            listName: Strings.Lists.Apps,
            docSetId: DataSource.DocSetItemId,
            templatesUrl: DataSource.Configuration.templatesLibraryUrl,
            onActionsRendered: (el, col, file) => {
                // Hide the properties button
                el.querySelector(".btn-actions-properties").parentElement.classList.add("d-none");

                // See if this is a package file
                if (file.ServerRelativeUrl.endsWith(".sppkg")) {
                    // Disable the view & delete button
                    el.querySelectorAll(".btn-actions-view, .btn-actions-delete").forEach((el: HTMLElement) => {
                        el.setAttribute("disabled", "");
                        el.parentElement.classList.add("pe-none");
                    });
                }
            },
            onNavigationRendered: (nav) => {
                // Fix the Templates dropdown offset value
                (nav.el.querySelector("ul:last-child li:first-child a") as any)._tippy.setProps({ offset: [0, 4] });
            },
            onNavigationRendering: props => {
                // Clear the subNav brand
                props.brand = "";
            },
            table: {
                columns: [
                    {
                        name: "Type",
                        title: "Type",
                    },
                    {
                        name: "Name",
                        title: "Name"
                    },
                    {
                        name: "FileSize",
                        title: "File Size"
                    },
                    {
                        name: "Created",
                        title: "Created"
                    },
                    {
                        name: "Author",
                        title: "Created By"
                    },
                    {
                        name: "Modified",
                        title: "Modified"
                    },
                    {
                        name: "ModifiedBy",
                        title: "Modified By"
                    },
                    {
                        className: "text-end text-nowrap",
                        name: "Actions",
                        title: ""
                    }
                ]
            }
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
        let elNavDocs: HTMLElement = null;
        let elNavInfo: HTMLElement = null;
        let itemsEnd: Components.INavbarItem[] = [
            {
                className: "btn-outline-light ms-2 ps-2 pt-1",
                classNameItem: "d-none",
                iconClassName: "me-1",
                iconSize: 24,
                iconType: layoutTextWindow,
                isButton: true,
                text: "App Details",
                onRender: el => { elNavInfo = el; },
                onClick: () => {
                    // Show the info
                    this._el.querySelector("#app-info").classList.remove("d-none");
                    elNavInfo.classList.add("d-none");

                    // Hide the documents
                    this._el.querySelector("#app-docs").classList.add("d-none");
                    elNavDocs.classList.remove("d-none");

                    // Set the bread crumb
                    crumb.setItems([{
                        text: DataSource.DocSetItem.Title,
                        href: "#",
                        isActive: true
                    }]);
                }
            },
            {
                className: "btn-outline-light ms-2 ps-2 pt-1",
                iconClassName: "me-1",
                iconSize: 24,
                iconType: folderSymlink,
                isButton: true,
                text: "Documents",
                onRender: el => { elNavDocs = el; },
                onClick: (item) => {
                    // Show the documents
                    this._el.querySelector("#app-docs").classList.remove("d-none");
                    elNavDocs.classList.add("d-none");

                    // Hide the info
                    this._el.querySelector("#app-info").classList.add("d-none");
                    elNavInfo.classList.remove("d-none");

                    crumb.setItems([
                        { text: "App Dashboard", href: Strings.DashboardUrl, className: "pe-auto" },
                        {
                            text: DataSource.DocSetItem.Title, className: "pe-auto", href: "#", onClick: () => {
                                // Show the info
                                this._el.querySelector("#app-info").classList.remove("d-none");
                                elNavInfo.classList.add("d-none");

                                // Hide the documents
                                this._el.querySelector("#app-docs").classList.add("d-none");
                                elNavDocs.classList.remove("d-none");

                                // Set the bread crumb
                                crumb.setItems([{
                                    text: DataSource.DocSetItem.Title,
                                    href: "#",
                                    isActive: true
                                }]);
                            }
                        },
                        { text: item.text, href: "#", isActive: true }
                    ]);
                }
            }
        ];

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
                { text: "App Dashboard", href: Strings.DashboardUrl, className: "pe-auto" },
                { text: DataSource.DocSetItem.Title, href: "#", isActive: true }
            ]
        });

        // Update the breadcrumb divider to use a bootstrap icon
        crumb.el.setAttribute("style", "--bs-breadcrumb-divider: " + Common.generateEmbeddedSVG(caretRightFill(18, 18)).replace("currentColor", "%23fff"));

        // Render the navigation
        let nav = Components.Navbar({
            el: this._el.querySelector("#app-nav"),
            brand: crumb.el,
            className: "bg-sharepoint rounded-top",
            type: Components.NavbarTypes.Primary,
            itemsEnd
        });

        // Adjust the nav alignment
        nav.el.querySelector("nav div.container-fluid").classList.add("pe-75");
        nav.el.querySelector("nav div.container-fluid").classList.add("ps-3");

        // Disable the link on the root nav brand
        nav.el.querySelector("nav div.container-fluid a.navbar-brand").classList.add("pe-none");
    }
}