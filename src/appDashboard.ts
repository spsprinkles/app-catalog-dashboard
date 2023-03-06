import { Documents, LoadingDialog } from "dattatable";
import { Components, Types } from "gd-sprest-bs";
import { caretRightFill } from "gd-sprest-bs/build/icons/svgs/caretRightFill";
import { folderSymlink } from "gd-sprest-bs/build/icons/svgs/folderSymlink";
import { layoutTextWindow } from "gd-sprest-bs/build/icons/svgs/layoutTextWindow";
import { questionLg } from "gd-sprest-bs/build/icons/svgs/questionLg";
import { AppActions } from "./appActions";
import { AppConfig } from "./appCfg";
import { AppView } from "./appView";
import { ButtonActions } from "./btnActions";
import * as Common from "./common";
import { DataSource } from "./ds";
import Strings from "./strings";

/**
 * Application Dashboard
 * Gives a more detailed overview of the application.
 */
export class AppDashboard {
    private _docSetId: number = null;
    private _el: HTMLElement = null;
    private _elDashboard: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement, elDashboard: HTMLElement, docSetId: number) {
        // Initialize the component
        this._docSetId = docSetId || DataSource.AppItem.Id;
        this._el = el;
        this._elDashboard = elDashboard;

        // Render the template
        this._el.classList.add("bs");
        this._el.innerHTML = `
            <div id="app-dashboard" class="row">
                <div id="app-status-info" class="col-12"></div>
                <div id="app-nav" class="col-12"></div>
                <div id="app-error" class="col-12"></div>
                <div id="app-info" class="col-12"></div>
                <div id="app-docs" class="col-12 d-none"></div>
            </div>
        `.trim();

        // Render the navigation
        this.renderNav();

        // Render the alert
        this.renderAlertStatus();
        this.renderAlertError();

        // Render the info
        this.renderInfo();

        // Render the actions
        this.renderActions();

        // Render the documents
        this.renderDocuments();
    }

    // Redirects to the dashboard
    private redirectToDashboard() {
        // See if the dashboard exists
        if (this._elDashboard && this._elDashboard.firstChild) {
            // Hide the details
            this._el.classList.add("d-none");

            // Clear the dashboard
            while (this._elDashboard.firstChild) { this._elDashboard.removeChild(this._elDashboard.firstChild); }

            // Show the dashboard
            this._elDashboard.classList.remove("d-none");

            // Render the app view
            new AppView(this._elDashboard, this._el);
        } else {
            // Get the page name
            let names = window.location.pathname.split('/');
            let pageName = names[names.length - 1].toLowerCase();

            // See if this is the document set home page
            if (pageName == "docsethomepage.aspx") {
                // Open the dashboard in a new tab
                window.open(AppConfig.Configuration.dashboardUrl, "_blank");
            }
            else {
                // Show a loading dialog
                LoadingDialog.setHeader("Loading the Applications");
                LoadingDialog.setBody("This will close after the data is loaded.");
                LoadingDialog.show();

                // Load the data
                DataSource.refresh().then(() => {
                    // Hide the details
                    this._el.classList.add("d-none");

                    // Render the app view
                    new AppView(this._elDashboard, this._el);

                    // Hide the dialog
                    LoadingDialog.hide();
                });
            }
        }
    }

    // Refreshes the dashboard
    private refresh() {
        // Show a loading dialog
        LoadingDialog.setHeader("Refreshing the Data");
        LoadingDialog.setBody("This will close after the data is loaded.");
        LoadingDialog.show();

        // Refresh the data
        DataSource.refresh(this._docSetId).then(() => {
            // Render the alerts
            this.renderAlertStatus();
            this.renderAlertError();

            // Render the info
            this.renderInfo();

            // Render the actions
            this.renderActions();

            // Render the documents
            this.renderDocuments();

            // Hide the dialog
            LoadingDialog.hide();
        });
    }

    // Renders the dashboard
    private renderActions() {
        // Define the card group parent
        let cardGroup = this._el.querySelector("#app-info > div.card-group");

        // Clear the actions
        let elActions = cardGroup.querySelector("div.app-actions");
        while (elActions) { elActions.parentElement.removeChild(elActions); }

        // Render a card
        Components.Card({
            el: cardGroup,
            className: "app-actions",
            body: [{
                title: "Actions",
                onRender: el => {
                    // Render the actions
                    new ButtonActions(el, () => {
                        // Refresh the dashboard
                        this.refresh();
                    });
                }
            }]
        });
    }

    // Renders the status information from the configuration
    private renderAlertStatus() {
        // Clear the element
        let elAlert = this._el.querySelector("#app-status-info");
        while (elAlert.firstChild) { elAlert.removeChild(elAlert.firstChild); }

        // See if an alert exists
        let cfg = AppConfig.Status[DataSource.AppItem.AppStatus];
        if (cfg && cfg.alert) {
            // Show the element
            elAlert.classList.remove("d-none");

            // Render an alert
            Components.Alert({
                el: elAlert,
                className: "m-0 rounded-0",
                header: cfg.alert.header || "",
                content: cfg.alert.content || "",
                isDismissible: false,
                type: Components.AlertTypes.Info
            });
        }
        else {
            // Hide the element
            elAlert.classList.add("d-none");
        }
    }

    // Renders the error in an alert
    private renderAlertError() {
        // Clear the element
        let elAlert = this._el.querySelector("#app-error");
        while (elAlert.firstChild) { elAlert.removeChild(elAlert.firstChild); }

        // See if an alert exists
        if (DataSource.AppItem.AppIsRejected && DataSource.AppItem.AppComments) {
            // Show the element
            elAlert.classList.remove("d-none");

            // Render an alert
            Components.Alert({
                el: elAlert,
                className: "m-0 rounded-0",
                header: "Request Rejected",
                content: DataSource.AppItem.AppComments,
                type: Components.AlertTypes.Danger
            });
        }
        else {
            // Hide the element
            elAlert.classList.add("d-none");
        }

        // See if the app is deployed and has an error message
        let errorMessage = DataSource.AppCatalogSiteItem ? DataSource.AppCatalogSiteItem.ErrorMessage : null;
        errorMessage = (DataSource.AppCatalogItem ? DataSource.AppCatalogItem.AppPackageErrorMessage : null) || errorMessage;
        errorMessage = (DataSource.AppItem ? DataSource.AppItem.AppPackageErrorMessage : null) || errorMessage;
        if (errorMessage && errorMessage != "No errors.") {
            // Show the element
            elAlert.classList.remove("d-none");

            // Render an alert
            Components.Alert({
                el: elAlert,
                className: "m-0 rounded-0",
                header: "App Deployment Error",
                content: errorMessage,
                type: Components.AlertTypes.Danger
            });
        }
    }

    // Renders the documents
    private renderDocuments() {
        // Clear the docs element
        let elDocs = this._el.querySelector("#app-docs") as HTMLElement;
        while (elDocs.firstChild) { elDocs.removeChild(elDocs.firstChild); }

        // Render the documents
        new Documents({
            el: elDocs,
            listName: Strings.Lists.Apps,
            docSetId: this._docSetId,
            templatesUrl: AppConfig.Configuration.templatesLibraryUrl,
            webUrl: Strings.SourceUrl,
            onActionsRendered: (el, col, file) => {
                // See if this is a package file and not archived
                let fileUrl = file.ServerRelativeUrl.toLowerCase();
                if (fileUrl.indexOf("/archive/") < 0 && fileUrl.endsWith(".sppkg")) {
                    // Disable the view & delete button
                    el.querySelectorAll(".btn-actions-view, .btn-actions-delete").forEach((el: HTMLElement) => {
                        el.setAttribute("disabled", "");
                        el.parentElement.classList.add("pe-none");
                    });
                }
            },
            onFileAdding: (fileInfo) => {
                // Return a promise
                return new Promise(resolve => {
                    // See if this is an spfx package
                    if (fileInfo.name.toLowerCase().endsWith(".sppkg")) {
                        // Update the app package
                        AppActions.uploadAppPackage(fileInfo).then(() => {
                            // Refresh the dashboard
                            this.refresh();
                        })
                    } else {
                        // Resolve the request
                        resolve(true);
                    }
                });
            },
            onNavigationRendered: (nav) => {
                // Update the upload button tooltip
                (nav.el.querySelector("#navbar_content ul:last-child li:last-child button") as any)._tippy.setContent("Upload Document");
            },
            onNavigationRendering: props => {
                // Clear the subNav brand
                props.brand = "";
            }
        });
    }

    // Renders the information
    private renderInfo() {
        // Clear the information
        let elInfo = this._el.querySelector("#app-info");
        while (elInfo.firstChild) { elInfo.removeChild(elInfo.firstChild); }

        // Update the fields after rendering them
        let postUpdateFields = (ctrl: Components.IFormControl, field: Types.SP.Field) => {
            // See if this is the API permission field
            if (field && field.InternalName == "AppAPIPermissions") {
                // Check if a value exists
                if (ctrl.getValue()) {
                    // Select the textarea element
                    let textarea = ctrl.el.querySelector("textarea");
                    // Shrink the text size
                    textarea.style.fontSize = "0.75rem";
                    // Extend the textarea number of rows
                    textarea.setAttribute("rows", "4");
                }
            }
        }

        // Render a card
        Components.CardGroup({
            el: elInfo,
            cards: [
                {
                    body: [{
                        onRender: el => {
                            // Render the display form
                            Components.ListForm.renderDisplayForm({
                                info: DataSource.AppFormInfo,
                                el,
                                includeFields: AppConfig.Configuration.appDetails && AppConfig.Configuration.appDetails.left ? AppConfig.Configuration.appDetails.left : [
                                    "AppIsTenant",
                                    "AppStatus",
                                    "AppDevelopers",
                                    "AppSponsor",
                                    "AppProductID",
                                    "AppIsClientSideSolution"
                                ],
                                onControlRendered: (ctrl, field) => {
                                    // Update the fields
                                    postUpdateFields(ctrl, field);
                                }
                            });
                        }
                    }]
                },
                {
                    body: [{
                        onRender: el => {
                            // Render the display form
                            Components.ListForm.renderDisplayForm({
                                info: DataSource.AppFormInfo,
                                el,
                                includeFields: AppConfig.Configuration.appDetails && AppConfig.Configuration.appDetails.right ? AppConfig.Configuration.appDetails.right : [
                                    "AppVersion",
                                    "AppSharePointMinVersion",
                                    "AppIsDomainIsolated",
                                    "AppSkipFeatureDeployment",
                                    "AppAPIPermissions"
                                ],
                                onControlRendered: (ctrl, field) => {
                                    // Update the fields
                                    postUpdateFields(ctrl, field);
                                }
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
                    // Show the info and actions
                    this._el.querySelector("#app-info").classList.remove("d-none");
                    elNavInfo.classList.add("d-none");

                    // Hide the documents
                    this._el.querySelector("#app-docs").classList.add("d-none");
                    elNavDocs.classList.remove("d-none");

                    crumb.remove();
                    crumb.remove();
                    crumb.add({
                        text: DataSource.AppItem.Title,
                        href: "#",
                        isActive: true
                    });
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

                    // Hide the info and actions
                    this._el.querySelector("#app-info").classList.add("d-none");
                    elNavInfo.classList.remove("d-none");

                    // Hide the actions
                    crumb.setItems([
                        {
                            text: "App Dashboard",
                            className: "pe-auto",
                            href: "#",
                            onClick: () => {
                                // Redirect to the app dashboard
                                this.redirectToDashboard();
                            }
                        },
                        {
                            text: DataSource.AppItem.Title, className: "pe-auto", href: "#", onClick: () => {
                                // Show the info
                                this._el.querySelector("#app-info").classList.remove("d-none");
                                elNavInfo.classList.add("d-none");

                                // Hide the documents
                                this._el.querySelector("#app-docs").classList.add("d-none");
                                elNavDocs.classList.remove("d-none");

                                crumb.remove();
                                crumb.remove();
                                crumb.add({
                                    text: DataSource.AppItem.Title,
                                    href: "#",
                                    isActive: true
                                });
                            }
                        },
                        { text: item.text, href: "#", isActive: true }
                    ]);
                }
            }
        ];

        // See if the help url exists
        if (AppConfig.Configuration.helpPageUrl) {
            // Add the item
            itemsEnd.push({
                className: "btn-outline-light ms-2 ps-1 pt-1",
                iconSize: 24,
                iconType: questionLg,
                isButton: true,
                text: "Help",
                onClick: () => {
                    // Display in a new tab
                    window.open(AppConfig.Configuration.helpPageUrl, "_blank");
                }
            });
        }

        // Render a breadcrumb for the nav brand
        let crumb = Components.Breadcrumb({
            el: this._el,
            items: [
                {
                    text: "App Dashboard",
                    className: "pe-auto",
                    href: "#",
                    onClick: () => {
                        // Redirect to the app dashboard
                        this.redirectToDashboard();
                    }
                },
                //{ text: "App Dashboard", href: AppConfig.Configuration.dashboardUrl, className: "pe-auto" },
                { text: DataSource.AppItem.Title, href: "#", isActive: true }
            ]
        });

        // Update the breadcrumb divider to use a bootstrap icon
        crumb.el.setAttribute("style", "--bs-breadcrumb-divider: " + Common.generateEmbeddedSVG(caretRightFill(18, 18)).replace("currentColor", "%23fff"));

        // Render the navigation
        let nav = Components.Navbar({
            el: this._el.querySelector("#app-nav"),
            brand: crumb.el,
            className: "navbar-expand bg-sharepoint rounded-top",
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