import { Documents, LoadingDialog, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
import { caretRightFill } from "gd-sprest-bs/build/icons/svgs/caretRightFill";
import { folderSymlink } from "gd-sprest-bs/build/icons/svgs/folderSymlink";
import { layoutTextWindow } from "gd-sprest-bs/build/icons/svgs/layoutTextWindow";
import { questionLg } from "gd-sprest-bs/build/icons/svgs/questionLg";
import { AppActions } from "./appActions";
import { AppConfig } from "./appCfg";
import { ButtonActions } from "./btnActions";
import * as Common from "./common";
import { DataSource } from "./ds";
import Strings from "./strings";

/**
 * Application Dashboard
 * Gives a more detailed overview of the application.
 */
export class AppDashboard {
    private _el: HTMLElement = null;
    private _elDashboard: HTMLElement = null;

    // Constructor
    constructor(el: HTMLElement, elDashboard?: HTMLElement, docSetId?: number) {
        // Initialize the component
        this._el = el;
        this._elDashboard = elDashboard;

        // Render the template
        this._el.classList.add("bs");
        this._el.innerHTML = `
            <div id="app-dashboard" class="row">
                <div id="app-nav" class="col-12"></div>
                <div id="app-info" class="col-9"></div>
                <div id="app-actions" class="col-3"></div>
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
        this.renderDocuments(docSetId);
    }

    // Redirects to the dashboard
    private redirectToDashboard() {
        // See if the parent element exists
        if (this._elDashboard) {
            // Hide the details
            this._el.style.display = "none";

            // Show the dashboard
            this._elDashboard.style.display = "";
        } else {
            // Redirect to the dashboard
            window.open(Strings.DashboardUrl, "_self");
        }
    }

    // Refreshes the dashboard
    private refresh() {
        // Show a loading dialog
        LoadingDialog.setHeader("Refreshing the Data");
        LoadingDialog.setBody("This will close after the data is loaded.");

        // Load the the document set information
        DataSource.loadDocSet().then(() => {
            // Render the info
            this.renderInfo();

            // Render the actions
            this.renderActions();

            // Hide the dialog
            LoadingDialog.hide();
        });
    }

    // Renders the dashboard
    private renderActions() {
        // Clear the actions
        let elActions = this._el.querySelector("#app-actions");
        while (elActions.firstChild) { elActions.removeChild(elActions.firstChild); }

        // Render a card
        Components.Card({
            el: elActions,
            body: [{
                title: "Actions",
                onRender: el => {
                    // Render the actions
                    new ButtonActions(el, DataSource.DocSetItem, () => {
                        // Refresh the dashboard
                        this.refresh();
                    })
                }
            }]
        });
    }

    // Renders the documents
    private renderDocuments(docSetId: number = DataSource.DocSetItemId) {
        // Render the documents
        new Documents({
            el: this._el.querySelector("#app-docs"),
            listName: Strings.Lists.Apps,
            docSetId,
            templatesUrl: AppConfig.Configuration.templatesLibraryUrl,
            webUrl: Strings.SourceUrl,
            onActionsRendered: (el, col, file) => {
                // See if this is a package file
                if (file.ServerRelativeUrl.endsWith(".sppkg")) {
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
                        // See if the app isn't approved or requires an assessment
                        if ([AppConfig.ApprovedStatus, AppConfig.TechReviewStatus, AppConfig.TestCasesStatus].indexOf(DataSource.DocSetItem.AppStatus) < 0) {
                            // Add the file
                            resolve(true);
                            return;
                        }

                        // Read the package
                        AppActions.readPackage(fileInfo.data).then(appInfo => {
                            let errorMessage = null;

                            // Deny if the product id doesn't match
                            if (appInfo.AppProductID != DataSource.DocSetItem.AppProductID) {
                                // Set the error message
                                errorMessage = "The app's product id doesn't match the current one.";
                            }
                            // Else, deny if the version is less than the current
                            else if (appInfo.AppVersion != DataSource.DocSetItem.AppVersion) {
                                // Compare the values
                                let appVersion = DataSource.DocSetItem.AppVersion.split('.');
                                let newAppVersion = appInfo.AppVersion.split('.');

                                // See if the new version is greater
                                for (let i = 0; i < appVersion.length; i++) {
                                    // See if this version is not greater than
                                    if (newAppVersion[0] < appVersion[0]) {
                                        // Set the error message
                                        errorMessage = "The app's version is less than the current one."
                                        break;
                                    }
                                }
                            }

                            // See if an error message exists
                            if (errorMessage) {
                                // Deny the file upload
                                resolve(false);

                                // Clear the modal
                                Modal.clear();

                                // Set the header
                                Modal.setHeader("Error Uploading File");

                                // Set the body
                                Modal.setBody("<p>The app package being uploaded has been denied for the following reason:</p><br/>" + errorMessage)

                                // Show the modal
                                Modal.show();
                            } else {
                                // Update the status and set it back to the testing status
                                DataSource.DocSetItem.update({
                                    AppStatus: AppConfig.TestCasesStatus
                                }).execute(() => {
                                    // See if the item is currently approved
                                    if (DataSource.DocSetItem.AppStatus == AppConfig.TestCasesStatus) {
                                        // Archive the file
                                        AppActions.archivePackage(DataSource.DocSetItem, () => {
                                            // Add the file
                                            resolve(true);
                                        });
                                    } else {
                                        // Add the file
                                        resolve(true);
                                    }
                                });
                            }
                        });
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

        // Render a card
        Components.CardGroup({
            el: elInfo,
            cards: [
                {
                    body: [{
                        onRender: el => {
                            // Render the properties
                            Components.ListForm.renderDisplayForm({
                                info: DataSource.DocSetInfo,
                                el,
                                includeFields: [
                                    "AppSponser",
                                    "AppStatus",
                                    "AppDevelopers",
                                    "AppProductID",
                                    "AppIsClientSideSolution"
                                ],
                                onControlRendering: (ctrl, field) => {
                                    // See if this is the status field and the item is rejected
                                    if (field.InternalName == "AppStatus") {
                                        // Set the status value
                                        ctrl.value = Common.appStatus(DataSource.DocSetItem);
                                    }
                                }
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
                                    "AppVersion",
                                    "AppSharePointMinVersion",
                                    "AppIsDomainIsolated",
                                    "AppSkipFeatureDeployment",
                                    "AppAPIPermissions"
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
                    // Show the info and actions
                    this._el.querySelector("#app-info").classList.remove("d-none");
                    this._el.querySelector("#app-actions").classList.remove("d-none");
                    elNavInfo.classList.add("d-none");

                    // Hide the documents
                    this._el.querySelector("#app-docs").classList.add("d-none");
                    elNavDocs.classList.remove("d-none");

                    crumb.remove();
                    crumb.remove();
                    crumb.add({
                        text: DataSource.DocSetItem.Title,
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
                    this._el.querySelector("#app-actions").classList.add("d-none");
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
                            text: DataSource.DocSetItem.Title, className: "pe-auto", href: "#", onClick: () => {
                                // Show the info
                                this._el.querySelector("#app-info").classList.remove("d-none");
                                this._el.querySelector("#app-actions").classList.remove("d-none");
                                elNavInfo.classList.add("d-none");

                                // Hide the documents
                                this._el.querySelector("#app-docs").classList.add("d-none");
                                elNavDocs.classList.remove("d-none");

                                crumb.remove();
                                crumb.remove();
                                crumb.add({
                                    text: DataSource.DocSetItem.Title,
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
                //{ text: "App Dashboard", href: Strings.DashboardUrl, className: "pe-auto" },
                { text: DataSource.DocSetItem.Title, href: "#", isActive: true }
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