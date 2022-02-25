import { Documents, LoadingDialog } from "dattatable";
import { Components, ContextInfo, Web } from "gd-sprest-bs";
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
import { DataSource, IStatus } from "./ds";
import Strings from "./strings";

/**
 * Main Application
 */
export class App {
    private _el: HTMLElement = null;
    private _elDashboard: HTMLElement = null;
    private _forms: AppForms = null;

    // Constructor
    constructor(el: HTMLElement, elDashboard?: HTMLElement, docSetId?: number) {
        // Set the global variables
        this._el = el;
        this._elDashboard = elDashboard;
        this._forms = new AppForms();

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
        let elActionButtons = null;
        Components.Card({
            el: elActions,
            body: [{
                title: "Actions",
                onRender: el => {
                    // Set the element
                    elActionButtons = el;
                }
            }]
        });

        // Ensure the actions have been defined
        let actions = DataSource.Configuration.actions;
        let status: IStatus = DataSource.Status[DataSource.DocSetItem.AppStatus] || {} as any;
        let btnActions = actions ? actions[DataSource.DocSetItem.AppIsRejected ? "Rejected" : DataSource.DocSetItem.AppStatus] : null;
        if (btnActions == null || btnActions.length == 0) { return; }

        // Determine if the user can edit items
        let canEdit = Common.canEdit(DataSource.DocSetItem);

        // Parse the button actions
        for (let i = 0; i < btnActions.length; i++) {
            // Render the action button
            switch (btnActions[i]) {
                // Approve/Reject
                case "ApproveReject":
                    // Ensure this is an approver
                    if (DataSource.IsApprover) {
                        // See if this is the sponser state
                        let isSponserApproval = status.name == "Pending Sponser Approval";

                        // Render the approval button
                        Components.Tooltip({
                            el: elActionButtons,
                            content: "Approves the application.",
                            btnProps: {
                                text: "Approve",
                                iconClassName: "me-1",
                                iconSize: 20,
                                iconType: chatSquareDots,
                                isSmall: true,
                                isDisabled: isSponserApproval ? DataSource.DocSetItem.AppSponserId != ContextInfo.userId : false,
                                type: Components.ButtonTypes.OutlineSuccess,
                                onClick: () => {
                                    // Display the approval form
                                    this._forms.approve(DataSource.DocSetItem, () => {
                                        // Refresh the dashboard
                                        this.refresh();
                                    });
                                }
                            }
                        });

                        // Render the reject button
                        Components.Tooltip({
                            el: elActionButtons,
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
                                        this.refresh();
                                    });
                                }
                            }
                        });
                    }
                    break;

                // Delete
                case "Delete":
                    // Ensure this is an approver
                    if (DataSource.IsApprover) {
                        // Render the button
                        Components.Tooltip({
                            el: elActionButtons,
                            content: "Deletes the app.",
                            btnProps: {
                                text: "Delete App/Solution",
                                iconClassName: "me-1",
                                iconSize: 20,
                                iconType: trash,
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
                    break;

                // Delete Test Site
                case "DeleteTestSite":
                    // TODO
                    break;

                // Deploy Site Catalog
                case "DeploySiteCatalog":
                    // TODO
                    break;

                // Deploy Tenant Catalog
                case "DeployTenantCatalog":
                    // See if the app is deployed
                    if (DataSource.DocSetTenantApp && DataSource.DocSetTenantApp.Deployed) {
                        // Render the retract button
                        Components.Tooltip({
                            el: elActionButtons,
                            content: "Retracts the solution from the tenant app catalog.",
                            btnProps: {
                                text: "Retract from Tenant",
                                iconClassName: "me-1",
                                iconSize: 20,
                                //iconType: trash,
                                isDisabled: DataSource.IsTenantAppCatalogOwner,
                                isSmall: true,
                                type: Components.ButtonTypes.OutlineDanger,
                                onClick: () => {
                                    // Retract the app
                                    this._forms.retract(DataSource.DocSetItem, true, false, () => {
                                        // Refresh the page
                                        this.refresh();
                                    });
                                }
                            }
                        });

                        // Load the context of the app catalog
                        ContextInfo.getWeb(DataSource.Configuration.tenantAppCatalogUrl).execute(context => {
                            let requestDigest = context.GetContextWebInformation.FormDigestValue;
                            let web = Web(DataSource.Configuration.tenantAppCatalogUrl, { requestDigest });

                            // Ensure this app can be deployed to the tenant
                            web.TenantAppCatalog().solutionContainsTeamsComponent(DataSource.DocSetTenantApp.ID).execute((resp: any) => {
                                // See if we can deploy this app to teams
                                if (resp.SolutionContainsTeamsComponent) {
                                    // Render the deploy to teams button
                                    Components.Tooltip({
                                        el: elActionButtons,
                                        content: "Deploys the solution to Teams.",
                                        btnProps: {
                                            text: "Deploy to Teams",
                                            iconClassName: "me-1",
                                            iconSize: 20,
                                            //iconType: trash,
                                            isDisabled: DataSource.IsTenantAppCatalogOwner,
                                            isSmall: true,
                                            type: Components.ButtonTypes.OutlineWarning,
                                            onClick: () => {
                                                // Deploy the app
                                                this._forms.deployToTeams(DataSource.DocSetItem, () => {
                                                    // Refresh the page
                                                    this.refresh();
                                                });
                                            }
                                        }
                                    });
                                }
                            });
                        });
                    } else {
                        // Render the deploy button
                        Components.Tooltip({
                            el: elActionButtons,
                            content: "Deploys the application to the tenant app catalog.",
                            btnProps: {
                                text: "Deploy",
                                iconClassName: "me-1",
                                iconSize: 20,
                                //iconType: trash,
                                isSmall: true,
                                type: Components.ButtonTypes.OutlineSuccess,
                                onClick: () => {
                                    // Deploy the app
                                    this._forms.deploy(DataSource.DocSetItem, true, () => {
                                        // Refresh the page
                                        this.refresh();
                                    });
                                }
                            }
                        });
                    }
                    break;

                // Edit
                case "Edit":
                    // Render the button
                    Components.Tooltip({
                        el: elActionButtons,
                        content: "Displays the edit form to update the app properties.",
                        btnProps: {
                            text: "Edit Properties",
                            iconClassName: "me-1",
                            iconSize: 20,
                            iconType: pencilSquare,
                            isDisabled: !canEdit,
                            isSmall: true,
                            type: Components.ButtonTypes.OutlineSecondary,
                            onClick: () => {
                                // Display the edit form
                                this._forms.edit(DataSource.DocSetItem.Id, () => {
                                    // Refresh the page
                                    this.refresh();
                                });
                            }
                        }
                    });
                    break;

                // Edit Tech Review
                case "EditTechReview":
                    // Render the button
                    Components.Tooltip({
                        el: elActionButtons,
                        content: "Edits the technical review for this app.",
                        btnProps: {
                            text: "Edit Technical Review",
                            iconClassName: "me-1",
                            iconSize: 20,
                            iconType: chatSquareDots,
                            isSmall: true,
                            type: Components.ButtonTypes.OutlinePrimary,
                            onClick: () => {
                                // Display the test cases
                                this._forms.editTechReview(DataSource.DocSetItem, () => {
                                    // Refresh the page
                                    this.refresh();
                                });
                            }
                        }
                    });
                    break;

                // Edit Test Cases
                case "EditTestCases":
                    // Render the button
                    Components.Tooltip({
                        el: elActionButtons,
                        content: "Edit the test cases for this app.",
                        btnProps: {
                            text: "Edit Test Cases",
                            iconClassName: "me-1",
                            iconSize: 20,
                            iconType: chatSquareDots,
                            isDisabled: !canEdit,
                            isSmall: true,
                            type: Components.ButtonTypes.OutlinePrimary,
                            onClick: () => {
                                // Display the review form
                                this._forms.editTestCases(DataSource.DocSetItem, () => {
                                    // Refresh the page
                                    this.refresh();
                                });
                            }
                        }
                    });
                    break;

                // Submit
                case "Submit":
                    // Render the button
                    Components.Tooltip({
                        el: elActionButtons,
                        content: "Submits the app for approval",
                        btnProps: {
                            text: "Submit for Approval",
                            iconClassName: "me-1",
                            iconSize: 20,
                            iconType: appIndicator,
                            isDisabled: !canEdit || (status.lastStep),
                            isSmall: true,
                            type: Components.ButtonTypes.OutlinePrimary,
                            onClick: () => {
                                // Display the submit form
                                this._forms.submit(DataSource.DocSetItem, () => {
                                    // Refresh the page
                                    this.refresh();
                                });
                            }
                        }
                    });
                    break;

                // Test Site
                case "TestSite":
                    // See if a test site exists
                    DataSource.loadTestSite(DataSource.DocSetItem).then(
                        // Test site exists
                        web => {
                            // Render the view button
                            Components.Tooltip({
                                el: elActionButtons,
                                content: "Opens the test site in a new tab.",
                                btnProps: {
                                    text: "View Test Site",
                                    iconClassName: "me-1",
                                    iconSize: 20,
                                    iconType: chatSquareDots,
                                    isSmall: true,
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // Open the test site in a new tab
                                        window.open(web.Url, "_blank");
                                    }
                                }
                            });
                        },

                        // Test site doesn't exist
                        () => {
                            // Render the create button
                            Components.Tooltip({
                                el: elActionButtons,
                                content: "Creates the test site for the app.",
                                btnProps: {
                                    text: "Create Test Site",
                                    iconClassName: "me-1",
                                    iconSize: 20,
                                    iconType: chatSquareDots,
                                    isDisabled: !canEdit,
                                    isSmall: true,
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // Display the create test site form
                                        this._forms.createTestSite(DataSource.DocSetItem, () => {
                                            // Refresh the page
                                            this.refresh();
                                        });
                                    }
                                }
                            });
                        }
                    );
                    break;

                // View
                case "View":
                    // Render the button
                    Components.Tooltip({
                        el: elActionButtons,
                        content: "Displays the the app properties.",
                        btnProps: {
                            text: "View Properties",
                            iconClassName: "me-1",
                            iconSize: 20,
                            iconType: pencilSquare,
                            isSmall: true,
                            type: Components.ButtonTypes.OutlineSecondary,
                            onClick: () => {
                                // Display the edit form
                                this._forms.display(DataSource.DocSetItem.Id);
                            }
                        }
                    });
                    break;

                // View Tech Review
                case "ViewTechReview":
                    // Render the button
                    Components.Tooltip({
                        el: elActionButtons,
                        content: "Views the technical review for this app.",
                        btnProps: {
                            text: "View Technical Review",
                            iconClassName: "me-1",
                            iconSize: 20,
                            iconType: chatSquareDots,
                            isSmall: true,
                            type: Components.ButtonTypes.OutlinePrimary,
                            onClick: () => {
                                // Display the test cases
                                this._forms.displayTechReview(DataSource.DocSetItem);
                            }
                        }
                    });
                    break;

                // View Test Cases
                case "ViewTestCases":
                    // Render the button
                    Components.Tooltip({
                        el: elActionButtons,
                        content: "View the test cases for this app.",
                        btnProps: {
                            text: "View Test Cases",
                            iconClassName: "me-1",
                            iconSize: 20,
                            iconType: chatSquareDots,
                            isDisabled: !canEdit,
                            isSmall: true,
                            type: Components.ButtonTypes.OutlinePrimary,
                            onClick: () => {
                                // Display the review form
                                this._forms.displayTestCases(DataSource.DocSetItem);
                            }
                        }
                    });
                    break;
            }
        }
    }

    // Renders the documents
    private renderDocuments(docSetId: number = DataSource.DocSetItemId) {
        // Render the documents
        new Documents({
            el: this._el.querySelector("#app-docs"),
            listName: Strings.Lists.Apps,
            docSetId,
            templatesUrl: DataSource.Configuration.templatesLibraryUrl,
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
                                    "FileLeafRef",
                                    "AppStatus",
                                    "AppDevelopers",
                                    "AppProductID",
                                    "AppIsClientSideSolution"
                                ],
                                onControlRendering: (ctrl, field) => {
                                    // See if this is the status field and the item is rejected
                                    if (field.InternalName == "AppStatus" && DataSource.DocSetItem.AppIsRejected) {
                                        // Set the status value
                                        ctrl.value = "Rejected";
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
                    // Show the info
                    this._el.querySelector("#app-info").classList.remove("d-none");
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

                    // Hide the info
                    this._el.querySelector("#app-info").classList.add("d-none");
                    elNavInfo.classList.remove("d-none");

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