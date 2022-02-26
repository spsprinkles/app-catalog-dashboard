import { Components, ContextInfo, Web } from "gd-sprest-bs";
import { appIndicator } from "gd-sprest-bs/build/icons/svgs/appIndicator";
import { chatSquareDots } from "gd-sprest-bs/build/icons/svgs/chatSquareDots";
import { pencilSquare } from "gd-sprest-bs/build/icons/svgs/pencilSquare";
import { trash } from "gd-sprest-bs/build/icons/svgs/trash";
import { AppConfig, IStatus } from "./appCfg";
import * as Common from "./common";
import { DataSource, IAppItem } from "./ds";
import { AppForms } from "./itemForms";
import Strings from "./strings";

/**
 * Actions
 * Renders the actions available, based on the status.
 */
export class Actions {
    private _el: HTMLElement = null;
    private _forms: AppForms = null;
    private _item: IAppItem = null;
    private _onUpdate: () => void = null;

    // Constructor
    constructor(el: HTMLElement, item: IAppItem, onUpdate: () => void) {
        // Initialize the component
        this._el = el;
        this._forms = new AppForms();
        this._item = item;
        this._onUpdate = onUpdate;

        // Render the actions
        this.render();
    }

    // Status
    private get Status(): IStatus { return AppConfig.Status[this._item.AppStatus] || {} as any; }

    // Determines if the user is an approver
    private isApprover() {
        // See if there is an approver setting for this status
        let approvers = this.Status.approval || [];
        if (approvers.length == 0) { return false; }

        // Parse the approver
        for (let i = 0; i < approvers.length; i++) {
            // See which approver type it is
            switch (approvers[i]) {
                // Approver
                case "Approver":
                    // See if this user is an approver
                    return DataSource.IsApprover;

                // Developer
                case "Developer":
                    // See if this user is a developer
                    return DataSource.IsDeveloper;

                // Sponser
                case "Sponser":
                    // See if this user is a sponser
                    return this._item.AppSponserId == ContextInfo.userId;
            }
        }
    }

    // Renders the actions
    render() {
        // Ensure actions exist
        let btnActions = this.Status.actions;
        if (btnActions == null || btnActions.length == 0) { return; }

        // Determine if the user can edit items
        let canEdit = Common.canEdit(this._item);

        // Render the tooltip group
        let tooltips = Components.TooltipGroup({
            el: this._el,
            isVertical: true
        });

        // Parse the button actions
        for (let i = 0; i < btnActions.length; i++) {
            // Render the action button
            switch (btnActions[i]) {
                // Approve/Reject
                case "ApproveReject":
                    // Ensure this user can approve this item
                    if (this.isApprover()) {
                        // Render the approval button
                        tooltips.add({
                            content: "Approves the application.",
                            btnProps: {
                                text: "Approve",
                                iconClassName: "me-1",
                                iconSize: 20,
                                iconType: chatSquareDots,
                                isSmall: true,
                                type: Components.ButtonTypes.OutlineSuccess,
                                onClick: () => {
                                    // Display the approval form
                                    this._forms.approve(this._item, () => {
                                        // Call the update event
                                        this._onUpdate();
                                    });
                                }
                            }
                        });

                        // Render the reject button
                        tooltips.add({
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
                                    this._forms.reject(this._item, () => {
                                        // Call the update event
                                        this._onUpdate();
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
                        tooltips.add({
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
                                    this._forms.delete(this._item, () => {
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
                    // See if this is an approver
                    if (DataSource.IsSiteAppCatalogOwner) {
                        // See if a test site exists
                        DataSource.loadTestSite(this._item).then(
                            // Test site exists
                            () => {
                                // Render the button
                                tooltips.add({
                                    content: "Deletes the test site.",
                                    btnProps: {
                                        text: "Delete Test Site",
                                        iconClassName: "me-1",
                                        iconSize: 20,
                                        iconType: trash,
                                        isSmall: true,
                                        type: Components.ButtonTypes.OutlineDanger,
                                        onClick: () => {
                                            // Display the delete site form
                                            this._forms.deleteSite(this._item, () => {
                                                // Redirect to the dashboard
                                                window.open(Strings.DashboardUrl, "_self");
                                            });
                                        }
                                    }
                                });
                            },

                            // Test site doesn't exist
                            () => { }
                        );
                    }
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
                        tooltips.add({
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
                                    this._forms.retract(this._item, true, false, () => {
                                        // Call the update event
                                        this._onUpdate();
                                    });
                                }
                            }
                        });

                        // Load the context of the app catalog
                        ContextInfo.getWeb(AppConfig.Configuration.tenantAppCatalogUrl).execute(context => {
                            let requestDigest = context.GetContextWebInformation.FormDigestValue;
                            let web = Web(AppConfig.Configuration.tenantAppCatalogUrl, { requestDigest });

                            // Ensure this app can be deployed to the tenant
                            web.TenantAppCatalog().solutionContainsTeamsComponent(DataSource.DocSetTenantApp.ID).execute((resp: any) => {
                                // See if we can deploy this app to teams
                                if (resp.SolutionContainsTeamsComponent) {
                                    // Render the deploy to teams button
                                    tooltips.add({
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
                                                this._forms.deployToTeams(this._item, () => {
                                                    // Call the update event
                                                    this._onUpdate();
                                                });
                                            }
                                        }
                                    });
                                }
                            });
                        });
                    } else {
                        // Render the deploy button
                        tooltips.add({
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
                                    this._forms.deploy(this._item, true, () => {
                                        // Call the update event
                                        this._onUpdate();
                                    });
                                }
                            }
                        });
                    }
                    break;

                // Edit
                case "Edit":
                    // Render the button
                    tooltips.add({
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
                                this._forms.edit(this._item.Id, () => {
                                    // Call the update event
                                    this._onUpdate();
                                });
                            }
                        }
                    });
                    break;

                // Edit Tech Review
                case "EditTechReview":
                    // Render the button
                    tooltips.add({
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
                                this._forms.editTechReview(this._item, () => {
                                    // Call the update event
                                    this._onUpdate();
                                });
                            }
                        }
                    });
                    break;

                // Edit Test Cases
                case "EditTestCases":
                    // Render the button
                    tooltips.add({
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
                                this._forms.editTestCases(this._item, () => {
                                    // Call the update event
                                    this._onUpdate();
                                });
                            }
                        }
                    });
                    break;

                // Submit
                case "Submit":
                    // Render the button
                    tooltips.add({
                        content: "Submits the app for approval",
                        btnProps: {
                            text: "Submit for Approval",
                            iconClassName: "me-1",
                            iconSize: 20,
                            iconType: appIndicator,
                            isDisabled: !canEdit || (this.Status.lastStep),
                            isSmall: true,
                            type: Components.ButtonTypes.OutlinePrimary,
                            onClick: () => {
                                // Display the submit form
                                this._forms.submit(this._item, () => {
                                    // Call the update event
                                    this._onUpdate();
                                });
                            }
                        }
                    });
                    break;

                // Test Site
                case "TestSite":
                    // See if a test site exists
                    DataSource.loadTestSite(this._item).then(
                        // Test site exists
                        web => {
                            // Render the view button
                            tooltips.add({
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
                            tooltips.add({
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
                                        this._forms.createTestSite(this._item, () => {
                                            // Call the update event
                                            this._onUpdate();
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
                    tooltips.add({
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
                                this._forms.display(this._item.Id);
                            }
                        }
                    });
                    break;

                // View Tech Review
                case "ViewTechReview":
                    // Render the button
                    tooltips.add({
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
                                this._forms.displayTechReview(this._item);
                            }
                        }
                    });
                    break;

                // View Test Cases
                case "ViewTestCases":
                    // Render the button
                    tooltips.add({
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
                                this._forms.displayTestCases(this._item);
                            }
                        }
                    });
                    break;
            }
        }
    }
}