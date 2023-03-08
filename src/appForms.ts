import { LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, List, SPTypes, Web } from "gd-sprest-bs";
import { AppActions } from "./appActions";
import { AppConfig, IStatus } from "./appCfg";
import { AppNotifications } from "./appNotifications";
import { AppSecurity } from "./appSecurity";
import { DataSource, IAppCatalogRequestItem, IAppItem, IAssessmentItem } from "./ds";
import { ErrorDialog } from "./errorDialog";
import Strings from "./strings";

// Define the fields to exclude from the main tab
const AppExcludeFields = [
    "AppStatus",
    "IsAppPackageEnabled",
    "IsDefaultAppMetadataLocale"
];

// Define the fields to be read-only in the main tab
const AppReadOnlyFields = [
    "FileLeafRef"
];

// Define the app deployment fields
const AppDeploymentFields = [
    "AppIsTenantDeployed",
    "AppSiteDeployments"
];

// Define the app package properties
const AppPackageFields = [
    "Title",
    "AppAPIPermissions",
    "AppIsClientSideSolution",
    "AppIsDomainIsolated",
    "AppProductID",
    "AppSharePointMinVersion",
    "AppSkipFeatureDeployment",
    "AppVersion",
    "AppManifest"
];

// Define the app store properties
const AppStoreFields = [
    "AppPublisher",
    "AppShortDescription",
    "AppThumbnailURL",
    "AppImageURL1",
    "AppImageURL2",
    "AppImageURL3",
    "AppImageURL4",
    "AppImageURL5",
    "AppSupportURL",
    "AppVideoURL"
];

/**
 * App Item Forms
 */
export class AppForms {
    // Constructor
    constructor() { }

    // Approval form
    approve(item: IAppItem, onUpdate: () => void) {
        // Clear the modal
        Modal.clear();

        // Get the status information
        let status = AppConfig.Status[item.AppStatus];

        // Set the header
        Modal.setHeader("Approve " + item.AppStatus + " State");

        // Set the body
        Modal.setBody("Click 'Approve' to complete this approval step.");

        // Render the footer
        Modal.setFooter(Components.Tooltip({
            content: "This will approve the current state of the application.",
            btnProps: {
                text: "Approve",
                type: Components.ButtonTypes.OutlineSuccess,
                onClick: () => {
                    // Close the modal
                    Modal.hide();

                    // Ensure the metadata is complete
                    this.isMetadataComplete(item).then(isValid => {
                        // Ensure it's valid
                        if (!isValid) { return; }

                        // See if the tech review is valid
                        this.isValidTechReview(item, status).then(isValid => {
                            // Ensure it's valid
                            if (!isValid) { return; }

                            // See if the test cases are valid
                            this.isValidTestCases(item, status).then(isValid => {
                                // Ensure it's valid
                                if (!isValid) { return; }

                                // Show a loading dialog
                                LoadingDialog.setHeader("Updating the Status");
                                LoadingDialog.setBody("This dialog will close after the status is updated.");
                                LoadingDialog.show();

                                // Log
                                DataSource.logItem({
                                    LogUserId: ContextInfo.userId,
                                    ParentId: item.AppProductID,
                                    ParentListName: Strings.Lists.Apps,
                                    Title: DataSource.AuditLogStates.AppApproved,
                                    LogComment: `The app ${item.Title} was approved, from ${item.AppStatus} to ${status.nextStep}.`
                                }, DataSource.AppItem);

                                // Update the item
                                item.update({
                                    AppStatus: status.nextStep
                                }).execute(() => {
                                    // Close the dialog
                                    LoadingDialog.hide();

                                    // Run the flow associated for this status
                                    AppActions.runFlow(status.flowId);

                                    // Send the notifications
                                    AppNotifications.sendEmail(status.notification, item).then(() => {
                                        // See if the test app catalog exists and we are creating the test site
                                        if (AppSecurity.IsSiteAppCatalogOwner && status.createTestSite) {
                                            // Load the test site
                                            DataSource.loadTestSite(item).then(
                                                // Exists
                                                web => {
                                                    // See if the current version is deployed
                                                    if (item.AppVersion == DataSource.AppCatalogSiteItem.InstalledVersion && !DataSource.AppCatalogSiteItem.SkipDeploymentFeature) {
                                                        // Call the update event
                                                        onUpdate();
                                                    } else {
                                                        // Update the app
                                                        AppActions.updateApp(web.ServerRelativeUrl, true, onUpdate).then(() => {
                                                            // Call the update event
                                                            onUpdate();
                                                        });
                                                    }
                                                },

                                                // Doesn't exist
                                                () => {
                                                    // Create the test site
                                                    AppActions.createTestSite(onUpdate);
                                                }
                                            );
                                        } else {
                                            // Call the update event
                                            onUpdate();
                                        }
                                    }, ex => {
                                        // Log the error
                                        ErrorDialog.show("Updating Status", "There was an error updating the status.", ex);
                                    });
                                });
                            });
                        });
                    });
                }
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Creates the test site for the application
    createTestSite(errorMessage: string, onUpdate: () => void) {
        // Set the header
        Modal.setHeader(errorMessage ? "App Error" : "Create Test Site");

        // Set the body
        Modal.setBody(errorMessage || "Are you sure you want to create the test site for this application?");

        // Render the footer
        Modal.setFooter(Components.Button({
            text: "Create Site",
            type: Components.ButtonTypes.OutlineSuccess,
            isDisabled: errorMessage ? true : false,
            onClick: () => {
                // Close the modal
                Modal.hide();

                // Code to run after the logic below completes
                let onComplete = () => {
                    // Create the test site
                    AppActions.createTestSite(web => {
                        // Log
                        DataSource.logItem({
                            LogUserId: ContextInfo.userId,
                            ParentId: DataSource.AppItem.AppProductID,
                            ParentListName: Strings.Lists.Apps,
                            Title: DataSource.AuditLogStates.CreateTestSite,
                            LogComment: `The app ${DataSource.AppItem.Title} test site was created successfully at: ${web.ServerRelativeUrl}`
                        }, DataSource.AppItem);

                        // Call the update event
                        onUpdate();
                    });
                }

                // See if the app was deployed, but errored out
                if (DataSource.AppCatalogItem && DataSource.AppCatalogItem.AppPackageErrorMessage && DataSource.AppCatalogItem.AppPackageErrorMessage != "No errors.") {
                    // Delete the item
                    DataSource.AppCatalogItem.delete().execute(() => {
                        // Create the test site
                        onComplete();
                    });
                } else {
                    // Create the test site
                    onComplete();
                }
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Delete form
    delete(onUpdate: () => void) {
        // Set the header
        Modal.setHeader("Delete App/Solution Package");

        // Set the body
        Modal.setBody("Are you sure you want to delete this app package? This will remove it from the tenant app catalog.");

        // Render the footer
        Modal.setFooter(Components.Button({
            text: "Delete",
            type: Components.ButtonTypes.OutlineDanger,
            onClick: () => {
                // Close the modal
                Modal.hide();

                // Show a loading dialog
                LoadingDialog.setHeader("Retracting the App");
                LoadingDialog.setBody("Removing the app from the app catalog(s).");
                LoadingDialog.show();

                // Retract the solution from the site collection app catalog
                AppActions.retract(false, true, () => {
                    // Retract the solution from the tenant app catalog
                    AppActions.retract(true, true, () => {
                        // Delete the test site
                        AppActions.deleteTestSite().then(() => {
                            // Update the loading dialog
                            LoadingDialog.setHeader("Removing Assessments");
                            LoadingDialog.setBody("Removing the assessments associated with this app.");

                            // Delete the assessments w/ this app
                            this.deleteAssessments(DataSource.AppItem.Id).then(() => {
                                // Update the loading dialog
                                LoadingDialog.setHeader("Deleting the App Request");
                                LoadingDialog.setBody("This dialog will close after the app request is deleted.");

                                // Delete this folder
                                DataSource.AppItem.delete().execute(() => {
                                    // Close the dialog
                                    LoadingDialog.hide();

                                    // Log
                                    DataSource.logItem({
                                        LogUserId: ContextInfo.userId,
                                        ParentId: DataSource.AppItem.AppProductID,
                                        ParentListName: Strings.Lists.Apps,
                                        Title: DataSource.AuditLogStates.DeleteApp,
                                        LogComment: `The app ${DataSource.AppItem.Title} was deleted.`
                                    }, DataSource.AppItem);

                                    // Execute the update event
                                    onUpdate();
                                });
                            });
                        });
                    });
                });
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Method to delete the assessments associated with the app
    private deleteAssessments(itemId: number): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the assoicated item
            List(Strings.Lists.Assessments).Items().query({
                Filter: "RelatedAppId eq " + itemId
            }).execute(
                items => {
                    // Parse the items
                    Helper.Executor(items.results, item => {
                        // Return a promise
                        return new Promise(resolve => {
                            // Delete the item
                            item.delete().execute(() => { resolve(null); });
                        });
                    }).then(() => {
                        // Resolve the request
                        resolve();
                    });
                },
                ex => {
                    // Log the error
                    ErrorDialog.show("Deleting Assessments", "There was an error deleting the assessments", ex);
                }
            );
        });
    }

    // Delete test site form
    deleteSite(item: IAppItem, onUpdate: () => void) {
        // Set the header
        Modal.setHeader("Delete Test Site");

        // Set the body
        Modal.setBody("Are you sure you want to delete the test site?");

        // Render the footer
        Modal.setFooter(Components.Button({
            text: "Delete",
            type: Components.ButtonTypes.OutlineDanger,
            onClick: () => {
                // Close the modal
                Modal.hide();

                // Delete the test site
                AppActions.deleteTestSite().then(() => {
                    // Log
                    DataSource.logItem({
                        LogUserId: ContextInfo.userId,
                        ParentId: item.AppProductID,
                        ParentListName: Strings.Lists.Apps,
                        Title: DataSource.AuditLogStates.DeleteTestSite,
                        LogComment: `The app ${item.Title} test site was deleted.`
                    }, item);

                    // Execute the update event
                    onUpdate();
                });
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Deploys the solution to the app catalog
    deploy(item: IAppItem, tenantFl: boolean, onUpdate: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Deploy to " + (tenantFl ? " Tenant " : " Site Collection ") + "App Catalog");

        // See if the app can be deployed to all sites
        let form: Components.IForm = null;
        if (item.AppSkipFeatureDeployment) {
            // Create the form
            form = Components.Form({
                el: Modal.BodyElement,
                controls: [
                    {
                        name: "SkipFeatureDeployment",
                        label: "Skip Feature Deployment?",
                        description: "Select this option if you are deploying to teams or if you want to make the solution available to all sites immediately.",
                        type: Components.FormControlTypes.Switch,
                        value: true
                    }
                ]
            });

            // Append the confirmation message
            let elConfirmation = document.createElement("p");
            elConfirmation.innerHTML = "Click 'Deploy' to deploy the solution to the app catalog.";
            Modal.BodyElement.appendChild(elConfirmation);
        } else {
            // Set the body
            Modal.setBody("Are you sure you want to deploy the application?");
        }

        // Render the footer
        Modal.setFooter(Components.Button({
            text: "Deploy",
            type: Components.ButtonTypes.OutlineSuccess,
            onClick: () => {
                // Close the modal
                Modal.hide();

                // Set the flag
                let skipFeatureDeployment = form ? form.getValues()["SkipFeatureDeployment"] : false;

                // Deploy the app
                AppActions.deploy(tenantFl, skipFeatureDeployment, onUpdate, () => {
                    // See if there is a flow
                    if (AppConfig.Configuration.appFlows && AppConfig.Configuration.appFlows.deployToTenant) {
                        // Execute the flow
                        AppActions.runFlow(AppConfig.Configuration.appFlows.deployToTenant);
                    }

                    // Log
                    DataSource.logItem({
                        LogUserId: ContextInfo.userId,
                        ParentId: item.AppProductID,
                        ParentListName: Strings.Lists.Apps,
                        Title: tenantFl ? DataSource.AuditLogStates.AppTenantDeployed : DataSource.AuditLogStates.AppDeployed,
                        LogComment: `The app ${item.Title} was deployed to: ${tenantFl ? AppConfig.Configuration.tenantAppCatalogUrl : AppConfig.Configuration.appCatalogUrl}`
                    }, item);

                    // Call the update event
                    onUpdate();
                });
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Deploys the solution to a site collection app catalog
    deployToSite(item: IAppItem, onUpdate: () => void) {
        let appCatalogExists = false;
        let canDeploy = false;
        let errorMessage = "";
        let webUrl = null;

        // Set the header
        Modal.setHeader("Deploy to App Catalog");

        // Create the form
        let form = Components.Form({
            controls: [
                {
                    name: "Url",
                    title: "Site Collection Url",
                    description: "The relative url to the site collection containing the app catalog to deploy the application to.",
                    type: Components.FormControlTypes.TextField,
                    required: true,
                    onControlRendering: ctrl => {
                        (ctrl as Components.IFormControlPropsTextField).onChange = () => {
                            // Disable the buttons
                            btnDeploy.disable();
                            btnRequest.disable();
                        }
                    },
                    onValidate: (ctrl, results) => {
                        // See if there is an error loading the site
                        if (errorMessage) {
                            // Set the flag
                            results.isValid = false;

                            // Set the error message
                            results.invalidMessage = errorMessage;
                        } else {
                            // Set the flag
                            results.isValid = true;
                        }

                        // Return the results
                        return results;
                    }
                }
            ]
        });

        // Set the body
        Modal.setBody(form.el);

        // Render the footer
        let btnDeploy: Components.IButton = null;
        let btnRequest: Components.IButton = null;
        Modal.setFooter(Components.TooltipGroup({
            tooltips: [
                {
                    content: "Loads the site collection and validates the app catalog.",
                    btnProps: {
                        text: "Load Site",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Reset the flags
                            appCatalogExists = false;
                            canDeploy = false;
                            errorMessage = "";
                            webUrl = null;

                            // Disable the buttons
                            btnDeploy.disable();
                            btnRequest.disable();

                            // Validate the url
                            // Note - trim the '/' from the end of the url. It can cause the web query to fail.
                            let url = (form.getValues()["Url"] || "").trim().replace(/\/$/, "");
                            if (url) {
                                // Show a loading dialog
                                LoadingDialog.setHeader("Loading App Catalog");
                                LoadingDialog.setBody("This dialog will close after the app catalog is loaded.");
                                LoadingDialog.show();

                                // Determine if the user has access to the app catalog
                                ((): PromiseLike<void> => {
                                    // Return a promise
                                    return new Promise((resolve) => {
                                        // Query the web
                                        Web(url).query({
                                            Expand: ["Lists", "Lists/EffectiveBasePermissions"],
                                            Select: ["Title", "Lists/Title", "Lists/EffectiveBasePermissions"]
                                        }).execute(web => {
                                            // Parse the lists
                                            for (let i = 0; i < web.Lists.results.length; i++) {
                                                // See if this is the app catalog
                                                let list = web.Lists.results[i];
                                                if (list.Title == Strings.Lists.AppCatalog) {
                                                    // Set the flag
                                                    appCatalogExists = true;

                                                    // Determine if the user can deploy to it
                                                    canDeploy = Helper.hasPermissions(list.EffectiveBasePermissions, [
                                                        SPTypes.BasePermissionTypes.FullMask
                                                    ]);

                                                    // Break from the loop
                                                    break;
                                                }
                                            }

                                            // See if the app catalog doesn't exists
                                            if (!appCatalogExists) {
                                                // Set the error message
                                                errorMessage = "The app catalog does not exist in this site.";
                                            }
                                            // Else, see if the user can't deploy
                                            else if (!canDeploy) {
                                                // Set the error message
                                                errorMessage = "The user does not have the permissions to deploy to the app catalog.";
                                            }

                                            // Resolve the request
                                            resolve();
                                        }, () => {
                                            // Set the error message
                                            errorMessage = "Site does not exist or user does not have access to it.";

                                            // Resolve the request
                                            resolve();
                                        });
                                    });
                                })().then(() => {
                                    // Ensure the form is valid
                                    if (form.isValid()) {
                                        // See if the app catalog doesn't exists
                                        if (!appCatalogExists) {
                                            // Enable the request button
                                            btnRequest.enable();
                                        }

                                        // See if the user can deploy
                                        if (canDeploy) {
                                            // Enable the deploy button
                                            btnDeploy.enable();
                                        }
                                    }

                                    // Hide the loading dialog
                                    LoadingDialog.hide();
                                });
                            } else {
                                // Set the error message
                                errorMessage = "Please enter a valid url to the site collection.";

                                // Show the default error message
                                form.isValid();
                            }
                        }
                    }
                },
                {
                    content: "Deploys the application to the site collection's app catalog.",
                    btnProps: {
                        assignTo: btn => { btnDeploy = btn; },
                        isDisabled: true,
                        text: "Deploy",
                        type: Components.ButtonTypes.OutlineSuccess,
                        onClick: () => {
                            // Close the modal
                            Modal.hide();

                            // Deploy the app
                            let siteUrl = form.getValues()["Url"];
                            AppActions.deployToSite(siteUrl, () => {
                                // See if there is a flow
                                if (AppConfig.Configuration.appFlows && AppConfig.Configuration.appFlows.deployToSiteCollection) {
                                    // Execute the flow
                                    AppActions.runFlow(AppConfig.Configuration.appFlows.deployToSiteCollection);
                                }

                                // Log
                                DataSource.logItem({
                                    LogUserId: ContextInfo.userId,
                                    ParentId: item.AppProductID,
                                    ParentListName: Strings.Lists.Apps,
                                    Title: DataSource.AuditLogStates.AppDeployed,
                                    LogComment: `The app ${item.Title} was deployed to: ${siteUrl}`
                                }, item);

                                // Call the update event
                                onUpdate();
                            });
                        }
                    }
                },
                {
                    content: "Creates a request for an app catalog.",
                    btnProps: {
                        assignTo: btn => { btnRequest = btn; },
                        isDisabled: true,
                        text: "Request",
                        type: Components.ButtonTypes.OutlineInfo,
                        onClick: () => {
                            // Display the request form
                            this.requestAppCatalog(webUrl);
                        }
                    }
                },
            ]
        }).el);

        // Show the modal
        Modal.show();
    }

    // Deploys the solution to teams
    deployToTeams(item: IAppItem, onUpdate: () => void) {
        // Set the header
        Modal.setHeader("Deploy to Teams");

        // Set the body
        Modal.setBody("Are you sure you want to deploy teams?");

        // Render the footer
        Modal.setFooter(Components.Button({
            text: "Deploy",
            type: Components.ButtonTypes.OutlineSuccess,
            onClick: () => {
                // Close the modal
                Modal.hide();

                // Deploy the app
                AppActions.deployToTeams(() => {
                    // See if there is a flow
                    if (AppConfig.Configuration.appFlows && AppConfig.Configuration.appFlows.deployToTeams) {
                        // Execute the flow
                        AppActions.runFlow(AppConfig.Configuration.appFlows.deployToTeams);
                    }

                    // Call the update event
                    onUpdate();
                });
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Display form
    display(itemId: number) {
        // Show the display form
        DataSource.DocSetList.viewForm({
            itemId,
            webUrl: Strings.SourceUrl,
            onSetHeader: el => {
                // Update the header
                el.className = "h5 m-0";
                el.innerHTML = "App Details";
            },
            onGetListInfo: props => {
                // Set the content type
                props.contentType = "App";

                // Return the properties
                return props;
            },
            tabInfo: {
                tabs: [
                    {
                        title: "Properties",
                        excludeFields: AppExcludeFields.concat(AppDeploymentFields, AppPackageFields, AppStoreFields)
                    },
                    {
                        title: "Metadata",
                        fields: AppPackageFields
                    },
                    {
                        title: "Store Details",
                        fields: AppStoreFields
                    },
                    {
                        title: "Deployment",
                        fields: AppDeploymentFields
                    }
                ]
            },
            onCreateViewForm: props => {
                // Update the field
                props.onControlRendering = (ctrl, field) => {
                    // See if this is a url field
                    if (field.InternalName.indexOf("URL") > 0) {
                        // Hide the description field
                        (ctrl as Components.IFormControlUrlProps).showDescription = false;
                    }
                }

                // Return the properties
                return props;
            }
        });
    }

    // Technical review form
    displayTechReview(item: IAppItem) {
        // Get the assessment item
        this.getAssessmentItem(item, false).then(assessment => {
            // See if an item exists
            if (assessment) {
                // Display the the view form
                DataSource.AppAssessments.viewForm({
                    itemId: assessment.Id,
                    webUrl: Strings.SourceUrl,
                    onGetListInfo: (props) => {
                        // Exclude the related app field
                        props.excludeFields = ["RelatedApp"];
                        return props;
                    }
                });
            } else {
                // No assessment found
                Modal.clear();
                Modal.setHeader(item.Title + " - Technical Review");
                Modal.setBody("No technical review assessment was found.");
                Modal.show();
            }
        });
    }

    // Views the tests for the application
    displayTestCases(item: IAppItem) {
        // Get the assessment item
        this.getAssessmentItem(item).then(assessment => {
            // See if an item exists
            if (assessment) {
                // Display the the view form
                DataSource.AppAssessments.viewForm({
                    itemId: assessment.Id,
                    webUrl: Strings.SourceUrl,
                    onGetListInfo: props => {
                        // Exclude the related app field
                        props.excludeFields = ["RelatedApp"];

                        // Set the content type
                        props.contentType = "TestCases";

                        // Return the properties
                        return props;
                    }
                });
            } else {
                // No assessment found
                Modal.clear();
                Modal.setHeader(item.Title + " - Technical Review");
                Modal.setBody("No technical review assessment was found.");
                Modal.show();
            }
        });
    }

    // Edit form
    edit(itemId: number, onUpdate: () => void) {
        // Display the edit form
        DataSource.DocSetList.editForm({
            itemId,
            webUrl: Strings.SourceUrl,
            onSetFooter: el => {
                // Add a cancel button
                Components.Button({
                    el,
                    text: "Cancel",
                    type: Components.ButtonTypes.OutlineSecondary,
                    onClick: () => {
                        Modal.hide();
                    }
                });
            },
            onSetHeader: el => {
                // Update the header
                el.className = "h5 m-0";
                el.innerHTML = "App Details";
            },
            onGetListInfo: props => {
                // Set the content type
                props.contentType = "App";

                // Return the properties
                return props;
            },
            tabInfo: {
                tabs: [
                    {
                        title: "Properties",
                        excludeFields: AppExcludeFields.concat(AppDeploymentFields, AppPackageFields, AppStoreFields)
                    },
                    {
                        title: "Metadata",
                        fields: AppPackageFields
                    },
                    {
                        title: "Store Details",
                        fields: AppStoreFields
                    },
                    {
                        title: "Deployment",
                        fields: AppDeploymentFields
                    }
                ]
            },
            onCreateEditForm: props => {
                // Set the control rendering event
                props.onControlRendering = (ctrl, field) => {
                    // See if this is a read-only field
                    if (AppReadOnlyFields.indexOf(field.InternalName) >= 0 || AppPackageFields.indexOf(field.InternalName) >= 0 || AppDeploymentFields.indexOf(field.InternalName) >= 0) {
                        // Make it read-only
                        ctrl.isReadonly = true;
                    }

                    // See if this is the permissions justification
                    if (field.InternalName == "AppPermissionsJustification") {
                        // Add validation
                        ctrl.onValidate = (ctrl, results) => {
                            // See if permissions exist
                            let apiPermissions = DataSource.DocSetList.EditForm.getItem()["AppAPIPermissions"];
                            if (apiPermissions) {
                                // Set the flag
                                results.isValid = (results.value || "").trim().length > 0;

                                // Set the error message
                                results.invalidMessage = "A justification is required if API permissions exist.";
                            } else {
                                // It's valid
                                results.isValid = true;
                            }

                            // Return the results
                            return results;
                        }
                    }

                    // See if this is the teams field
                    if (field.InternalName == "AppIsTeams") {
                        // Disable it by default
                        ctrl.isDisabled = true;

                        try {
                            // Get the manifest information
                            let manifest = JSON.parse(DataSource.AppItem.AppManifest);
                            if (manifest) {
                                // See if the app supports teams
                                if (manifest.supportedHosts.indexOf("TeamsTab") >= 0 ||
                                    manifest.supportedHosts.indexOf("TeamsPersonalApp") >= 0) {
                                    // Enable the control
                                    ctrl.isDisabled = false;
                                }
                            }
                        } catch { }
                    }

                    // See if this is a url field
                    if (field.InternalName.indexOf("URL") > 0 || field.InternalName == "AppSourceControl") {
                        // Hide the description field
                        (ctrl as Components.IFormControlUrlProps).showDescription = false;
                    }
                    /* TODO - Need to get the CSS corrected for switches
                    // Else, see if this is a boolean field
                    else if(field.FieldTypeKind == SPTypes.FieldType.Boolean) {
                        // Set the option to display as a switch
                        (ctrl as Components.IFormControlPropsCheckbox).type = Components.FormControlTypes.Switch;
                    }
                    */
                }

                // Set the control rendered event
                props.onControlRendered = (ctrl, field) => {
                    // See if this is a URL Image or Video field
                    if (field.InternalName.startsWith("AppImageURL") || field.InternalName == "AppVideoURL") {
                        let isImage = field.InternalName != "AppVideoURL";

                        // Make this field read-only
                        ctrl.textbox.elTextbox.readOnly = true;

                        // Set a click event
                        ctrl.textbox.elTextbox.addEventListener("click", () => {
                            // Method to upload a file to the docset folder
                            let uploadFile = (fileName: string, file: Helper.IListFormAttachmentInfo) => {
                                // Return a promise
                                return new Promise((resolve, reject) => {
                                    // Show a loading dialog
                                    LoadingDialog.setHeader("Uploading File");
                                    LoadingDialog.setBody("This will close after the file is uploaded");
                                    LoadingDialog.show();

                                    // Upload the file
                                    Web(Strings.SourceUrl).Lists(Strings.Lists.Apps).Items(DataSource.AppItem.Id).Folder().Files().add(fileName, true, file.data).execute(
                                        // Success
                                        file => {
                                            // Close the dialog
                                            LoadingDialog.hide();

                                            // Resolve the request
                                            resolve(file);
                                        },
                                        // Error
                                        () => {
                                            // Close the dialog
                                            LoadingDialog.hide();

                                            // Reject the request
                                            reject();
                                        }
                                    );
                                });
                            }

                            // Display a file upload dialog
                            Helper.ListForm.showFileDialog().then(file => {
                                // Clear the value
                                ctrl.textbox.setValue("");

                                // Get the file name
                                let fileName = file.name.toLowerCase();

                                // Validate the file type
                                if (isImage) {
                                    if (fileName.endsWith(".png") || fileName.endsWith(".jpg") || fileName.endsWith(".jpeg") || fileName.endsWith(".gif")) {
                                        // Upload the file
                                        let fileInfo = fileName.split('.');
                                        let dstFileName = "AppImage" + field.InternalName.replace("AppImageURL", '') + "." + fileInfo[fileInfo.length - 1];
                                        let fileUrl = [document.location.origin, DataSource.AppFolder.ServerRelativeUrl.replace(/^\//, ''), dstFileName].join('/');
                                        uploadFile(dstFileName, file).then(
                                            () => {
                                                // Set the value
                                                ctrl.textbox.setValue(fileUrl);

                                                // Validate the control
                                                ctrl.updateValidation(ctrl.el, {
                                                    isValid: true
                                                });
                                            },
                                            // Error
                                            () => {
                                                // Display an error message
                                                ctrl.updateValidation(ctrl.el, {
                                                    isValid: false,
                                                    invalidMessage: "Error uploading the file to the app folder. Refresh and try again."
                                                });
                                            }
                                        );
                                    } else {
                                        // Display an error message
                                        ctrl.updateValidation(ctrl.el, {
                                            isValid: false,
                                            invalidMessage: "The file must be a valid image file. Valid types: png, jpg, jpeg, gif"
                                        });
                                    }
                                } else {
                                    // Else, ensure it's a video
                                    if (fileName.endsWith(".mpg") || fileName.endsWith(".mpeg") || fileName.endsWith(".avi") || fileName.endsWith(".mp4")) {
                                        // Upload the file
                                        let fileInfo = fileName.split('.');
                                        let dstFileName = "AppVideo" + "." + fileInfo[fileInfo.length - 1];
                                        let fileUrl = [document.location.origin, DataSource.AppFolder.ServerRelativeUrl.replace(/^\//, ''), dstFileName].join('/');
                                        uploadFile(dstFileName, file).then(
                                            () => {
                                                // Set the value
                                                ctrl.textbox.setValue(fileUrl);

                                                // Validate the control
                                                ctrl.updateValidation(ctrl.el, {
                                                    isValid: true
                                                });
                                            },
                                            // Error
                                            () => {
                                                // Display an error message
                                                ctrl.updateValidation(ctrl.el, {
                                                    isValid: false,
                                                    invalidMessage: "Error uploading the file to the app folder. Refresh and try again."
                                                });
                                            }
                                        );
                                    } else {
                                        // Display an error message
                                        ctrl.updateValidation(ctrl.el, {
                                            isValid: false,
                                            invalidMessage: "The file must be a valid video file. Valid types: mpg, mpeg, avi, mp4"
                                        });
                                    }
                                }
                            });
                        });
                    }
                }

                // Return the properties
                return props;
            },
            onValidation: (values) => {
                // Save the form by default
                return true;
            },
            onUpdate: (item: IAppItem) => {
                // See if this is a new item
                if (item.AppStatus == "New") {
                    // Log
                    ErrorDialog.logInfo(`Validating the app sponsor id: '${item.AppSponsorId}'`);

                    // Get the sponsor
                    let sponsor = AppSecurity.AppWeb.getUserForGroup(Strings.Groups.Sponsors, item.AppSponsorId);
                    if (sponsor == null && item.AppSponsorId > 0) {
                        // Log
                        ErrorDialog.logInfo(`App sponsor not in group. Adding the user...`);

                        // Add the sponsor to the group
                        AppSecurity.AppWeb.addUserToGroup(Strings.Groups.Sponsors, item.AppSponsorId).then(() => {
                            // Get the status
                            let status = AppConfig.Status[item.AppStatus];

                            // Send the notifications
                            AppNotifications.sendEmail(status.notification, item, false).then(() => {
                                // Call the update event
                                onUpdate();

                                // Hide the dialog
                                LoadingDialog.hide();
                            });
                        });
                    } else {
                        // Call the update event
                        onUpdate();
                    }
                } else {
                    // Call the update event
                    onUpdate();
                }
            }
        });
    }

    // Technical review form
    editTechReview(item: IAppItem, onUpdate: () => void) {
        // Displays the assessment form
        let displayEditForm = (assessment: IAssessmentItem) => {
            let validateFl = false;

            // Display the edit form
            DataSource.AppAssessments.editForm({
                itemId: assessment.Id,
                webUrl: Strings.SourceUrl,
                onSetFooter: el => {
                    // Render a completed button
                    Components.Tooltip({
                        el,
                        content: "Validates the form.",
                        btnProps: {
                            text: "Validate",
                            type: Components.ButtonTypes.OutlineSuccess,
                            onClick: () => {
                                // Set the flag
                                validateFl = true;

                                // Validate the form
                                DataSource.AppAssessments.EditForm.isValid();
                            }
                        }
                    });
                },
                onCreateEditForm: props => {
                    // Set the rendering event
                    props.onControlRendering = (props, field) => {
                        // See if this is a test case field
                        if (field.InternalName.indexOf("TechReview") == 0 && field.FieldTypeKind == SPTypes.FieldType.Choice) {
                            // Set the validation
                            props.onValidate = (ctrl, results) => {
                                // Clear the invalid message
                                results.invalidMessage = "";

                                // See if we are validating the results
                                if (validateFl) {
                                    // See if there is validation values for this field
                                    let values = AppConfig.Configuration.validation && AppConfig.Configuration.validation.techReview ? AppConfig.Configuration.validation.techReview[field.InternalName] : null;
                                    if (values && values.length > 0) {
                                        // Set the flag
                                        let selectedItem: Components.IDropdownItem = results.value;
                                        if (selectedItem && selectedItem.text) {
                                            let matchFl = false;

                                            // Parse the valid values
                                            for (let i = 0; i < values.length; i++) {
                                                // See if the value matches
                                                if (values[i] == selectedItem.text) {
                                                    // Set the flag and break from the loop
                                                    matchFl = true;
                                                    break;
                                                }
                                            }

                                            // See if a match wasn't found
                                            if (!matchFl) {
                                                // Value isn't valid
                                                results.isValid = false;

                                                // Set the error message
                                                results.invalidMessage = "The selected value will fail to approve the item.";
                                            }
                                        } else {
                                            // Value isn't valid
                                            results.isValid = false;

                                            // Set the error message
                                            results.invalidMessage = "A selection is required for approval.";
                                        }
                                    }
                                }

                                // Return the results
                                return results;
                            }
                        }
                    }

                    // Return the properties
                    return props;
                },
                onValidation: (values) => {
                    // Save the form by default
                    return true;
                },
                onUpdate: () => {
                    // Call the update event
                    onUpdate();
                }
            });
        }

        // Get the assessment item
        this.getAssessmentItem(item, false).then(assessment => {
            // See if an item exists
            if (assessment) {
                // Show the assessment form
                displayEditForm(assessment);
            } else {
                // Show a loading dialog
                LoadingDialog.setHeader("Creating the Form");
                LoadingDialog.setBody("This dialog will close after the form is created.");

                // Create the item
                DataSource.AppAssessments.createItem({
                    RelatedAppId: item.Id,
                    Title: "Technical Review: " + item.Title
                }).then(
                    item => {
                        // Show the assessment form
                        displayEditForm(item);
                    },
                    ex => {
                        // Log the error
                        ErrorDialog.show("Adding Assessment", "There was an error adding the assessment.", ex);
                    }
                );
            }
        });
    }

    // Views the tests for the application
    editTestCases(item: IAppItem, onUpdate: () => void) {
        // Displays the form
        let displayEditForm = (assessment: IAssessmentItem) => {
            let validateFl = false;

            // Show the edit form
            DataSource.AppAssessments.editForm({
                itemId: assessment.Id,
                webUrl: Strings.SourceUrl,
                onGetListInfo: props => {
                    // Set the content type
                    props.contentType = "TestCases";

                    // Return the properties
                    return props;
                },
                onSetFooter: el => {
                    // Render a completed button
                    Components.Tooltip({
                        el,
                        content: "Validates the form.",
                        btnProps: {
                            text: "Validate",
                            type: Components.ButtonTypes.OutlineSuccess,
                            onClick: () => {
                                // Set the flag
                                validateFl = true;

                                // Validate the form
                                DataSource.AppAssessments.EditForm.isValid();
                            }
                        }
                    });
                },
                onCreateEditForm: props => {
                    // Set the rendering event
                    props.onControlRendering = (props, field) => {
                        // See if this is a test case field
                        if (field.InternalName.indexOf("TestCase") == 0 && field.FieldTypeKind == SPTypes.FieldType.Choice) {
                            // Set the validation
                            props.onValidate = (ctrl, results) => {
                                // See if we are validating the results
                                if (validateFl) {
                                    // See if there is validation values for this field
                                    let values = AppConfig.Configuration.validation && AppConfig.Configuration.validation.testCases ? AppConfig.Configuration.validation.testCases[field.InternalName] : null;
                                    if (values && values.length > 0) {
                                        // Set the flag
                                        let selectedItem: Components.IDropdownItem = results.value;
                                        if (selectedItem && selectedItem.text) {
                                            let matchFl = false;

                                            // Parse the valid values
                                            for (let i = 0; i < values.length; i++) {
                                                // See if the value matches
                                                if (values[i] == selectedItem.text) {
                                                    // Set the flag and break from the loop
                                                    matchFl = true;
                                                    break;
                                                }
                                            }

                                            // See if a match wasn't found
                                            if (!matchFl) {
                                                // Value isn't valid
                                                results.isValid = false;

                                                // Set the error message
                                                results.invalidMessage = "The selected value will fail to approve the item.";
                                            }
                                        } else {
                                            // Value isn't valid
                                            results.isValid = false;

                                            // Set the error message
                                            results.invalidMessage = "A selection is required for approval.";
                                        }
                                    }
                                }

                                // Return the results
                                return results;
                            }
                        }
                    }

                    // Return the properties
                    return props;
                },
                onValidation: (values) => {
                    // Save the form by default
                    return true;
                },
                onUpdate: () => {
                    // Call the update event
                    onUpdate();
                }
            });
        }

        // Get the assessment item
        this.getAssessmentItem(item).then(assessment => {
            // See if an item exists
            if (assessment) {
                // Show the assessment form
                displayEditForm(assessment);
            } else {
                // Show a loading dialog
                LoadingDialog.setHeader("Creating the Form");
                LoadingDialog.setBody("This dialog will close after the form is created.");

                // Get the test content type
                Web(Strings.SourceUrl).Lists(Strings.Lists.Assessments).ContentTypes().query({
                    Filter: "Name eq 'TestCases'"
                }).execute(cts => {
                    // Set the content type
                    let ct = cts.results[0];

                    // Create the item
                    DataSource.AppAssessments.createItem({
                        ContentTypeId: ct ? ct.StringId : null,
                        RelatedAppId: item.Id,
                        Title: "Test Cases: " + item.Title
                    }).then(
                        item => {
                            // Show the assessment form
                            displayEditForm(item);
                        },
                        ex => {
                            // Log the error
                            ErrorDialog.show("Adding Assessment", "There was an error adding the assessment.", ex);
                        }
                    );
                });
            }
        });
    }

    // Method to get the assessment item associated with the app
    private getAssessmentItem(item: IAppItem, testFl: boolean = true): PromiseLike<IAssessmentItem> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the associated items
            DataSource.loadAppAssessments(item.Id, testFl).then(items => {
                // Return the last item
                resolve(items[0]);
            }, reject);
        });
    }

    // Method to determine if the required metadata has been completed
    private isMetadataComplete(item: IAppItem): PromiseLike<boolean> {
        // Return a promise
        return new Promise(resolve => {
            let invalidFields: Components.IListGroupItem[] = [];

            // Show a loading dialog
            LoadingDialog.setHeader("Validating the Metadata");
            LoadingDialog.setBody("This dialog will close after the validation completes.");
            LoadingDialog.show();

            // Get the content type
            for (let i = 0; i < DataSource.DocSetList.ListContentTypes.length; i++) {
                let ct = DataSource.DocSetList.ListContentTypes[i];

                // See if this is the target content type
                if (ct.Name == "App") {
                    // Parse the fields
                    for (let i = 0; i < ct.FieldLinks.results.length; i++) {
                        let fieldLink = ct.FieldLinks.results[i];
                        let field = ct.Fields.results[i];

                        // Set the field name based on the type
                        let fieldName = field.InternalName;
                        if (field.FieldTypeKind == SPTypes.FieldType.Lookup || field.FieldTypeKind == SPTypes.FieldType.User) {
                            fieldName += "Id";
                        }

                        // See if this is a required field
                        if (fieldLink.Required) {
                            // Ensure a value exists
                            if (item[fieldName]) { continue; }

                            // Append the field
                            invalidFields.push({
                                content: field.Title
                            });
                        }
                    }

                    // Break from the loop
                    break;
                }
            }

            // Hide the dialog
            LoadingDialog.hide();

            // See if it's not valid
            if (invalidFields.length > 0) {
                // Clear the modal
                Modal.clear();

                // Set the header
                Modal.setHeader("Error");

                // Set the body
                Modal.setBody("The following app data was not completed for this app:");

                // Set the fields
                Components.ListGroup({
                    className: "mt-2",
                    el: Modal.BodyElement,
                    items: invalidFields
                });

                // Show the dialog
                Modal.show();
            }

            // Resolve the request
            resolve(invalidFields.length == 0);
        });
    }

    // Method to determine if the tech review is valid
    private isValidTechReview(item: IAppItem, status: IStatus): PromiseLike<boolean> {
        // Return a promise
        return new Promise(resolve => {
            let isValid = true;

            // See if this approval requires a technical review and validation exists
            if (status.requiresTechReview && AppConfig.Configuration.validation && AppConfig.Configuration.validation.techReview) {
                // Show a loading dialog
                LoadingDialog.setHeader("Validating the Technical Review");
                LoadingDialog.setBody("This dialog will close after the validation completes.");
                LoadingDialog.show();

                // Get the assessment
                this.getAssessmentItem(item, false).then(item => {
                    // Ensure the item exists
                    if (item) {
                        // Parse the validation
                        for (let fieldName in AppConfig.Configuration.validation.techReview) {
                            // Ensure values exist
                            let values = AppConfig.Configuration.validation.techReview[fieldName];
                            if (values && values.length > 0) {
                                let matchFl = false;

                                // Parse the valid values
                                for (let i = 0; i < values.length; i++) {
                                    // See if the value matches
                                    if (values[i] == item[fieldName]) {
                                        // Set the flag and break from the loop
                                        matchFl = true;
                                        break;
                                    }
                                }

                                // See if a match wasn't found
                                if (!matchFl) {
                                    // Form isn't valid
                                    isValid = false;
                                    break;
                                }
                            }
                        }

                        // Hide the dialog
                        LoadingDialog.hide();

                        // See if it's valid
                        if (isValid) {
                            // Update the date completed field
                            item.update({ Completed: (new Date(Date.now())).toISOString() }).execute(() => {
                                // Resolve the request
                                resolve(isValid);
                            });
                        } else {
                            // Clear the modal
                            Modal.clear();

                            // Set the header
                            Modal.setHeader("Error");

                            // Set the body
                            Modal.setBody("The technical review is not valid. Please review it and try your request again.");

                            // Show the dialog
                            Modal.show();

                            // Resolve the request
                            resolve(isValid);
                        }
                    } else {
                        // Hide the dialog
                        LoadingDialog.hide();

                        // Clear the modal
                        Modal.clear();

                        // Set the header
                        Modal.setHeader("Error");

                        // Set the body
                        Modal.setBody("No technical review exists for this application.");

                        // Show the dialog
                        Modal.show();

                        // Not valid
                        resolve(false);
                    }
                });
            } else {
                // Resolve the request
                resolve(isValid);
            }
        });
    }

    // Method to determine if the test cases are valid
    private isValidTestCases(item: IAppItem, status: IStatus): PromiseLike<boolean> {
        // Return a promise
        return new Promise(resolve => {
            let isValid = true;

            // See if this approval requires test cases and validation exists
            if (status.requiresTestCases && AppConfig.Configuration.validation && AppConfig.Configuration.validation.testCases) {
                // Show a loading dialog
                LoadingDialog.setHeader("Validating the Technical Review");
                LoadingDialog.setBody("This dialog will close after the validation completes.");
                LoadingDialog.show();

                // Get the assessment
                this.getAssessmentItem(item, true).then(item => {
                    // Ensure the item exists
                    if (item) {
                        // Parse the validation
                        for (let fieldName in AppConfig.Configuration.validation.testCases) {
                            // Ensure values exist
                            let values = AppConfig.Configuration.validation.testCases[fieldName];
                            if (values && values.length > 0) {
                                let matchFl = false;

                                // Parse the valid values
                                for (let i = 0; i < values.length; i++) {
                                    // See if the value matches
                                    if (values[i] == item[fieldName]) {
                                        // Set the flag and break from the loop
                                        matchFl = true;
                                        break;
                                    }
                                }

                                // See if a match wasn't found
                                if (!matchFl) {
                                    // Form isn't valid
                                    isValid = false;
                                    break;
                                }
                            }
                        }

                        // Hide the dialog
                        LoadingDialog.hide();

                        // See if it's valid
                        if (isValid) {
                            // Update the date completed field
                            item.update({ Completed: (new Date(Date.now())).toISOString() }).execute(() => {
                                // Resolve the request
                                resolve(isValid);
                            });
                        } else {
                            // Clear the modal
                            Modal.clear();

                            // Set the header
                            Modal.setHeader("Error");

                            // Set the body
                            Modal.setBody("The test cases are not valid. Please review them and try your request again.");

                            // Show the dialog
                            Modal.show();

                            // Resolve the request
                            resolve(isValid);
                        }
                    } else {
                        // Hide the dialog
                        LoadingDialog.hide();

                        // Clear the modal
                        Modal.clear();

                        // Set the header
                        Modal.setHeader("Error");

                        // Set the body
                        Modal.setBody("No test cases were found for this application.");

                        // Show the dialog
                        Modal.show();

                        // Not valid
                        resolve(false);
                    }
                });
            } else {
                // Resolve the request
                resolve(isValid);
            }
        });
    }

    // Last Assessment form
    lastAssessment(item: IAppItem) {
        // Get the assessment item
        this.getAssessmentItem(item, false).then(assessment => {
            // See if an item exists
            if (assessment) {
                // Display the view form
                DataSource.AppAssessments.viewForm({
                    itemId: assessment.Id,
                    webUrl: Strings.SourceUrl
                });
            }
            else {
                // Clear the modal
                Modal.clear();

                // Set the header & body
                Modal.setHeader("Assessment not found")
                Modal.setBody("Unable to find an assessment for app '" + DataSource.AppItem.Title + "'.")

                // Add the footer button
                Components.Button({
                    el: Modal.FooterElement,
                    text: "Close",
                    type: Components.ButtonTypes.OutlineSecondary,
                    onClick: () => {
                        Modal.hide();
                    }
                });

                // Show the modal
                Modal.show();
            }
        });
    }

    // Reject form
    reject(item: IAppItem, onUpdate: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Reject App");

        // Set the body
        Modal.setBody("Are you sure you want to send this back to the developer?");

        // Create the form
        let form = Components.Form({
            el: Modal.BodyElement,
            controls: [
                {
                    name: "Reason",
                    label: "Reason",
                    errorMessage: "A reason is required to reject the app request.",
                    required: true,
                    type: Components.FormControlTypes.TextArea
                }
            ]
        });

        // Set the footer
        Modal.setFooter(Components.Button({
            text: "Reject",
            type: Components.ButtonTypes.OutlineDanger,
            onClick: () => {
                // Ensure the form is valid
                if (form.isValid()) {
                    let comments = form.getValues()["Reason"];

                    // Close the modal
                    Modal.hide();

                    // Show a loading dialog
                    LoadingDialog.setHeader("Rejecting Request");
                    LoadingDialog.setBody("This dialog will close after the app is sent back to the developer.");
                    LoadingDialog.show();

                    // Update the status
                    let status = AppConfig.Status[item.AppStatus];
                    let values = {
                        AppComments: comments,
                        AppIsRejected: true,
                        AppStatus: status && status.prevStep ? status.prevStep : item.AppStatus
                    };
                    item.update(values).execute(
                        () => {
                            // Send the notification
                            AppNotifications.rejectEmail(item, comments).then(() => {
                                // Call the update event
                                onUpdate();

                                // Hide the dialog
                                LoadingDialog.hide();
                            });

                            // Log
                            DataSource.logItem({
                                LogUserId: ContextInfo.userId,
                                ParentId: item.AppProductID,
                                ParentListName: Strings.Lists.Apps,
                                Title: DataSource.AuditLogStates.AppRejected,
                                LogComment: `The app ${item.Title} was rejected. Justification: ${values.AppComments}`
                            }, { ...item, ...values });
                        },
                        ex => {
                            // Log the error
                            ErrorDialog.show("Updating Request", "There was an error updating the item.", ex);
                        }
                    );
                }
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Displays the list form for requesting an app catalog
    requestAppCatalog(url: string) {
        // Display the new form
        DataSource.AppCatalogRequests.newForm({
            webUrl: Strings.SourceUrl,
            onSetHeader(el) {
                el.innerHTML = el.innerHTML;
            },
            onCreateEditForm: props => {
                props.onSetFieldDefaultValue = (field, value) => {
                    // See if this is the url field
                    if (field.InternalName == "SiteCollectionUrl") {
                        // Set the default value
                        return {
                            Description: url,
                            Url: url
                        };
                    }
                    // Else, see if this is the user field
                    else if (field.InternalName == "Requesters") {
                        // Set the default value to the current user
                        return [ContextInfo.userId];
                    }

                    // Return the default value
                    return value;
                }

                // Return the properties
                return props;
            },
            onSave: (values) => {
                // Set the default status
                values["RequestStatus"] = "New";

                // Return the values
                return values;
            },
            onUpdate: (item: IAppCatalogRequestItem) => {
                // Reload the app catalog items
                DataSource.AppCatalogRequests.refresh();

                // Log
                DataSource.logItem({
                    LogUserId: ContextInfo.userId,
                    ParentId: DataSource.AppItem.AppProductID,
                    ParentListName: Strings.Lists.Apps,
                    Title: DataSource.AuditLogStates.AppCatalogRequest,
                    LogComment: `An app catalog was requested for site: ${item.SiteCollectionUrl.Url}`
                }, item as any);
            }
        });
    }

    // Retracts the solution from the tenant app catalog
    removeFromTenant(item: IAppItem, onUpdate: () => void) {
        // Set the header
        Modal.setHeader("Remove App");

        // Set the body
        Modal.setBody("Are you sure you want to remove the app from the tenant app catalog?");

        // Render the footer
        Modal.setFooter(Components.Button({
            text: "Remove",
            type: Components.ButtonTypes.OutlineDanger,
            onClick: () => {
                // Close the modal
                Modal.hide();

                // Retract the app
                AppActions.retract(true, true, () => {
                    // Log
                    DataSource.logItem({
                        LogUserId: ContextInfo.userId,
                        ParentId: item.AppProductID,
                        ParentListName: Strings.Lists.Apps,
                        Title: DataSource.AuditLogStates.AppRetracted,
                        LogComment: `The app ${item.Title} was removed from the tenant app catalog.`
                    }, item);

                    // Call the update event
                    onUpdate();
                });
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Retracts the solution from the tenant app catalog
    retractFromTenant(item: IAppItem, onUpdate: () => void) {
        // Set the header
        Modal.setHeader("Retract App");

        // Set the body
        Modal.setBody("Are you sure you want to retract the app from the tenant app catalog? The app will still be in the app catalog, but not available for use.");

        // Render the footer
        Modal.setFooter(Components.Button({
            text: "Retract",
            type: Components.ButtonTypes.OutlineDanger,
            onClick: () => {
                // Close the modal
                Modal.hide();

                // Retract the app
                AppActions.retract(true, false, () => {
                    // Log
                    DataSource.logItem({
                        LogUserId: ContextInfo.userId,
                        ParentId: item.AppProductID,
                        ParentListName: Strings.Lists.Apps,
                        Title: DataSource.AuditLogStates.AppRetracted,
                        LogComment: `The app ${item.Title} was retracted from the tenant app catalog.`
                    }, item);

                    // Call the update event
                    onUpdate();
                });
            }
        }).el);

        // Show the modal
        Modal.show();
    }


    // Form to send a notification
    sendNotification(item: IAppItem, to?: string[], subject?: string, body?: string) {
        // Set the header
        Modal.setHeader("Send Notification");

        // See if the to value exists
        let sendToValue = [];
        if (to && to.length > 0) {
            // Parse the options
            for (let i = 0; i < to.length; i++) {
                // Copy the value
                switch (to[i]) {
                    case "ApproversGroup":
                        // Add the value
                        sendToValue.push("Approver's Group");
                        break;
                    case "Developers":
                        // Add the value
                        sendToValue.push("App Developers");
                        break;
                    case "DevelopersGroup":
                        // Add the value
                        sendToValue.push("Developer's Group");
                        break;
                    case "Sponsor":
                        // Add the value
                        sendToValue.push("App Sponsor");
                        break;
                    case "SponsorsGroup":
                        // Add the value
                        sendToValue.push("Sponsor's Group");
                        break;
                }
            }
        }

        // Create a form
        let form = Components.Form({
            controls: [
                {
                    name: "UserTypes",
                    label: "Send To",
                    type: Components.FormControlTypes.MultiSwitch,
                    required: true,
                    value: sendToValue,
                    items: [
                        {
                            label: "Approver's Group",
                            data: "ApproversGroup"
                        },
                        {
                            label: "App Developers",
                            data: "Developers"
                        },
                        {
                            label: "Developer's Group",
                            data: "DevelopersGroup"
                        },
                        {
                            label: "App Sponsor",
                            data: "Sponsor"
                        },
                        {
                            label: "Sponsor's Group",
                            data: "SponsorsGroup"
                        }
                    ]
                } as Components.IFormControlPropsMultiCheckbox,
                {
                    name: "EmailSubject",
                    label: "Email Subject",
                    type: Components.FormControlTypes.TextField,
                    required: true,
                    value: subject
                },
                {
                    name: "EmailBody",
                    label: "Email Body",
                    type: Components.FormControlTypes.TextArea,
                    required: true,
                    rows: 10,
                    value: body
                } as Components.IFormControlPropsTextField
            ]
        });

        // Set the body
        Modal.setBody(form.el);

        // Render the footer
        Modal.setFooter(Components.Button({
            text: "Send",
            type: Components.ButtonTypes.OutlineSuccess,
            onClick: () => {
                // Ensure the form is valid
                if (form.isValid()) {
                    let values = form.getValues();

                    // Parse the user types
                    let userTypes = [];
                    for (let i = 0; i < values["UserTypes"].length; i++) {
                        let cbValue = values["UserTypes"][i] as Components.ICheckboxGroupItem;

                        // Append the value
                        userTypes.push(cbValue.data);
                    }

                    // Send an email
                    AppNotifications.sendNotification(item, userTypes, values["EmailSubject"], values["EmailBody"]).then(() => {
                        // Close the modal
                        Modal.hide();
                    });
                }
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Submit Form
    submit(item: IAppItem, onUpdate: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader((item.AppIsRejected ? "Resubmit" : "Submit") + " App for Approval");

        // Create the body element
        let elBody = document.createElement("div");

        // Append the content element
        let elContent = document.createElement("p");
        elBody.appendChild(elContent);

        // Get the status information
        let status = AppConfig.Status[item.AppStatus];

        // See if a checklist exists
        let checklist = status.checklists || [];
        if (checklist.length > 0) {
            let elChecklist = document.createElement("ul");
            elBody.appendChild(elChecklist);

            // Set the content
            elContent.innerHTML = "Review the following checklist before submitting the app for approval.";

            // Parse the items
            for (let i = 0; i < checklist.length; i++) {
                // Append the item
                let elItem = document.createElement("li");
                elItem.innerHTML = checklist[i];
                elChecklist.appendChild(elItem);
            }
        } else {
            // Set the content
            elContent.innerHTML = "Are you sure you want to submit this app for approval?";
        }

        // Set the body
        Modal.setBody(elBody);

        // Set the footer
        Modal.setFooter(Components.Button({
            text: item.AppIsRejected ? "Resubmit" : "Submit",
            type: Components.ButtonTypes.OutlinePrimary,
            onClick: () => {
                // Close the modal
                Modal.hide();

                // Ensure the metadata is complete
                this.isMetadataComplete(item).then(isValid => {
                    // Ensure it's valid
                    if (!isValid) { return; }

                    // See if the tech review is valid
                    this.isValidTechReview(item, status).then(isValid => {
                        // Ensure it's valid
                        if (!isValid) { return; }

                        // See if the test cases are valid
                        this.isValidTestCases(item, status).then(isValid => {
                            // Ensure it's valid
                            if (!isValid) { return; }

                            // Show a loading dialog
                            LoadingDialog.setHeader("Updating App Submission");
                            LoadingDialog.setBody("This dialog will close after the app submission is completed.");
                            LoadingDialog.show();

                            // Update the status
                            let status = AppConfig.Status[item.AppStatus];
                            let newStatus = status && status.nextStep ? status.nextStep : item.AppStatus;
                            let values = {
                                AppIsRejected: false,
                                AppStatus: newStatus
                            };
                            item.update(values).execute(() => {
                                // Code to run after the sponsor is added to the security group
                                let onComplete = () => {
                                    // Log
                                    DataSource.logItem({
                                        LogUserId: ContextInfo.userId,
                                        ParentId: item.AppProductID,
                                        ParentListName: Strings.Lists.Apps,
                                        Title: item.AppIsRejected ? DataSource.AuditLogStates.AppResubmitted : DataSource.AuditLogStates.AppSubmitted,
                                        LogComment: `The app ${item.Title} was ${item.AppIsRejected ? "resubmitted" : "submitted"} for approval.`
                                    }, { ...item, ...values });

                                    // Run the flow associated for this status
                                    AppActions.runFlow(status.flowId);

                                    // Send the notifications
                                    AppNotifications.sendEmail(status.notification, item, false).then(() => {
                                        // Call the update event
                                        onUpdate();

                                        // Hide the dialog
                                        LoadingDialog.hide();
                                    });
                                }

                                // Log
                                ErrorDialog.logInfo(`Validating the app sponsor id: '${item.AppSponsorId}'`);

                                // Get the sponsor
                                let sponsor = AppSecurity.AppWeb.getUserForGroup(Strings.Groups.Sponsors, item.AppSponsorId);
                                if (sponsor == null && item.AppSponsorId > 0) {
                                    // Log
                                    ErrorDialog.logInfo(`App sponsor not in group. Adding the user...`);

                                    // Add the sponsor to the group
                                    AppSecurity.AppWeb.addUserToGroup(Strings.Groups.Sponsors, item.AppSponsorId).then(() => {
                                        // Complete the request
                                        onComplete();
                                    });
                                } else {
                                    // Complete the request
                                    onComplete();
                                }
                            });
                        });
                    });
                });
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Updates the app
    updateApp(item: IAppItem, siteUrl: string, onUpdate: () => void) {
        // Set the header
        Modal.setHeader("Update App");

        // Set the body
        Modal.setBody("Are you sure you want to update the app?");

        // Render the footer
        Modal.setFooter(Components.Button({
            text: "Update",
            type: Components.ButtonTypes.OutlineSuccess,
            onClick: () => {
                // Close the modal
                Modal.hide();

                // Update the app
                AppActions.updateApp(siteUrl, true, onUpdate).then(() => {
                    // Send the notifications
                    AppNotifications.sendAppTestSiteUpgradedEmail(DataSource.AppItem).then(() => {
                        // Call the update event
                        onUpdate();
                    });

                    // Log
                    DataSource.logItem({
                        LogUserId: ContextInfo.userId,
                        ParentId: item.AppProductID,
                        ParentListName: Strings.Lists.Apps,
                        Title: DataSource.AuditLogStates.AppUpdated,
                        LogComment: `The app ${item.Title} was updated for site: ${siteUrl}`
                    }, item);
                });
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Displays the ugprade from
    upgrade(appItem: IAppItem) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Upgrade App to v" + appItem.AppVersion);

        // Display a loading dialog
        LoadingDialog.setHeader("Loading Deployment Information");
        LoadingDialog.setBody("Gathering the app deployment information...");
        LoadingDialog.show();

        // Parse the site collection urls and generate the check group items
        let items: Components.ICheckboxGroupItem[] = [];
        let urls = (appItem.AppSiteDeployments || "").split('\n');
        Helper.Executor(urls, url => {
            // Return a promise
            return new Promise(resolve => {
                // Ensure the url exists
                if (url) {
                    // Load the web containing this app
                    Web(url).SiteCollectionAppCatalog().AvailableApps(appItem.AppProductID).execute(
                        // Exists
                        (app) => {
                            let appVersion = app.InstalledVersion || app.AppCatalogVersion;

                            // Append the item
                            items.push({
                                data: url,
                                label: url + " (v" + appVersion + ")",
                                type: Components.CheckboxGroupTypes.Switch,
                                isDisabled: appItem.AppVersion == appVersion
                            });

                            // Check the next web
                            resolve(null);
                        },

                        // Denied Access
                        () => {
                            // Append the item
                            items.push({
                                label: url,
                                type: Components.CheckboxGroupTypes.Switch,
                                isDisabled: true
                            });

                            // Check the next web
                            resolve(null);
                        }
                    );
                } else {
                    // Check the next web
                    resolve(null);
                }
            });
        }).then(() => {
            // Hide the loading dialog
            LoadingDialog.hide();

            // Render a form
            let form = Components.Form({
                el: Modal.BodyElement,
                controls: [
                    {
                        items,
                        name: "urls",
                        required: true,
                        type: Components.FormControlTypes.MultiCheckbox,
                        errorMessage: "Please select one or more site collections to upgrade."
                    } as Components.IFormControlPropsMultiCheckbox
                ]
            });

            // Render the footer
            Components.Button({
                el: Modal.FooterElement,
                text: "Upgrade",
                type: Components.ButtonTypes.OutlineSuccess,
                onClick: () => {
                    // Validate the form
                    if (form.isValid()) {
                        let counter = 0;
                        let items = form.getValues()["urls"] as Components.ICheckboxGroupItem[];

                        // Close the modal
                        Modal.hide();

                        // Show a loading dialog
                        LoadingDialog.setHeader("Upgrading Apps");
                        LoadingDialog.setBody("Upgrading " + (++counter) + " of " + items.length);
                        LoadingDialog.show();

                        // Parse the items
                        Helper.Executor(items, item => {
                            // Return a promise
                            return new Promise(resolve => {
                                // Upgrade the app
                                AppActions.updateApp(item.data, false, () => { resolve(null); }).then(() => {
                                    // Update the loading dialog
                                    LoadingDialog.setHeader("Upgrading Apps");
                                    LoadingDialog.setBody("Upgrading " + (++counter) + " of " + items.length);

                                    // Log
                                    DataSource.logItem({
                                        LogUserId: ContextInfo.userId,
                                        ParentId: appItem.AppProductID,
                                        ParentListName: Strings.Lists.Apps,
                                        Title: DataSource.AuditLogStates.AppUpgraded,
                                        LogComment: `The app ${appItem.Title} was upgraded on site: ${item.data}`
                                    }, appItem);

                                    // Upgrade the next site
                                    resolve(null);
                                });
                            });
                        }).then(() => {
                            // Close the dialog
                            LoadingDialog.hide();
                        });
                    }
                }
            });

            // Show the modal
            Modal.show();
        });
    }
}