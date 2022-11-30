import { ItemForm, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, List, SPTypes, Web } from "gd-sprest-bs";
import { AppActions } from "./appActions";
import { AppConfig, IStatus } from "./appCfg";
import { AppNotifications } from "./appNotifications";
import { AppSecurity } from "./appSecurity";
import { DataSource, IAppItem, IAssessmentItem } from "./ds";
import { ErrorDialog } from "./errorDialog";
import Strings from "./strings";

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

                    // See if the tech review is valid
                    this.isValidTechReview(item, status).then(isValid => {
                        // Ensure it's valid
                        if (isValid) {
                            // See if the test cases are valid
                            this.isValidTestCases(item, status).then(isValid => {
                                // Ensure it's valid
                                if (isValid) {
                                    // Show a loading dialog
                                    LoadingDialog.setHeader("Updating the Status");
                                    LoadingDialog.setBody("This dialog will close after the status is updated.");
                                    LoadingDialog.show();

                                    // Update the item
                                    item.update({
                                        AppStatus: status.nextStep
                                    }).execute(
                                        () => {
                                            // Close the dialog
                                            LoadingDialog.hide();

                                            // Send the notifications
                                            AppNotifications.sendEmail(status.notification, item).then(() => {
                                                // See if the test app catalog exists and we are creating the test site
                                                if (DataSource.SiteCollectionAppCatalogExists && status.createTestSite) {
                                                    // Load the test site
                                                    DataSource.loadTestSite(item).then(
                                                        // Exists
                                                        webInfo => {
                                                            // See if the current version is deployed
                                                            if (item.AppVersion == webInfo.app.InstalledVersion && !webInfo.app.SkipDeploymentFeature) {
                                                                // Call the update event
                                                                onUpdate();
                                                            } else {
                                                                // Update the app
                                                                AppActions.updateApp(item, webInfo.web.ServerRelativeUrl, true, onUpdate).then(() => {
                                                                    // Call the update event
                                                                    onUpdate();
                                                                });
                                                            }
                                                        },

                                                        // Doesn't exist
                                                        () => {
                                                            // Create the test site
                                                            AppActions.createTestSite(item, onUpdate);
                                                        }
                                                    );
                                                } else {
                                                    // Call the update event
                                                    onUpdate();
                                                }
                                            },
                                                ex => {
                                                    // Log the error
                                                    ErrorDialog.show("Updating Status", "There was an error updating the status.", ex);
                                                }
                                            );
                                        });

                                }
                            });
                        }
                    });
                }
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Creates the test site for the application
    createTestSite(item: IAppItem, onUpdate: () => void) {
        // Set the header
        Modal.setHeader("Create Test Site");

        // Set the body
        Modal.setBody("Are you sure you want to create the test site for this application?");

        // Render the footer
        Modal.setFooter(Components.Button({
            text: "Create Site",
            type: Components.ButtonTypes.OutlineSuccess,
            onClick: () => {
                // Close the modal
                Modal.hide();

                // Code to run after the logic below completes
                let onComplete = () => {
                    // Create the test site
                    AppActions.createTestSite(item, () => {
                        // Call the update event
                        onUpdate();
                    });
                }

                // See if the app was deployed, but errored out
                if (DataSource.DocSetSCAppItem && DataSource.DocSetSCAppItem.AppPackageErrorMessage && DataSource.DocSetSCAppItem.AppPackageErrorMessage != "No errors.") {
                    // Delete the item
                    DataSource.DocSetItem.delete().execute(() => {
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
    delete(item: IAppItem, onUpdate: () => void) {
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
                AppActions.retract(item, false, true, () => {
                    // Retract the solution from the tenant app catalog
                    AppActions.retract(item, true, true, () => {
                        // Delete the test site
                        AppActions.deleteTestSite(item).then(() => {
                            // Update the loading dialog
                            LoadingDialog.setHeader("Removing Assessments");
                            LoadingDialog.setBody("Removing the assessments associated with this app.");

                            // Delete the assessments w/ this app
                            this.deleteAssessments(item).then(() => {
                                // Update the loading dialog
                                LoadingDialog.setHeader("Deleting the App Request");
                                LoadingDialog.setBody("This dialog will close after the app request is deleted.");

                                // Delete this folder
                                item.delete().execute(() => {
                                    // Close the dialog
                                    LoadingDialog.hide();

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
    private deleteAssessments(item: IAppItem): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the assoicated item
            List(Strings.Lists.Assessments).Items().query({
                Filter: "RelatedAppId eq " + item.Id
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
                AppActions.deleteTestSite(item).then(() => {
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
                AppActions.deploy(item, tenantFl, skipFeatureDeployment, onUpdate, () => {
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
        let loadError = null;
        let isOwner = null;
        let isRoot = null;
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
                        // Disable the request button by default
                        btnRequest.disable();

                        // See if there is an error loading the site
                        if (loadError) {
                            // Set the flag
                            results.isValid = false;

                            // Set the error message
                            results.invalidMessage = "Unable to load the site. Please check the url and try again.<br/>Note - Please ensure the url is relative or absolute and does not include a page. (Example: " + window.location.origin + "/sites/demo)";
                        }
                        // Else, see if the site is not the root web
                        else if (isRoot === false) {
                            // Set the flag
                            results.isValid = false;

                            // Set the error message
                            results.invalidMessage = "The site entered is a subsite. Please enter the url for the root web of a site collection.";
                        }
                        // Else, see if the user is not an owner
                        else if (isOwner === false) {
                            // Set the flag
                            results.isValid = false;

                            // Set the error message
                            results.invalidMessage = "You do not have the appropriate rights to deploy to this site collection.";
                        }
                        // Else, ensure a value exists
                        else if (results.value) {
                            // Set the flag
                            results.isValid = hasAppCatalog;

                            // Set the error message
                            results.invalidMessage = "The site does not contain an app catalog.";

                            // Enable the button
                            btnRequest.enable();
                        }
                        else {
                            // Set the flag
                            results.isValid = false;

                            // Set the error message
                            results.invalidMessage = "Please enter the relative site collection url to the app catalog.";
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
        let hasAppCatalog: boolean = false;
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
                            // Clear the owner flag and url
                            isOwner = null;
                            isRoot = null;
                            loadError = null;
                            webUrl = null;

                            // Validate the form
                            let url = form.getValues()["Url"];
                            if (url) {
                                // Show a loading dialog
                                LoadingDialog.setHeader("Loading App Catalog");
                                LoadingDialog.setBody("This dialog will close after the app catalog is loaded.");
                                LoadingDialog.show();

                                // Ensure the user is an owner of the site
                                DataSource.isOwner(url).then(
                                    // Successfully loaded site
                                    info => {
                                        // Set the flag and url
                                        isOwner = info.isOwner;
                                        isRoot = info.isRoot;
                                        webUrl = info.url;

                                        // See if the user has the rights to deploy the solution
                                        if (isOwner) {
                                            // Load the app catalog
                                            hasAppCatalog = false;
                                            Web(url).SiteCollectionAppCatalog().AvailableApps().execute(
                                                // Success
                                                apps => {
                                                    // Set the flag
                                                    hasAppCatalog = apps.results && typeof (apps.results.length) === "number" ? true : false;

                                                    // Validate the form and disable/enable the deploy button
                                                    form.isValid() ? btnDeploy.enable() : btnDeploy.disable();

                                                    // Hide the dialog
                                                    LoadingDialog.hide();
                                                },

                                                // The app catalog doesn't exist
                                                () => {
                                                    // Validate the form
                                                    form.isValid();

                                                    // Hide the dialog
                                                    LoadingDialog.hide();
                                                }
                                            );
                                        } else {
                                            // Show the error message
                                            form.isValid();

                                            // Hide the dialog
                                            LoadingDialog.hide();
                                        }
                                    },
                                    // Error loading site
                                    () => {
                                        // Set the flag
                                        loadError = true;

                                        // Validate the form
                                        form.isValid();

                                        // Hide the dialog
                                        LoadingDialog.hide();
                                    }
                                );
                            } else {
                                // Show the default error message
                                form.isValid();

                                // Hide the dialog
                                LoadingDialog.hide();
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
                            AppActions.deployToSite(item, form.getValues()["Url"], () => {
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
                            // Display the notification modal
                            this.sendNotification(item, AppConfig.Configuration.appCatalogRequests, "Request for Site Collection App Catalog",
                                `App Catalog Admins,

We are requesting an app catalog to be created on the following site(s):

${webUrl}

r/,
${ContextInfo.userDisplayName}`.trim());
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
                AppActions.deployToTeams(item, () => {
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
        // Set the form properties
        ItemForm.AutoClose = false;
        ItemForm.ListName = Strings.Lists.Apps;

        // Show the item form
        ItemForm.view({
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
            onCreateViewForm: props => {
                // Exclude fields
                props.excludeFields = ["AppVideoURL", "AppStatus", "IsAppPackageEnabled", "IsDefaultAppMetadataLocale"];

                // See if the item is not approved
                if ((props.info.item as IAppItem).AppStatus != "Approved") {
                    // Exclude the site deployments field
                    props.excludeFields.push("AppSiteDeployments");
                }

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
        // Set the form properties
        ItemForm.AutoClose = false;
        ItemForm.ListName = Strings.Lists.Assessments;

        // Get the assessment item
        this.getAssessmentItem(item, false).then(assessment => {
            // See if an item exists
            if (assessment) {
                // Show the view form
                ItemForm.view({
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
        // Set the form properties
        ItemForm.AutoClose = false;
        ItemForm.ListName = Strings.Lists.Assessments;

        // Get the assessment item
        this.getAssessmentItem(item).then(assessment => {
            // See if an item exists
            if (assessment) {
                // Show the view form
                ItemForm.view({
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
        // Set the form properties
        ItemForm.AutoClose = false;
        ItemForm.ListName = Strings.Lists.Apps;

        // Show the item form
        ItemForm.edit({
            itemId,
            webUrl: Strings.SourceUrl,
            onSetFooter: el => {
                // Add a cancel button if form is in a modal
                if (ItemForm.UseModal) {
                    Components.Button({
                        el,
                        text: "Cancel",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            Modal.hide();
                        }
                    });
                }
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
            onCreateEditForm: props => {
                // Exclude fields
                props.excludeFields = ["AppVideoURL", "AppStatus", "IsAppPackageEnabled", "IsDefaultAppMetadataLocale"];

                // Update the field
                props.onControlRendering = (ctrl, field) => {
                    // See if this is a read-only field
                    if (["AppAPIPermissions", "AppIsClientSideSolution", "AppIsDomainIsolated",
                        "AppProductID", "AppSharePointMinVersion", "AppSkipFeatureDeployment",
                        "AppVersion", "FileLeafRef", "Title"].indexOf(field.InternalName) >= 0) {
                        // Make it read-only
                        ctrl.isReadonly = true;
                    }

                    // See if this is the permissions justification
                    if (field.InternalName == "AppPermissionsJustification") {
                        // Add validation
                        ctrl.onValidate = (ctrl, results) => {
                            // See if permissions exist
                            let apiPermissions = ItemForm.EditForm.getItem()["AppAPIPermissions"];
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

                    // See if this is a url field
                    if (field.InternalName.indexOf("URL") > 0) {
                        // Hide the description field
                        (ctrl as Components.IFormControlUrlProps).showDescription = false;
                    }
                }

                // Return the properties
                return props;
            },
            onUpdate: (item: IAppItem) => {
                // See if this is a new item
                if (item.AppStatus == "New") {
                    // Log
                    ErrorDialog.logInfo(`Validating the app sponsor id: '${item.AppSponsorId}'`);

                    // Get the sponsor
                    let sponsor = AppSecurity.getSponsor(item.AppSponsorId);
                    if (sponsor == null && item.AppSponsorId > 0) {
                        // Log
                        ErrorDialog.logInfo(`App sponsor not in group. Adding the user...`);

                        // Add the sponsor to the group
                        AppSecurity.addSponsor(item.AppSponsorId).then(() => {
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

    // Edit form displaying the app catalog metadata
    editAppData(itemId: number, onUpdate: () => void) {
        // Set the form properties
        ItemForm.AutoClose = false;
        ItemForm.ListName = Strings.Lists.Apps;

        // Show the item form
        ItemForm.edit({
            itemId,
            webUrl: Strings.SourceUrl,
            onSetFooter: el => {
                // Add a cancel button if form is in a modal
                if (ItemForm.UseModal) {
                    Components.Button({
                        el,
                        text: "Cancel",
                        type: Components.ButtonTypes.OutlineSecondary,
                        onClick: () => {
                            Modal.hide();
                        }
                    });
                }
            },
            onSetHeader: el => {
                // Update the header
                el.className = "h5 m-0";
                el.innerHTML = "App Catalog Details";
            },
            onGetListInfo: props => {
                // Update the query
                props.fields = [
                    "AppDescription", "AppThumbnailURL", "AppImageURL1", "AppImageURL2",
                    "AppImageURL3", "AppImageURL4", "AppImageURL5", "AppSupportURL",
                    "AppVideoURL"
                ];

                // Return the properties
                return props;
            },
            onCreateEditForm: props => {
                // Exclude fields
                props.includeFields = [
                    "AppDescription", "AppThumbnailURL", "AppImageURL1", "AppImageURL2",
                    "AppImageURL3", "AppImageURL4", "AppImageURL5", "AppSupportURL",
                    "AppVideoURL"
                ];

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
            },
            onUpdate: () => {
                // Call the update event
                onUpdate();
            }
        });
    }

    // Technical review form
    editTechReview(item: IAppItem, onUpdate: () => void) {
        // Set the form properties
        ItemForm.AutoClose = false;
        ItemForm.ListName = Strings.Lists.Assessments;

        // Displays the assessment form
        let displayForm = (assessment: IAssessmentItem) => {
            let validateFl = false;

            // Show the edit form
            ItemForm.edit({
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
                                ItemForm.EditForm.isValid();
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
                displayForm(assessment);
            } else {
                // Show a loading dialog
                LoadingDialog.setHeader("Creating the Form");
                LoadingDialog.setBody("This dialog will close after the form is created.");

                // Create the item
                Web(Strings.SourceUrl).Lists(Strings.Lists.Assessments).Items().add({
                    RelatedAppId: item.Id,
                    Title: item.Title + " Review " + (new Date(Date.now()).toDateString())
                }).execute(
                    item => {
                        // Show the assessment form
                        displayForm(item as any);
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
        // Set the form properties
        ItemForm.AutoClose = false;
        ItemForm.ListName = Strings.Lists.Assessments;

        // Displays the form
        let displayForm = (assessment: IAssessmentItem) => {
            let validateFl = false;

            // Show the edit form
            ItemForm.edit({
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
                                ItemForm.EditForm.isValid();
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
                displayForm(assessment);
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
                    Web(Strings.SourceUrl).Lists(Strings.Lists.Assessments).Items().add({
                        ContentTypeId: ct ? ct.StringId : null,
                        RelatedAppId: item.Id,
                        Title: item.Title + " Tests " + (new Date(Date.now()).toDateString())
                    }).execute(
                        item => {
                            // Show the assessment form
                            displayForm(item as any);
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
            // Get the assoicated item
            List(Strings.Lists.Assessments).Items().query({
                Filter: "RelatedAppId eq " + item.Id + " and ContentType eq '" + (testFl ? "TestCases" : "Item") + "'",
                OrderBy: ["Completed desc"]
            }).execute(items => {
                // Return the last item
                resolve(items.results[0] as any);
            }, reject);
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

                        // See if it's not valid
                        if (!isValid) {
                            // Clear the modal
                            Modal.clear();

                            // Set the header
                            Modal.setHeader("Error");

                            // Set the body
                            Modal.setBody("The technical review is not valid. Please review it and try your request again.");

                            // Wait for the loading dialog to complete it's closing event
                            setTimeout(() => {
                                // Show the dialog
                                Modal.show();
                            }, 250)
                        }

                        // Resolve the request
                        resolve(isValid);
                    } else {
                        // Hide the dialog
                        LoadingDialog.hide();

                        // Clear the modal
                        Modal.clear();

                        // Set the header
                        Modal.setHeader("Error");

                        // Set the body
                        Modal.setBody("No technical review exists for this application.");

                        // Wait for the loading dialog to complete it's closing event
                        setTimeout(() => {
                            // Show the dialog
                            Modal.show();
                        }, 250)

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

                        // See if it's not valid
                        if (!isValid) {
                            // Clear the modal
                            Modal.clear();

                            // Set the header
                            Modal.setHeader("Error");

                            // Set the body
                            Modal.setBody("The test cases are not valid. Please review them and try your request again.");

                            // Wait for the loading dialog to complete it's closing event
                            setTimeout(() => {
                                // Show the dialog
                                Modal.show();
                            }, 250)
                        }

                        // Resolve the request
                        resolve(isValid);
                    } else {
                        // Hide the dialog
                        LoadingDialog.hide();

                        // Clear the modal
                        Modal.clear();

                        // Set the header
                        Modal.setHeader("Error");

                        // Set the body
                        Modal.setBody("No test cases were found for this application.");

                        // Wait for the loading dialog to complete it's closing event
                        setTimeout(() => {
                            // Show the dialog
                            Modal.show();
                        }, 250)

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
        // Set the list name
        ItemForm.AutoClose = false;
        ItemForm.ListName = Strings.Lists.Assessments;

        // Get the assessment item
        this.getAssessmentItem(item, false).then(assessment => {
            // See if an item exists
            if (assessment) {
                // Show the edit form
                ItemForm.view({
                    itemId: assessment.Id,
                    webUrl: Strings.SourceUrl
                });
            }
            else {
                // Show 'assessment not found' modal
                Modal.clear();
                Modal.setHeader("Assessment not found")
                Modal.setBody("Unable to find an assessment for app '" + DataSource.DocSetItem.Title + "'.")
                let close = Components.Button({
                    el: document.createElement("div"),
                    text: "Close",
                    type: Components.ButtonTypes.OutlineSecondary,
                    onClick: () => {
                        Modal.hide();
                    }
                });
                Modal.setFooter(close.el);
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

                    // Get the current status configuration
                    let status = AppConfig.Status[item.AppStatus];

                    // Set the new status
                    let newStatus = status && status.prevStep ? status.prevStep : item.AppStatus;

                    // Update the status
                    item.update({
                        AppComments: comments,
                        AppIsRejected: true,
                        AppStatus: newStatus
                    }).execute(
                        () => {
                            // Send the notification
                            AppNotifications.rejectEmail(newStatus, item, comments).then(() => {
                                // Call the update event
                                onUpdate();

                                // Hide the dialog
                                LoadingDialog.hide();
                            });
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

    // Retracts the solution from the tenant app catalog
    retractFromTenant(item: IAppItem, onUpdate: () => void) {
        // Set the header
        Modal.setHeader("Deploy to Teams");

        // Set the body
        Modal.setBody("Are you sure you want to retract the solution?");

        // Render the footer
        Modal.setFooter(Components.Button({
            text: "Retract",
            type: Components.ButtonTypes.OutlineDanger,
            onClick: () => {
                // Close the modal
                Modal.hide();

                // Retract the app
                AppActions.retract(item, true, false, () => {
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

                // See if the tech review is valid
                this.isValidTechReview(item, status).then(isValid => {
                    // Ensure it's valid
                    if (isValid) {
                        // See if the test cases are valid
                        this.isValidTestCases(item, status).then(isValid => {
                            // Ensure it's valid
                            if (isValid) {
                                // Show a loading dialog
                                LoadingDialog.setHeader("Updating App Submission");
                                LoadingDialog.setBody("This dialog will close after the app submission is completed.");
                                LoadingDialog.show();

                                // Update the status
                                let status = AppConfig.Status[item.AppStatus];
                                let newStatus = status ? status.nextStep : AppConfig.Status[0].name;
                                item.update({
                                    AppIsRejected: false,
                                    AppStatus: item.AppIsRejected ? item.AppStatus : newStatus
                                }).execute(() => {
                                    // Code to run after the sponsor is added to the security group
                                    let onComplete = () => {
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
                                    let sponsor = AppSecurity.getSponsor(item.AppSponsorId);
                                    if (sponsor == null && item.AppSponsorId > 0) {
                                        // Log
                                        ErrorDialog.logInfo(`App sponsor not in group. Adding the user...`);

                                        // Add the sponsor to the group
                                        AppSecurity.addSponsor(item.AppSponsorId).then(() => {
                                            // Complete the request
                                            onComplete();
                                        });
                                    } else {
                                        // Complete the request
                                        onComplete();
                                    }
                                });
                            }
                        });
                    }
                });
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Updates the app
    updateApp(item: IAppItem, appCatalogUrl: string, siteUrl: string, onUpdate: () => void) {
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
                AppActions.updateApp(item, siteUrl, true, onUpdate).then(() => {
                    // Call the update event
                    onUpdate();
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
                                AppActions.updateApp(appItem, item.data, false, () => { resolve(null); }).then(() => {
                                    // Update the loading dialog
                                    LoadingDialog.setHeader("Upgrading Apps");
                                    LoadingDialog.setBody("Upgrading " + (++counter) + " of " + items.length);

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