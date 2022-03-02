import { ItemForm, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, List, SPTypes, Types, Web } from "gd-sprest-bs";
import { AppActions } from "./appActions";
import { AppConfig, IStatus } from "./appCfg";
import { AppNotifications } from "./appNotifications";
import { DataSource, IAppItem, IAssessmentItem } from "./ds";
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
                                    }).execute(() => {
                                        // Close the dialog
                                        LoadingDialog.hide();

                                        // Send the notification
                                        AppNotifications.sendEmail(status.nextStep, item).then(onUpdate);
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

                // Create the test site
                AppActions.createTestSite(item, onUpdate);
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
                this.retract(item, false, true, () => {
                    // Retract the solution from the tenant app catalog
                    this.retract(item, true, true, () => {
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
            }).execute(items => {
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
            }, reject);
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
        // Set the header
        Modal.setHeader("Deploy to " + (tenantFl ? " Tenant " : " Site Collection ") + "App Catalog");

        // Set the body
        Modal.setBody("Are you sure you want to deploy the application?");

        // Render the footer
        Modal.setFooter(Components.Button({
            text: "Deploy",
            type: Components.ButtonTypes.OutlineSuccess,
            onClick: () => {
                // Close the modal
                Modal.hide();

                // Deploy the app
                AppActions.deploy(item, tenantFl, onUpdate);
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Deploys the solution to a site collection app catalog
    deployToSite(item: IAppItem, onUpdate: () => void) {
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
                            // Disable the deploy button
                            btnDeploy.disable();
                        }
                    },
                    onValidate: (ctrl, results) => {
                        // Ensure a value exists
                        if (results.value) {
                            // Set the flag
                            results.isValid = hasAppCatalog;

                            // Set the error message
                            results.invalidMessage = "The site does not contain an app catalog.";
                        } else {
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
        Modal.setFooter(Components.TooltipGroup({
            tooltips: [
                {
                    content: "Loads the site collection and validates the app catalog.",
                    btnProps: {
                        text: "Load Site",
                        type: Components.ButtonTypes.OutlinePrimary,
                        onClick: () => {
                            // Validate the form
                            let url = form.getValues()["Url"];
                            if (url) {
                                // Show a loading dialog
                                LoadingDialog.setHeader("Loading App Catalog");
                                LoadingDialog.setBody("This dialog will close after the app catalog is loaded.");
                                LoadingDialog.show();

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
                                )
                            } else {
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
                            AppActions.deployToTeams(item, onUpdate);
                        }
                    }
                }
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
                AppActions.deployToTeams(item, onUpdate);
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
                }).execute(item => {
                    // Show the assessment form
                    displayForm(item as any);
                });
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
                    }).execute(item => {
                        // Show the assessment form
                        displayForm(item as any);
                    });
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
                    }).execute(() => {
                        // Send the notification
                        AppNotifications.rejectEmail(newStatus, item, comments).then(onUpdate);
                    });
                }
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Retracts the solution to the app catalog
    retract(item: IAppItem, tenantFl: boolean, removeFl: boolean, onUpdate: () => void) {
        // Show a loading dialog
        LoadingDialog.setHeader("Retracting the Package");
        LoadingDialog.setBody("Retracting the spfx package to the app catalog.");
        LoadingDialog.show();

        // Load the context of the app catalog
        let catalogUrl = tenantFl ? AppConfig.Configuration.tenantAppCatalogUrl : AppConfig.Configuration.appCatalogUrl;
        ContextInfo.getWeb(catalogUrl).execute(context => {
            // Load the apps
            let web = Web(catalogUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue });
            (tenantFl ? web.TenantAppCatalog() : web.SiteCollectionAppCatalog()).AvailableApps(item.AppProductID).retract().execute(() => {
                // See if we are removing the app
                if (removeFl) {
                    // Remove the app
                    (tenantFl ? web.TenantAppCatalog() : web.SiteCollectionAppCatalog()).AvailableApps(item.AppProductID).remove().execute(() => {
                        // Call the update event
                        onUpdate();

                        // Close the dialog
                        LoadingDialog.hide();
                    });
                } else {
                    // Call the update event
                    onUpdate();

                    // Close the dialog
                    LoadingDialog.hide();
                }
            }, () => {
                // This shouldn't happen
                // The app was already checked to be deployed

                // Call the update event
                onUpdate();

                // Close the dialog
                LoadingDialog.hide();
            });
        });
    }

    // Submit Form
    submit(item: IAppItem, onUpdate: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Submit App for Approval");

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
                                    AppStatus: newStatus
                                }).execute(() => {
                                    // Send an email
                                    AppNotifications.sendEmail(newStatus, item).then(onUpdate);
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
}