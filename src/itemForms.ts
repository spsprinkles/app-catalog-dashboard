import { ItemForm, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, List, SPTypes, Types, Utility, Web } from "gd-sprest-bs";
import { loadAsync } from "jszip";
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
        let status = DataSource.Status[item.AppStatus];

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

                        // Call the update event
                        onUpdate();
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

                // Show a loading dialog
                LoadingDialog.setHeader("Deploying the Application");
                LoadingDialog.setBody("Adding the application to the test site collection app catalog.");
                LoadingDialog.show();

                // Deploy the solution
                this.deploy(item, false, () => {
                    // Update the loading dialog
                    LoadingDialog.setHeader("Creating the Test Site");
                    LoadingDialog.setBody("Creating the sub-web for testing the application.");
                    LoadingDialog.show();

                    // Load the context of the app catalog
                    ContextInfo.getWeb(DataSource.Configuration.appCatalogUrl).execute(context => {
                        let requestDigest = context.GetContextWebInformation.FormDigestValue;

                        // Create the test site
                        Web(DataSource.Configuration.appCatalogUrl, { requestDigest }).WebInfos().add({
                            Description: "The test site for the " + item.Title + " application.",
                            Title: item.Title,
                            Url: item.AppProductID,
                            WebTemplate: SPTypes.WebTemplateType.Site
                        }).execute(
                            // Success
                            web => {
                                // Update the loading dialog
                                LoadingDialog.setHeader("Installing the App");
                                LoadingDialog.setBody("Installing the application to the test site.");

                                // Install the application to the test site
                                Web(web.ServerRelativeUrl, { requestDigest }).SiteCollectionAppCatalog().AvailableApps(item.AppProductID).install().execute(
                                    // Success
                                    () => {
                                        // Update the loading dialog
                                        LoadingDialog.setHeader("Sending Email Notifications");
                                        LoadingDialog.setBody("Everything is done. Sending an email to the developer poc(s).");

                                        // Get the app developers
                                        let to = [];
                                        let pocs = item.AppDevelopers && item.AppDevelopers.results ? item.AppDevelopers.results : [];
                                        for (let i = 0; i < pocs.length; i++) {
                                            // Append the email
                                            to.push(pocs[i].EMail);
                                        }

                                        // Ensure emails exist
                                        if (to.length > 0) {
                                            // Send an email
                                            Utility().sendEmail({
                                                To: to,
                                                Subject: "Test Site Created",
                                                Body: "<p>App Developers,</p><br />" +
                                                    "<p>The '" + item.Title + "' app test site has been created.</p>"
                                            }).execute(() => {
                                                // Close the dialog
                                                LoadingDialog.hide();

                                                // Execute the update event
                                                onUpdate();
                                            });
                                        } else {
                                            // Close the dialog
                                            LoadingDialog.hide();

                                            // Execute the update event
                                            onUpdate();
                                        }
                                    },

                                    // Error
                                    () => {
                                        // Hide the loading dialog
                                        LoadingDialog.hide();

                                        // Error creating the site
                                        Modal.clear();
                                        Modal.setHeader("Error Deploying Application");
                                        Modal.setBody("There was an error deploying the application to the test site.");
                                        Modal.show();
                                    }
                                )
                            },
                            // Error
                            () => {
                                // Hide the loading dialog
                                LoadingDialog.hide();

                                // Error creating the site
                                Modal.clear();
                                Modal.setHeader("Error Creating Test Site");
                                Modal.setBody("There was an error creating the test site.");
                                Modal.show();
                            }
                        );
                    });
                });
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
                        // Update the loading dialog
                        LoadingDialog.setHeader("Deleting the Test Site");
                        LoadingDialog.setBody("Removing the test site created for this app.");
                        LoadingDialog.show();

                        // Delete the test site
                        DataSource.deleteTestSite(item).then(() => {
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

    // Deploys the solution to the app catalog
    deploy(item: IAppItem, tenantFl: boolean, onUpdate: () => void) {
        // Show a loading dialog
        LoadingDialog.setHeader("Uploading Package");
        LoadingDialog.setBody("Uploading the spfx package to the app catalog.");
        LoadingDialog.show();

        // Get the package file
        DataSource.DocSetItem.Folder().Files().execute(files => {
            let appFile = null;

            // Find the package file
            for (let i = 0; i < files.results.length; i++) {
                let file = files.results[i];

                // See if this is the package
                if (file.Name.endsWith(".sppkg")) {
                    // Set the file
                    appFile = file;
                    break;
                }
            }

            // Ensure a file exists
            if (appFile == null) {
                // This shouldn't happen
                LoadingDialog.hide();
                return;
            }

            // Upload the package file
            appFile.content().execute(content => {
                let catalogUrl = tenantFl ? DataSource.Configuration.tenantAppCatalogUrl : DataSource.Configuration.appCatalogUrl;

                // Load the context of the app catalog
                ContextInfo.getWeb(catalogUrl).execute(context => {
                    let requestDigest = context.GetContextWebInformation.FormDigestValue;

                    // Upload the file to the app catalog
                    let web = Web(catalogUrl, { requestDigest });
                    (tenantFl ? web.TenantAppCatalog() : web.SiteCollectionAppCatalog()).add(item.FileLeafRef, true, content).execute(file => {
                        // Update the dialog
                        LoadingDialog.setHeader("Deploying the Package");
                        LoadingDialog.setBody("This will close after the app is deployed.");

                        // Get the app item
                        file.ListItemAllFields().execute(appItem => {
                            // Update the metadata
                            this.updateApp(item, appItem.Id, tenantFl, catalogUrl, requestDigest).then(() => {
                                // Get the app catalog
                                let web = Web(catalogUrl, { requestDigest });
                                let appCatalog = (tenantFl ? web.TenantAppCatalog() : web.SiteCollectionAppCatalog());

                                // Deploy the app
                                appCatalog.AvailableApps(item.AppProductID).deploy().execute(app => {
                                    // Hide the dialog
                                    LoadingDialog.hide();

                                    // Call the update event
                                    onUpdate();
                                }, () => {
                                    // Hide the dialog
                                    LoadingDialog.hide();

                                    // Error deploying the app
                                    // TODO - Show an error
                                    // Call the update event
                                    onUpdate();
                                });
                            });
                        });
                    });
                });
            });
        });
    }

    // Deploys the solution to teams
    deployToTeams(item: IAppItem, onUpdate: () => void) {
        // Show a loading dialog
        LoadingDialog.setHeader("Deploying to Teams");
        LoadingDialog.setBody("Syncing the app to Teams.");
        LoadingDialog.show();

        // Load the context of the app catalog
        ContextInfo.getWeb(DataSource.Configuration.tenantAppCatalogUrl).execute(context => {
            let requestDigest = context.GetContextWebInformation.FormDigestValue;

            // Sync the app with Teams
            Web(DataSource.Configuration.tenantAppCatalogUrl, { requestDigest }).TenantAppCatalog().syncSolutionToTeams(item.Id).execute(
                // Success
                () => {
                    // Hide the dialog
                    LoadingDialog.hide();

                    // Call the update event
                    onUpdate();
                },

                // Error
                () => {
                    // Hide the dialog
                    LoadingDialog.hide();

                    // Error deploying the app
                    // TODO - Show an error
                    // Call the update event
                    onUpdate();
                }
            );
        });
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
                    webUrl: Strings.SourceUrl
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

        // Get the assessment item
        this.getAssessmentItem(item, false).then(assessment => {
            // See if an item exists
            if (assessment) {
                let completeFl = false;
                let alert: Components.IAlert = null;

                // Show the edit form
                ItemForm.edit({
                    itemId: assessment.Id,
                    webUrl: Strings.SourceUrl,
                    onSetHeader: el => {
                        // Render an alert
                        alert = Components.Alert({
                            el,
                            content: "You still need to update the item to complete the assessment."
                        });

                        // Hide it by default
                        alert.hide();
                    },
                    onSetFooter: el => {
                        // Render a completed button
                        let tooltip = Components.Tooltip({
                            el,
                            content: "Completes the review of the app.",
                            btnProps: {
                                text: "Complete Review",
                                type: Components.ButtonTypes.OutlineSuccess,
                                onClick: () => {
                                    // Set the flag
                                    completeFl = true;

                                    // Disable the button
                                    tooltip.button.disable();

                                    // Show the alert
                                    alert.show();
                                }
                            }
                        })
                    },
                    onCreateEditForm: props => {
                        // Set the rendering event
                        props.onControlRendering = (props, field) => {
                            // See if this is a test case field
                            if (field.InternalName.indexOf("TechReview") == 0 && field.FieldTypeKind == SPTypes.FieldType.Choice) {
                                // Set the validation
                                props.onValidate = (ctrl, results) => {
                                    // See if the complete flag is set
                                    if (completeFl) {
                                        // Set the flag
                                        let selectedItem: Components.IDropdownItem = results.value;
                                        results.isValid = selectedItem && selectedItem.text ? true : false;

                                        // Set the error message
                                        results.invalidMessage = "A selection is required.";
                                    }

                                    // Return the results
                                    return results;
                                }
                            }
                        }

                        // Return the properties
                        return props;
                    },
                    onSave: (props: IAssessmentItem) => {
                        // See if we are completing the assessment
                        if (completeFl) {
                            // Update the props
                            props.Completed = new Date(Date.now()) as any;
                        }

                        // Return the properties
                        return props;
                    },
                    onUpdate: () => {
                        // See if we are updating the status
                        if (completeFl) {
                            // Update the status
                            item.update({ AppStatus: "Pending Deployment" }).execute(onUpdate);
                        } else {
                            // Call the update event
                            onUpdate();
                        }
                    }
                });
            } else {
                // Show the new form
                ItemForm.create({
                    webUrl: Strings.SourceUrl,
                    onCreateEditForm: props => {
                        // Update the default title
                        props.onControlRendering = ((ctrl, field) => {
                            // See if this is the title field
                            if (field.InternalName == "Title") {
                                // Set the value
                                ctrl.value = item.Title;
                            }
                        });

                        // Return the properties
                        return props;
                    },
                    onSave: (props: IAssessmentItem) => {
                        // Set the title
                        props.Title = item.Title + " Review " + (new Date(Date.now()).toDateString());

                        // Set the item id
                        props.RelatedAppId = item.Id;

                        // Return the properties
                        return props;
                    },
                    onUpdate
                });
            }
        });
    }

    // Views the tests for the application
    editTestCases(item: IAppItem, onUpdate: () => void) {
        // Set the form properties
        ItemForm.AutoClose = false;
        ItemForm.ListName = Strings.Lists.Assessments;

        // Get the assessment item
        this.getAssessmentItem(item).then(assessment => {
            // See if an item exists
            if (assessment) {
                let completeFl = false;
                let alert: Components.IAlert = null;

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
                    onSetHeader: el => {
                        // Render an alert
                        alert = Components.Alert({
                            el,
                            content: "You still need to update the item to complete the assessment."
                        });

                        // Hide it by default
                        alert.hide();
                    },
                    onSetFooter: el => {
                        // Render a completed button
                        let tooltip = Components.Tooltip({
                            el,
                            content: "Completes the review of the app. You must update the item to complete it.",
                            btnProps: {
                                text: "Complete Review on Save",
                                type: Components.ButtonTypes.OutlineSuccess,
                                onClick: () => {
                                    // Set the flag
                                    completeFl = true;

                                    // Disable the button
                                    tooltip.button.disable();

                                    // Show the alert
                                    alert.show();
                                }
                            }
                        })
                    },
                    onCreateEditForm: props => {
                        // Set the rendering event
                        props.onControlRendering = (props, field) => {
                            // See if this is a test case field
                            if (field.InternalName.indexOf("TestCase") == 0 && field.FieldTypeKind == SPTypes.FieldType.Choice) {
                                // Set the validation
                                props.onValidate = (ctrl, results) => {
                                    // See if the complete flag is set
                                    if (completeFl) {
                                        // Set the flag
                                        let selectedItem: Components.IDropdownItem = results.value;
                                        results.isValid = selectedItem && selectedItem.text ? true : false;

                                        // Set the error message
                                        results.invalidMessage = "A selection is required.";
                                    }

                                    // Return the results
                                    return results;
                                }
                            }
                        }

                        // Return the properties
                        return props;
                    },
                    onSave: (props: IAssessmentItem) => {
                        // See if we are completing the assessment
                        if (completeFl) {
                            // Update the props
                            props.Completed = new Date(Date.now()) as any;
                        }

                        // Return the properties
                        return props;
                    },
                    onUpdate: () => {
                        // See if we are updating the status
                        if (completeFl) {
                            // Update the status
                            item.update({ AppStatus: "Pending Approval" }).execute(onUpdate);
                        } else {
                            // Call the update event
                            onUpdate();
                        }
                    }
                });
            } else {
                // Show the new form
                ItemForm.create({
                    webUrl: Strings.SourceUrl,
                    onGetListInfo: props => {
                        // Set the content type
                        props.contentType = "TestCases";

                        // Return the properties
                        return props;
                    },
                    onSave: (props: IAssessmentItem) => {
                        // Set the title
                        props.Title = item.Title + " Tests " + (new Date(Date.now()).toDateString());

                        // Set the related app
                        props.RelatedAppId = item.Id;

                        // Return the properties
                        return props;
                    },
                    onUpdate
                });
            }
        });
    }

    // Method to get the assessment item associated with the app
    private getAssessmentItem(item: IAppItem, testFl: boolean = true, lastFl: boolean = false): PromiseLike<IAssessmentItem> {
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

    // Last Assessment form
    lastAssessment(item: IAppItem) {
        // Set the list name
        ItemForm.AutoClose = false;
        ItemForm.ListName = Strings.Lists.Assessments;

        // Get the assessment item
        this.getAssessmentItem(item, false, true).then(assessment => {
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

    // Reads an app package file
    private readPackage(data): PromiseLike<IAppItem> {
        // Return a promise
        return new Promise(resolve => {
            // Unzip the package
            loadAsync(data).then(files => {
                let metadata: IAppItem = {} as any;
                let imageFound = false;

                // Parse the files
                files.forEach((path, fileInfo) => {
                    /** What are we doing here? */
                    /** Should we convert this to use the doc set dashboard solution? */

                    // See if this is an image
                    if (fileInfo.name.endsWith(".png") || fileInfo.name.endsWith(".jpg") || fileInfo.name.endsWith(".jpeg") || fileInfo.name.endsWith(".gif")) {
                        // NOTE: Reading of the file is asynchronous!
                        fileInfo.async("base64").then(function (content) {
                            //$get("appIcon").src = "data:image/png;base64," + content;
                            imageFound = true;
                        });

                        // Get image in different format for later uploading
                        fileInfo.async("arraybuffer").then(function (content) {
                            // Are we storing this somewhere?
                        });

                        // Check the next file
                        return;
                    }

                    // See if this is the app manifest
                    if (fileInfo.name != "AppManifest.xml") { return; }

                    // Read the file
                    fileInfo.async("string").then(content => {
                        var i = content.indexOf("<App");
                        var oParser = new DOMParser();
                        var oDOM = oParser.parseFromString(content.substring(i), "text/xml");

                        // Set the title
                        let elTitle = oDOM.getElementsByTagName("Title")[0];
                        if (elTitle) { metadata.Title = elTitle.textContent; }

                        // Set the version
                        let elVersion = oDOM.documentElement.attributes["Version"];
                        if (elVersion) { metadata.AppVersion = elVersion.value; }

                        // Set the permissions
                        let elPermissions = oDOM.documentElement.querySelector("WebApiPermissionRequests");
                        if (elPermissions) {
                            metadata.AppAPIPermissions = (elPermissions.innerHTML || "")
                                .replace(/&lt;/g, '<')
                                .replace(/&gt;/g, '>')
                                .replace(/&amp;/g, '&')
                                .replace(/&quot;/g, '"');
                        }

                        // Set the product id
                        let elProductId = oDOM.documentElement.attributes["ProductID"];
                        if (elProductId) { metadata.AppProductID = elProductId.value; }

                        // Set the client side solution flag
                        let elIsClientSideSolution = oDOM.documentElement.attributes["IsClientSideSolution"];
                        metadata.AppIsClientSideSolution = elIsClientSideSolution ? elIsClientSideSolution.value == "true" : false;

                        // Set the domain isolation flag
                        let elIsDomainIsolated = oDOM.documentElement.attributes["IsDomainIsolated"];
                        metadata.AppIsDomainIsolated = elIsDomainIsolated ? elIsDomainIsolated.value == "true" : false;

                        // Set the sp min version flag
                        let elSPMinVersion = oDOM.documentElement.attributes["SharePointMinVersion"];
                        if (elSPMinVersion) { metadata.AppSharePointMinVersion = elSPMinVersion.value; }

                        // Set the skip feature deployment flag
                        let elSkipFeatureDeployment = oDOM.documentElement.attributes["SkipFeatureDeployment"];
                        metadata.AppSkipFeatureDeployment = elSkipFeatureDeployment ? elSkipFeatureDeployment.value == "true" : false;

                        //var spfxElem = oDOM.documentElement.attributes["IsClientSideSolution"]; //Not for add-ins
                        //var isSPFx = (spfxElem ? spfxElem.value : "false");

                        // Set the status
                        metadata.AppStatus = "New";

                        // Resolve the request
                        resolve(metadata);
                    });
                });
            });
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
                    item.update({
                        AppComments: comments,
                        AppIsRejected: true
                    }).execute(() => {
                        // Parse the developers
                        let cc = [];
                        for (let i = 0; i < DataSource.DevGroup.Users.results.length; i++) {
                            // Append the email
                            cc.push(DataSource.DevGroup.Users.results[i].Email);
                        }

                        // Get the app developers
                        let to = [];
                        let owners = item.AppDevelopers && item.AppDevelopers.results ? item.AppDevelopers.results : [];
                        for (let i = 0; i < owners.length; i++) {
                            // Append the email
                            to.push(owners[i].EMail);
                        }

                        // Ensure owners exist
                        if (to.length > 0) {
                            // Send an email
                            Utility().sendEmail({
                                To: to,
                                CC: cc,
                                Subject: "App '" + item.Title + "' Sent Back",
                                Body: "App Developers,<br /><br />The '" + item.Title + "' app has been sent back based on the comments below." +
                                    ContextInfo.userDisplayName +
                                    ".<br /><br />" + comments
                            }).execute(() => {
                                // Call the update event
                                onUpdate();

                                // Close the loading dialog
                                LoadingDialog.hide();
                            });
                        } else {
                            // Call the update event
                            onUpdate();

                            // Close the loading dialog
                            LoadingDialog.hide();
                        }
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
        let catalogUrl = tenantFl ? DataSource.Configuration.tenantAppCatalogUrl : DataSource.Configuration.appCatalogUrl;
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

        // See if a checklist exists
        let checklist = (DataSource.Configuration.checklists ? DataSource.Configuration.checklists[item.AppStatus] : null) || [];
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

                // Show a loading dialog
                LoadingDialog.setHeader("Updating App Submission");
                LoadingDialog.setBody("This dialog will close after the app submission is completed.");
                LoadingDialog.show();

                // Update the status
                let status = DataSource.Status[item.AppStatus];
                item.update({
                    AppIsRejected: false,
                    AppStatus: status ? status.nextStep : DataSource.Status[0]
                }).execute(() => {
                    // Parse the developers
                    let to = DataSource.Configuration.appCatalogAdminEmailGroup ? [DataSource.Configuration.appCatalogAdminEmailGroup] : [];
                    for (let i = 0; i < DataSource.DevGroup.Users.results.length; i++) {
                        // Append the email
                        to.push(DataSource.DevGroup.Users.results[i].Email);
                    }

                    // Get the app developers
                    let cc = [];
                    let owners = item.AppDevelopers.results || [];
                    for (let i = 0; i < owners.length; i++) {
                        // Append the email
                        cc.push(owners[i].EMail);
                    }

                    // Ensure owners exist
                    if (to.length > 0) {
                        // Send an email
                        Utility().sendEmail({
                            To: to,
                            CC: cc,
                            Subject: "App '" + item.Title + "' submitted for approval",
                            Body: "App Developers,<br /><br />The '" + item.Title + "' app has been submitted for approval by " + ContextInfo.userDisplayName + ". Please take some time to test this app and submit an assessment/review using the App Dashboard."
                        }).execute(() => {
                            // Call the update event
                            onUpdate();

                            // Close the loading dialog
                            LoadingDialog.hide();
                        });
                    } else {
                        // Call the update event
                        onUpdate();

                        // Close the loading dialog
                        LoadingDialog.hide();
                    }
                });
            }
        }).el);

        // Show the modal
        Modal.show();
    }

    // Updates the app metadata
    updateApp(item: IAppItem, appItemId: number, tenantFl: boolean, catalogUrl: string, requestDigest: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Get the app catalog
            Web(catalogUrl, { requestDigest }).Lists().query({
                Filter: "BaseTemplate eq " + (tenantFl ? SPTypes.ListTemplateType.TenantAppCatalog : SPTypes.ListTemplateType.AppCatalog)
            }).execute(lists => {
                // Ensure the app catalog was found
                let list: Types.SP.List = lists.results[0] as any;
                if (list) {
                    // Update the metadata
                    list.Items(appItemId).update({
                        AppDescription: item.AppDescription,
                        AppImageURL1: item.AppImageURL1,
                        AppImageURL2: item.AppImageURL2,
                        AppImageURL3: item.AppImageURL3,
                        AppImageURL4: item.AppImageURL4,
                        AppImageURL5: item.AppImageURL5,
                        AppShortDescription: item.AppShortDescription,
                        AppSupportURL: item.AppSupportURL,
                        AppVideoURL: item.AppVideoURL
                    }).execute(() => {
                        // Resolve the request
                        resolve();
                    });
                } else {
                    // Resolve the request
                    resolve();
                }
            });
        });
    }

    // Upload form
    upload(onUpdate: () => void) {
        // Show the upload file dialog
        Helper.ListForm.showFileDialog().then(file => {
            // Ensure this is an spfx package
            if (file.name.toLowerCase().endsWith(".sppkg")) {
                // Display a loading dialog
                LoadingDialog.setHeader("Reading Package");
                LoadingDialog.setBody("Validating the package...");
                LoadingDialog.show();

                // Extract the metadata from the package
                this.readPackage(file.data).then(data => {
                    // Validate the data
                    if (data.AppProductID && data.AppVersion && data.Title) {
                        // Update the loading dialog
                        LoadingDialog.setHeader("Creating App Folder");
                        LoadingDialog.setBody("Creating the app folder...");

                        // Create the document set folder
                        Helper.createDocSet(data.Title, Strings.Lists.Apps).then(
                            // Success
                            item => {
                                // Update the loading dialog
                                LoadingDialog.setHeader("Updating Metadata");
                                LoadingDialog.setBody("Saving the package information...");

                                // Default the owner to the current user
                                data["AppDevelopersId"] = { results: [ContextInfo.userId] } as any;

                                // Update the metadata
                                item.update(data).execute(() => {
                                    // Execute the update event
                                    onUpdate();

                                    // Update the loading dialog
                                    LoadingDialog.setHeader("Uploading the Package");
                                    LoadingDialog.setBody("Uploading the app package...");

                                    // Upload the file
                                    item.Folder().Files().add(file.name, true, file.data).execute(
                                        // Success
                                        file => {
                                            // Close the loading dialog
                                            LoadingDialog.hide();

                                            // Display the edit form
                                            this.edit(item.Id, onUpdate);
                                        },

                                        // Error
                                        () => {
                                            // TODO
                                        }
                                    );
                                });
                            },

                            // Error
                            () => {
                                // TODO
                            }
                        );
                    } else {
                        // Close the loading dialog
                        LoadingDialog.hide();

                        // Display a modal
                        Modal.clear();
                        Modal.setHeader("Package Validation Error");
                        Modal.setBody("The spfx package is invalid.");
                        Modal.show();
                    }
                });
            } else {
                // Display a modal
                Modal.clear();
                Modal.setHeader("Error Adding Package");
                Modal.setBody("The file must be a valid SPFx app package file.");
                Modal.show();
            }
        });
    }
}