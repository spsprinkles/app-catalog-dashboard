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

    // Approve form
    approve(item: IAppItem, onUpdate: () => void) {
        // Set the header
        Modal.setHeader("Approve App/Solution Package");

        // Set the body
        Modal.setBody("Are you sure you want to approve this app package?");

        // Render the footer
        Modal.setFooter(Components.Tooltip({
            content: "This will approve the app and make it available for deployement.",
            btnProps: {
                text: "Approve",
                type: Components.ButtonTypes.OutlineSuccess,
                onClick: () => {
                    // Close the modal
                    Modal.hide();

                    // Show a loading dialog
                    LoadingDialog.setHeader("Approving the App");
                    LoadingDialog.setBody("This dialog will close after the app is approved.");
                    LoadingDialog.show();

                    // Update the item
                    item.update({
                        DevAppStatus: "Approved"
                    }).execute(() => {
                        // Close the dialog
                        LoadingDialog.hide();

                        // Execute the update event
                        onUpdate();
                    });
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

        // Determine if we can delete this package
        let deleteFl = item.HasUniqueRoleAssignments != true;

        // Set the body based on the flag
        if (deleteFl) {
            // Set the body
            Modal.setBody("Are you sure you want to delete this app package? This will remove it from the tenant app catalog.");
        } else {
            // Set the body
            Modal.setBody("This package is already published and cannot be deleted. Would you like to disable it?");
        }

        // Render the footer
        Modal.setFooter(Components.Button({
            text: deleteFl ? "Delete" : "Disable",
            type: Components.ButtonTypes.OutlineDanger,
            onClick: () => {
                // Close the modal
                Modal.hide();

                // Show a loading dialog
                LoadingDialog.setHeader(deleteFl ? "Deleting the App" : "Disabling the App");
                LoadingDialog.setBody("This dialog will close after the app is updated.");
                LoadingDialog.show();

                // See if we are deleting the package
                if (deleteFl) {
                    // Delete the assessments w/ this app
                    this.deleteAssessments(item).then(() => {
                        // Delete this folder
                        item.delete().execute(() => {
                            // Close the dialog
                            LoadingDialog.hide();

                            // Call the update event
                            onUpdate();
                        });
                    });
                }
                // Else, we are disabling the app
                else {
                    // Retract the solution from the site catalog
                    this.retract(item, false, () => {
                        // Retract the solution from the tenant
                        this.retract(item, true, () => {
                            // Update the item
                            item.update({
                                IsAppPackageEnabled: false
                            }).execute(() => {
                                // Close the dialog
                                LoadingDialog.hide();

                                // Execute the update event
                                onUpdate();
                            });
                        });
                    });
                }
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
                                    // Call the update event
                                    onUpdate();

                                    // App deployed
                                    LoadingDialog.hide();
                                }, () => {
                                    // Error deploying the app
                                    // TODO - Show an error
                                    // Call the update event
                                    onUpdate();

                                    // App deployed
                                    LoadingDialog.hide();
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
                    // Call the update event
                    onUpdate();

                    // App deployed
                    LoadingDialog.hide();
                },

                // Error
                () => {
                    // Error deploying the app
                    // TODO - Show an error
                    // Call the update event
                    onUpdate();

                    // App deployed
                    LoadingDialog.hide();
                }
            );
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
                props.excludeFields = ["AppVideoURL", "DevAppStatus", "IsAppPackageEnabled", "IsDefaultAppMetadataLocale"];

                // Update the field
                props.onControlRendering = (ctrl, field) => {
                    // See if this is a read-only field
                    if (["AppProductID", "AppVersion", "FileLeafRef", "Title"].indexOf(field.InternalName) >= 0) {
                        // Make it read-only
                        ctrl.isReadonly = true;
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

    // Method to get the assessment item associated with the app
    private getAssessmentItem(item: IAppItem, lastFl: boolean = false): PromiseLike<IAssessmentItem> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the assoicated item
            List(Strings.Lists.Assessments).Items().query({
                Filter: "RelatedAppId eq " + item.Id,
                OrderBy: ["Completed desc"]
            }).execute(items => {
                // Parse the results
                for (let i = 0; i < items.results.length; i++) {
                    let item: IAssessmentItem = items.results[i] as any;

                    // See if this is not completed
                    if (item.Completed == null) {
                        // Resolve the request
                        resolve(item);
                        return;
                    }
                    // Else, see if we are viewing the last assessment
                    else if (lastFl) {
                        resolve(item);
                        return;
                    }
                }

                // Not found
                resolve(null);
            }, reject);
        });
    }

    // Last Assessment form
    lastAssessment(item: IAppItem) {
        // Set the list name
        ItemForm.AutoClose = false;
        ItemForm.ListName = Strings.Lists.Assessments;

        // Get the assessment item
        this.getAssessmentItem(item, true).then(assessment => {
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

                        // Set the product id
                        let elProductId = oDOM.documentElement.attributes["ProductID"];
                        if (elProductId) { metadata.AppProductID = elProductId.value; }

                        // Are we storing this?
                        var spfxElem = oDOM.documentElement.attributes["IsClientSideSolution"]; //Not for add-ins
                        var isSPFx = (spfxElem ? spfxElem.value : "false");

                        //Attribute: SharePointMinVersion="16.0.0.0" // 15.0.0.0
                        //Attribute: SkipFeatureDeployment="true" (not for add-ins)

                        // Update the status
                        metadata.DevAppStatus = "Draft";

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
                        DevAppStatus: "Requires Attention"
                    }).execute(() => {
                        // Parse the developers
                        let cc = [];
                        for (let i = 0; i < DataSource.DevGroup.Users.results.length; i++) {
                            // Append the email
                            cc.push(DataSource.DevGroup.Users.results[i].Email);
                        }

                        // Get the app owners
                        let to = [];
                        let owners = item.Owners && item.Owners.results ? item.Owners.results : [];
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
    retract(item: IAppItem, tenantFl: boolean, onUpdate: () => void) {
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
                // Call the update event
                onUpdate();

                // Close the dialog
                LoadingDialog.hide();
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

    // Review form
    review(item: IAppItem, onUpdate: () => void) {
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
                            item.update({ DevAppStatus: "Requesting Approval" }).execute(onUpdate);
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
                        // Set the item id
                        props.RelatedAppId = item.Id;

                        // Update the status of the current item
                        if (item.DevAppStatus == "Submitted for Review") {
                            // Update the status
                            item.update({
                                DevAppStatus: "In Review"
                            }).execute(onUpdate);
                        }

                        // Return the properties
                        return props;
                    },
                    onUpdate
                });
            }
        });
    }

    // Submit form
    submit(item: IAppItem, onUpdate: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Submit App for Review");

        // Set the body
        Modal.setBody("Are you sure you want to submit this app for review?");

        // Set the footer
        Modal.setFooter(Components.Button({
            text: "Submit",
            type: Components.ButtonTypes.OutlinePrimary,
            onClick: () => {
                // Close the modal
                Modal.hide();

                // Show a loading dialog
                LoadingDialog.setHeader("Updating App Submission");
                LoadingDialog.setBody("This dialog will close after the app submission is updated.");
                LoadingDialog.show();

                // Update the status
                item.update({
                    DevAppStatus: "Submitted for Review"
                }).execute(() => {
                    // Parse the developers
                    let to = DataSource.Configuration.appCatalogAdminEmailGroup ? [DataSource.Configuration.appCatalogAdminEmailGroup] : [];
                    for (let i = 0; i < DataSource.DevGroup.Users.results.length; i++) {
                        // Append the email
                        to.push(DataSource.DevGroup.Users.results[i].Email);
                    }

                    // Get the app owners
                    let cc = [];
                    let owners = item.Owners.results || [];
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
                            Subject: "App '" + item.Title + "' submitted for review",
                            Body: "App Developers,<br /><br />The '" + item.Title + "' app has been submitted for review by " + ContextInfo.userDisplayName + ". Please take some time to test this app and submit an assessment/review using the App Dashboard."
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
                                data["OwnersId"] = { results: [ContextInfo.userId] } as any;

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