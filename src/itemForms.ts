import { ItemForm, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, List, Utility, Web } from "gd-sprest-bs";
import { loadAsync } from "jszip";
import { DataSource, IAppItem, IAssessmentItem } from "./ds";
import Strings from "./strings";

/**
 * App Item Forms
 */
export class AppForms {
    // Constructor
    constructor() { }

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
                    // Get the associated item
                    this.getAssessmentItem(item).then(
                        assessmentItem => {
                            // Delete it
                            assessmentItem.delete().execute(() => {
                                // Delete this item
                                item.delete().execute(() => {
                                    // Close the dialog
                                    LoadingDialog.hide();

                                    // Call the update event
                                    onUpdate();
                                });
                            });
                        },

                        () => {
                            // Delete this item
                            item.delete().execute(() => {
                                // Close the dialog
                                LoadingDialog.hide();

                                // Call the update event
                                onUpdate();
                            });
                        }
                    );
                }
                // Else, we are disabling the app
                else {
                    // Update the item
                    item.update({
                        IsAppPackageEnabled: false
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

    // Deploys the solution to the app catalog
    deploy(item: IAppItem, onUpdate: () => void) {
        // Show a loading dialog
        LoadingDialog.setHeader("Uploading Package");
        LoadingDialog.setBody("Uploading the spfx package to the app tenant catalog.");
        LoadingDialog.show();

        // Upload the package file
        item.File().content().execute(content => {
            // Load the context of the app catalog
            ContextInfo.getWeb(DataSource.Configuration.tenantAppCatalogUrl).execute(context => {
                // Upload the file to the app catalog
                Web(DataSource.Configuration.tenantAppCatalogUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue })
                    .TenantAppCatalog().add(item.FileLeafRef, true, content).execute(file => {
                        // Update the dialog
                        LoadingDialog.setHeader("Deploying the Package");
                        LoadingDialog.setBody("This will close after the app is deployed.");

                        // Check-in the file
                        file.checkIn().execute(() => {
                            // Do we need to do this by default?
                        });

                        // Get the apps
                        Web(DataSource.Configuration.tenantAppCatalogUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue })
                            .TenantAppCatalog().AvailableApps().execute(apps => {
                                for (let i = 0; i < apps.results.length; i++) {
                                    let app = apps.results[i];

                                    // See if this is the target app
                                    if (app.ProductId != item.AppProductID) { continue; }

                                    // See if it's not deployed
                                    if (app.Deployed != true) {
                                        // Deploy the app
                                        app.deploy().execute(() => {
                                            // Call the update event
                                            onUpdate();

                                            // App deployed
                                            LoadingDialog.hide();
                                        });
                                    } else {
                                        // Call the update event
                                        onUpdate();

                                        // Close the dialog
                                        LoadingDialog.hide();
                                    }
                                }
                            });
                    });
            });
        });
    }

    // Edit form
    edit(itemId: number, onUpdate: () => void) {
        // Set the list name
        ItemForm.ListName = Strings.Lists.Apps;

        // Show the item form
        ItemForm.edit({
            itemId,
            onSetHeader: el => {
                // Update the header
                el.innerHTML = "New App"
            },
            onGetListInfo: props => {
                // Set the content type
                props.contentType = "App";

                // Return the properties
                return props;
            },
            onCreateEditForm: props => {
                // Exclude fields
                props.excludeFields = ["DevAppStatus"];

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
    private getAssessmentItem(item: IAppItem): PromiseLike<IAssessmentItem> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the assoicated item
            List(Strings.Lists.Assessments).Items().query({
                Filter: "RelatedAppId eq " + item.Id
            }).execute(items => {
                let item = items.results[0] as any;

                // Resolve/Reject the promise
                item ? resolve(item) : reject();
            }, reject);
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

                        // See if the metadata is valid
                        if (metadata.Title && metadata.AppVersion && metadata.AppProductID) {
                            // Update the status
                            metadata.DevAppStatus = "In Review";
                        }

                        // Resolve the request
                        resolve(metadata);
                    });
                });
            });
        });
    }

    // Retracts the solution to the app catalog
    retract(item: IAppItem, onUpdate: () => void) {
        // Show a loading dialog
        LoadingDialog.setHeader("Retracting the Package");
        LoadingDialog.setBody("Retracting the spfx package to the app tenant catalog.");
        LoadingDialog.show();

        // Load the context of the app catalog
        ContextInfo.getWeb(DataSource.Configuration.tenantAppCatalogUrl).execute(context => {
            // Get the apps
            Web(DataSource.Configuration.tenantAppCatalogUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue })
                .TenantAppCatalog().AvailableApps().execute(apps => {
                    for (let i = 0; i < apps.results.length; i++) {
                        let app = apps.results[i];

                        // See if this is the target app
                        if (app.ProductId != item.AppProductID) { continue; }

                        // See if it's deployed
                        if (app.Deployed) {
                            // Retract the app
                            app.retract().execute(() => {
                                // Call the update event
                                onUpdate();

                                // App deployed
                                LoadingDialog.hide();
                            });
                        } else {
                            // Call the update event
                            onUpdate();

                            // Close the dialog
                            LoadingDialog.hide();
                        }
                    }
                });
        });
    }

    // Review form
    review(item: IAppItem, onUpdate: () => void) {
        // Set the list name
        ItemForm.ListName = Strings.Lists.Assessments;

        // Get the assessment item
        this.getAssessmentItem(item).then(
            // Exsists
            assessmentItem => {
                // Show the edit form
                ItemForm.edit({
                    itemId: assessmentItem.Id,
                    onCreateEditForm: props => {
                        // Update the field controls
                        props.onControlRendering = (ctrl, field) => {
                            // See if this is the app field
                            if (field.InternalName == "RelatedApp") {
                                // Make the field readonly
                                ctrl.isReadonly = true;
                            }
                        }

                        // Return the properties
                        return props;
                    },
                    onUpdate: () => {
                        if (item.DevAppStatus != "In Review") {
                            // Update the status of the item
                            item.update({
                                DevAppStatus: "In Review"
                            }).execute(() => {
                                // Call the update event
                                onUpdate();
                            });
                        } else {
                            // Call the update event
                            onUpdate();
                        }
                    }
                });
            },

            // Doesn't exist
            () => {
                // Show the new form
                ItemForm.create({
                    onCreateEditForm: props => {
                        // Exclude the related app id
                        props.excludeFields = ["RelatedApp"];

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
                    onSave: props => {
                        // Set the associated app id
                        props["RelatedAppId"] = item.Id;

                        // Return the properties
                        return props;
                    },
                    onUpdate: () => {
                        // Update the status of the item
                        item.update({
                            DevAppStatus: "In Review"
                        }).execute(() => {
                            // Call the update event
                            onUpdate();
                        });
                    }
                });
            });
    }

    // Submit form
    submit(item: IAppItem, onUpdate: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the header
        Modal.setHeader("Sumit App for Review");

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
                    DevAppStatus: "In Review"
                }).execute(() => {
                    // Get the app owners
                    let cc = [];
                    let owners = item.Owners.results || [];
                    for (let i = 0; i < owners.length; i++) {
                        // Append the email
                        cc.push(owners[i].EMail);
                    }

                    // Send an email
                    Utility().sendEmail({
                        To: [DataSource.Configuration.appCatalogAdminEmailGroup],
                        CC: cc,
                        Subject: "App '" + item.Title + "' submitted for review",
                        Body: "App Developers,<br /><br />The '" + item.Title + "' app has been submitted for review by " + ContextInfo.userDisplayName + ". Please take some time to test this app and submit an assessment/review using the App Dashboard."
                    }).execute(() => {
                        // Call the update event
                        onUpdate();

                        // Close the loading dialog
                        LoadingDialog.hide();
                    });
                });
            }
        }).el);

        // Show the modal
        Modal.show();
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