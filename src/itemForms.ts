import { Components } from ".pnpm/gd-bs@5.2.0/node_modules/gd-bs";
import { ItemForm, LoadingDialog, Modal } from "dattatable";
import { ContextInfo, Helper, List, Utility } from "gd-sprest-bs";
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
            onCreateEditForm: props => {
                // Exclude fields
                props.excludeFields = ["DevAppStatus"];

                // Update the field
                props.onControlRendering = (ctrl, field) => {
                    // See if this is a read-only field
                    if (field.InternalName == "FileLeafRef") {
                        // Make it read-only
                        ctrl.isReadonly = true;
                    }
                }

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
                // Upload the file
                List(Strings.Lists.Apps).RootFolder().Files().add(file.name, true, file.data).execute(
                    // Success
                    file => {
                        // Get the item id
                        file.ListItemAllFields().execute(item => {
                            // Display the edit form
                            this.edit(item.Id, onUpdate);
                        });
                    },

                    // Error
                    () => {
                        // TODO
                    }
                )
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