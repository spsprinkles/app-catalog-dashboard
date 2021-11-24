import { InstallationRequired, LoadingDialog } from "dattatable";
import { App } from "./app";
import { Configuration, createSecurityGroups } from "./cfg";
import { AppDashboard } from "./dashboard";
import { DataSource } from "./ds";
import Strings, { setContext } from "./strings";
import { UserAgreement } from "./userAgreement";

// Styling
import "./styles.scss";
import { Components } from "gd-sprest-bs";

// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    render: (el: HTMLElement, context?) => {
        // See if the page context exists
        if (context) {
            // Set the context
            setContext(context);
        }

        // Hide the first column of the webpart zones
        let wpZone: HTMLElement = document.querySelector("#DeltaPlaceHolderMain table > tbody > tr > td");
        wpZone ? wpZone.style.width = "0%" : null;

        // Make the second column of the webpart zones full width
        wpZone = document.querySelector("#DeltaPlaceHolderMain table > tbody > tr > td:last-child");
        wpZone ? wpZone.style.width = "100%" : null;

        // Initialize the solution
        DataSource.init().then(
            // Success
            () => {
                // Ensure the user is an approver or developer
                if (!DataSource.IsApprover && !DataSource.IsDeveloper) {
                    // Show the user agreement
                    new UserAgreement();
                }
                // Else, see if this is a document set
                else if (DataSource.IsDocSet) {
                    // Create the application
                    new App(el);
                } else {
                    // Render the dashboard
                    new AppDashboard(el);
                }
            },
            // Error
            () => {
                // See if an install is required
                InstallationRequired.requiresInstall(Configuration).then(installFl => {
                    let errors: Components.IListGroupItem[] = [];

                    // See if the security groups exist
                    if (DataSource.ApproverGroup == null || DataSource.DevGroup == null) {
                        // Add an error
                        errors.push({
                            content: "Security groups are not installed."
                        });
                    }

                    // See if an installation is required
                    if (installFl || errors.length > 0) {
                        // Show the installation dialog
                        InstallationRequired.showDialog({
                            errors,
                            onFooterRendered: el => {
                                // See if a custom error exists
                                if (errors.length > 0) {
                                    // Add the custom install button
                                    Components.Tooltip({
                                        el,
                                        content: "Creates the security groups.",
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        btnProps: {
                                            text: "Security",
                                            onClick: () => {
                                                // Show a loading dialog
                                                LoadingDialog.setHeader("Security Groups");
                                                LoadingDialog.setBody("Creating the security groups. This dialog will close after it completes.");
                                                LoadingDialog.show();

                                                // Create the security groups
                                                createSecurityGroups().then(() => {
                                                    // Close the dialog
                                                    LoadingDialog.hide();

                                                    // Refresh the page
                                                    window.location.reload();
                                                });
                                            }
                                        }
                                    });
                                }
                            }
                        });
                    } else {
                        // Log
                        console.error("[" + Strings.ProjectName + "] Error initializing the solution.");
                    }
                });
            }
        );
    }
};

// Update the DOM
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Remove the extra border spacing on the webpart
    let contentBox = document.querySelector("#contentBox table.ms-core-tableNoSpace");
    contentBox ? contentBox.classList.remove("ms-webpartPage-root") : null;
    
    // Render the application
    GlobalVariable.render(elApp);
}