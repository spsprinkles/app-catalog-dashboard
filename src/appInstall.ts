import { InstallationRequired, LoadingDialog } from "dattatable";
import { Components, Site } from "gd-sprest-bs";
import { AppConfig } from "./appCfg";
import { AppSecurity } from "./appSecurity";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import { ErrorDialog } from "./errorDialog";
import Strings from "./strings";

/**
 * App Install
 */
export class AppInstall {
    // Determines if the document set feature is enabled on this site
    private static docSetEnabled(): PromiseLike<boolean> {
        // Return a promise
        return new Promise((resolve) => {
            // See if the site collection has the feature enabled
            Site(Strings.SourceUrl).Features("3bae86a2-776d-499d-9db8-fa4cdc7884f8").execute(
                // Request was successful
                feature => {
                    // Ensure the feature exists and resolve the request
                    resolve(feature.DefinitionId ? true : false);
                },

                // Not enabled
                () => {
                    // Resolve the request
                    resolve(false);
                }
            )
        });
    }

    // Sees if an install is required and displays a dialog
    static InstallRequired(el: HTMLElement, showFl: boolean = false) {
        let errors: Components.IListGroupItem[] = [];
        let auditLogErrorFl = true;

        // See if an install is required
        InstallationRequired.requiresInstall({
            cfg: [DataSource.AuditLog.Configuration, Configuration],
            onError: (cfg) => {
                // See if this is the audit list
                if (DataSource.AuditLog.Configuration._configuration.ListCfg[0].ListInformation.Title == cfg._configuration.ListCfg[0].ListInformation.Title) {
                    // Set the flag
                    auditLogErrorFl = false;
                }
            }
        }).then(errorFl => {
            // Ensure the document set feature is enabled
            this.docSetEnabled().then(featureEnabledFl => {
                // See if the feature is enabled
                if (!featureEnabledFl) {
                    // Add an error
                    errors.push({
                        content: "Document Set site feature is not enabled.",
                        type: Components.ListGroupItemTypes.Danger
                    });
                }

                // See if the configuration is correct
                let cfgIsValid = true;
                if (AppConfig.Configuration == null) {
                    // Update the flag
                    cfgIsValid = false;

                    // Add an error
                    errors.push({
                        content: "App configuration doesn't exist or is not in the correct JSON format. Please edit the webpart and set the configuration property.",
                        type: featureEnabledFl ? null : Components.ListGroupItemTypes.Danger
                    });
                }
                // Else, ensure it's valid
                else if (!AppConfig.IsValid) {
                    // Update the flag
                    cfgIsValid = false;

                    // Add an error
                    errors.push({
                        content: "App configuration exists, but is invalid. Please contact your administrator.",
                        type: featureEnabledFl ? null : Components.ListGroupItemTypes.Danger
                    });
                }

                // See if the security groups are configured
                let securityGroupsExist = true;
                if (AppSecurity.hasErrors()) {
                    // Set the flag
                    securityGroupsExist = false;

                    // Add an error
                    errors.push({
                        content: "Security groups are not configured for the app.",
                        type: featureEnabledFl && cfgIsValid ? null : Components.ListGroupItemTypes.Danger
                    });
                }

                // See if an installation is required
                if ((errorFl || errors.length > 0) || showFl) {
                    // Show the installation dialog
                    InstallationRequired.showDialog({
                        errors,
                        onCompleted: (() => {
                            // Configure the web
                            AppSecurity.configureWeb().then(
                                // Success
                                () => {
                                    // Refresh the page
                                    window.location.reload();
                                },
                                // Error
                                ex => {
                                    // Show the error
                                    ErrorDialog.show("Error Configuring Web", "There was an error configuring the app web.", ex);
                                }
                            );
                        }),
                        onFooterRendered: el => {
                            // See if the configuration isn't defined
                            if (!cfgIsValid || !securityGroupsExist) {
                                // Disable the install button
                                (el.firstChild as HTMLButtonElement).disabled = true;
                            }

                            // See if the audit log doesn't exist
                            if (auditLogErrorFl) {
                                // Add the custom install button
                                Components.Tooltip({
                                    el,
                                    content: "Installs the audit log list.",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    btnProps: {
                                        text: "Install Audit Log",
                                        onClick: () => {
                                            // Show a loading dialog
                                            LoadingDialog.setHeader("Creating List");
                                            LoadingDialog.setBody("This dialog will close after list is created...");
                                            LoadingDialog.show();

                                            // Create the list
                                            DataSource.AuditLog.install().then(
                                                // Success
                                                () => {
                                                    // Close the dialog
                                                    LoadingDialog.hide();

                                                    // Refresh the page
                                                    window.location.reload();
                                                },
                                                // Error
                                                ex => {
                                                    // Show the error
                                                    ErrorDialog.show("Error Creating List", "There was an error creating the audit log list.", ex);
                                                }
                                            );
                                        }
                                    }
                                });
                            }

                            // See if the feature isn't enabled
                            if (!featureEnabledFl) {
                                // Add the custom install button
                                Components.Tooltip({
                                    el,
                                    content: "Enables the document set site collection feature.",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    btnProps: {
                                        text: "Enable Feature",
                                        onClick: () => {
                                            // Show a loading dialog
                                            LoadingDialog.setHeader("Enable Feature");
                                            LoadingDialog.setBody("Enabling the document set feature. This dialog will close after this requests completes.");
                                            LoadingDialog.show();

                                            // Enable the feature
                                            Site(Strings.SourceUrl).Features().add("3bae86a2-776d-499d-9db8-fa4cdc7884f8").execute(
                                                // Enabled
                                                () => {
                                                    // Close the dialog
                                                    LoadingDialog.hide();

                                                    // Refresh the page
                                                    window.location.reload();
                                                },
                                                ex => {
                                                    // Show the error
                                                    ErrorDialog.show("Enable Feature", "There was an error enabling the document set site collection feature.", ex);
                                                }
                                            );
                                        }
                                    }
                                });
                            }

                            // See if the security group doesn't exist
                            if (!securityGroupsExist || showFl) {
                                // Add the custom install button
                                Components.Tooltip({
                                    el,
                                    content: "Creates the security groups.",
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    btnProps: {
                                        text: "Security",
                                        isDisabled: !featureEnabledFl,
                                        onClick: () => {
                                            // Show a loading dialog
                                            LoadingDialog.setHeader("Security Groups");
                                            LoadingDialog.setBody("Creating the security groups. This dialog will close after it completes.");
                                            LoadingDialog.show();

                                            // Create the security groups
                                            AppSecurity.createAppSecurityGroups().then(() => {
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

                    // Clear the element
                    while (el.firstChild) { el.removeChild(el.firstChild); }

                    // Render the errors
                    Components.ListGroup({
                        el,
                        items: errors
                    });

                    // Show a button to display the modal
                    Components.Tooltip({
                        el,
                        content: "Displays the installation status modal.",
                        btnProps: {
                            text: "Status",
                            onClick: () => {
                                // Dispaly this modal
                                this.InstallRequired(el, true);
                            }
                        }
                    });
                } else {
                    // Log
                    console.error("[" + Strings.ProjectName + "] Error initializing the solution.");
                }
            });
        });
    }
}