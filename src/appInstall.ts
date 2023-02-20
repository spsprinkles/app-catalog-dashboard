import { InstallationRequired, LoadingDialog } from "dattatable";
import { Components, Site } from "gd-sprest-bs";
import { AppConfig } from "./appCfg";
import { AppSecurity, createSecurityGroups } from "./appSecurity";
import { Configuration } from "./cfg";
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
        // See if an install is required
        InstallationRequired.requiresInstall(Configuration).then(installFl => {
            let errors: Components.IListGroupItem[] = [];

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

                // See if the security groups exist
                let securityGroupsExist = true;
                if (AppSecurity.ApproverGroup == null || AppSecurity.DevGroup == null || AppSecurity.FinalApproverGroup == null || AppSecurity.SponsorGroup == null) {
                    // Set the flag
                    securityGroupsExist = false;

                    // Add an error
                    errors.push({
                        content: "Security groups are not installed.",
                        type: featureEnabledFl && cfgIsValid ? null : Components.ListGroupItemTypes.Danger
                    });
                }

                // See if an installation is required
                if ((installFl || errors.length > 0) || showFl) {
                    // Show the installation dialog
                    InstallationRequired.showDialog({
                        errors,
                        onFooterRendered: el => {
                            // See if the configuration isn't defined
                            if (!cfgIsValid) {
                                // Disable the install button
                                (el.firstChild as HTMLButtonElement).disabled = true;
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
                                        isDisabled: !featureEnabledFl || !cfgIsValid || !InstallationRequired.ListsExist,
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