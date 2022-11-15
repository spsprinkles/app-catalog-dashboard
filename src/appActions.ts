import { LoadingDialog, Modal } from "dattatable";
import { ContextInfo, Helper, SPTypes, Types, Utility, Web } from "gd-sprest-bs";
import * as JSZip from "jszip";
import { AppConfig } from "./appCfg";
import { AppNotifications } from "./appNotifications";
import { DataSource, IAppItem } from "./ds";
import { ErrorDialog } from "./errorDialog";
import Strings from "./strings";


/**
 * Application Actions
 * The actions taken for an app.
 */
export class AppActions {
    // Archives the current package file
    static archivePackage(item: IAppItem, onComplete: () => void) {
        let appFile: Types.SP.File = null;

        // Show a loading dialog
        LoadingDialog.setHeader("Archiving the Package");
        LoadingDialog.setBody("The current app package is being archived.");
        LoadingDialog.show();

        // Log
        ErrorDialog.logInfo("Archiving the package....");

        // Find the package file
        for (let i = 0; i < DataSource.DocSetFolder.Files.results.length; i++) {
            let file = DataSource.DocSetFolder.Files.results[i];

            // See if this is the package
            if (file.Name.toLowerCase().endsWith(".sppkg")) {
                // Set the file
                appFile = file;
                break;
            }
        }

        // Ensure a file exists
        if (appFile == null) {
            // Log
            ErrorDialog.logInfo("The app file was not found...");

            // This shouldn't happen
            LoadingDialog.hide();
            return;
        }

        // Create the archive folder
        this.createArchiveFolder(DataSource.DocSetFolder).then(archiveFolder => {
            // Log
            ErrorDialog.logInfo(`Getting the SPFx package's file content from ${appFile.ServerRelativeUrl}...`);

            // Get the package file contents
            Web(Strings.SourceUrl).getFileByServerRelativeUrl(appFile.ServerRelativeUrl).content().execute(content => {
                // Update the dialog
                LoadingDialog.setBody("Copying the file package.")

                // Get the name of the file
                let fileName = appFile.Name.toLowerCase().split(".sppkg")[0] + "_" + item.AppVersion + ".sppkg"

                // Log
                ErrorDialog.logInfo(`Renaming the archived package ${fileName}...`);

                // Copy the file to the archive folder
                archiveFolder.Files().add(fileName, true, content).execute(
                    file => {
                        // Log
                        ErrorDialog.logInfo("The app file was archived successfully...");

                        // Hide the dialog
                        LoadingDialog.hide();

                        // Call the event
                        onComplete();
                    },
                    ex => {
                        // Log the error
                        ErrorDialog.show("Archiving Package", "There was an error archiving the file.", ex);
                    }
                );
            });
        });
    }

    // Creates the archive folder
    private static createArchiveFolder(rootFolder: Types.SP.FolderOData): PromiseLike<Types.SP.IFolder> {
        // Log
        ErrorDialog.logInfo("Getting the archive folder...");

        // Return a promise
        return new Promise(resolve => {
            // Find the archive folder
            for (let i = 0; i < rootFolder.Folders.results.length; i++) {
                let folder = rootFolder.Folders.results[i];

                // See if this is the archive folder
                if (folder.Name.toLowerCase() == "archive") {
                    // Log
                    ErrorDialog.logInfo("Archive folder already exists...");

                    // Resolve the request
                    resolve(Web(Strings.SourceUrl).getFolderByServerRelativeUrl(folder.ServerRelativeUrl));
                    return;
                }
            }

            // Log
            ErrorDialog.logInfo("Creating the archive folder...");

            // Create the folder
            Web(Strings.SourceUrl).getFolderByServerRelativeUrl(rootFolder.ServerRelativeUrl).addSubFolder("archive").execute(
                folder => {
                    // Log
                    ErrorDialog.logInfo("Created the archive folder successfully...");

                    // Resolve the request
                    resolve(Web(Strings.SourceUrl).getFolderByServerRelativeUrl(rootFolder.ServerRelativeUrl + "/archive"));
                },
                ex => {
                    // Log the error
                    ErrorDialog.show("Archive Folder", "There was an error creating the archive folder.", ex);
                }
            );
        });
    }

    // Configures the test site collection
    static configureTestSite(webUrl: string, item: IAppItem): PromiseLike<void> {
        // Log
        ErrorDialog.logInfo("Configuring the Test Web...");

        // Show a loading dialog
        LoadingDialog.setHeader("Configuring the Test Web");
        LoadingDialog.setBody("Configuring the permissions of the test web.");
        LoadingDialog.show();

        // Return a promise
        return new Promise(resolve => {
            // Log
            ErrorDialog.logInfo(`Getting the web context information for ${AppConfig.Configuration.appCatalogUrl}...`);

            // Load the context of the app catalog
            ContextInfo.getWeb(AppConfig.Configuration.appCatalogUrl).execute(
                context => {
                    let web = Web(webUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue });

                    // Log
                    ErrorDialog.logInfo(`Resetting the web's permissions to inherit from the parent...`);

                    // Reset the permissions
                    web.resetRoleInheritance().execute();

                    // Log
                    ErrorDialog.logInfo(`Breaking the web's permissions to not inherit...`);

                    // Clear the permissions
                    web.breakRoleInheritance(true, false).execute(true);

                    // Log
                    ErrorDialog.logInfo(`Adding the developers to the web as 'Owners'...`);

                    // Parse the developers
                    let developers = item.AppDevelopers.results || [];
                    for (let i = 0; i < developers.length; i++) {
                        // Ensure the user exists in this site collection
                        web.ensureUser(developers[i].EMail).execute(user => {
                            // Update the developers user id
                            for (let j = 0; j < developers.length; j++) {
                                // See if this is the target user
                                if (developers[j].EMail == user.Email) {
                                    // Log
                                    ErrorDialog.logInfo(`Adding developer ${developers[j].EMail} as an owner...`);

                                    // Update the id
                                    developers[j].Id = user.Id;
                                    break;
                                }
                            }
                        }, true);
                    }

                    // Log
                    ErrorDialog.logInfo(`Getting the web's admin role definition...`);

                    // Get the owner role definition
                    web.RoleDefinitions().getByType(SPTypes.RoleType.Administrator).execute(role => {
                        // Log
                        ErrorDialog.logInfo(`Updating the web's permissions...`);

                        // Parse the developers
                        for (let i = 0; i < developers.length; i++) {
                            // Add the user
                            web.RoleAssignments().addRoleAssignment(developers[i].Id, role.Id).execute(true);
                        }

                        // Wait for the requests to complete
                        web.done(() => {
                            // Log
                            ErrorDialog.logInfo(`Web's permissions were updated successfully...`);

                            // Hide the loading dialog
                            LoadingDialog.hide();

                            // Resolve the request
                            resolve();
                        });
                    }, true);
                },
                ex => {
                    // Log the error
                    ErrorDialog.show("Getting Context", "There was an error getting the web context", ex);
                }
            );
        });
    }

    // Creates the test site for the application
    static createTestSite(item: IAppItem, onComplete: () => void) {
        // Log
        ErrorDialog.logInfo(`Creating the test site...`);

        // Show a loading dialog
        LoadingDialog.setHeader("Deploying the Application");
        LoadingDialog.setBody("Adding the application to the test site collection app catalog.");
        LoadingDialog.show();

        // Deploy the solution
        // Force the skip feature deployment to be false for a test site.
        this.deploy(item, false, false, () => {
            // Update the loading dialog
            LoadingDialog.setHeader("Creating the Test Site");
            LoadingDialog.setBody("Creating the sub-web for testing the application.");
            LoadingDialog.show();

            // Log
            ErrorDialog.logInfo(`Getting the web context information for ${AppConfig.Configuration.appCatalogUrl}...`);

            // Load the context of the app catalog
            ContextInfo.getWeb(AppConfig.Configuration.appCatalogUrl).execute(
                context => {
                    let requestDigest = context.GetContextWebInformation.FormDigestValue;

                    // Log
                    ErrorDialog.logInfo(`Creating the test web for app: ${item.AppProductID}...`);

                    // Create the test site
                    Web(AppConfig.Configuration.appCatalogUrl, { requestDigest }).WebInfos().add({
                        Description: "The test site for the " + item.Title + " application.",
                        Title: item.Title,
                        Url: item.AppProductID,
                        WebTemplate: SPTypes.WebTemplateType.Site
                    }).execute(
                        // Success
                        web => {
                            // Log
                            ErrorDialog.logInfo(`Creating the test web '${web.ServerRelativeUrl}' successfully...`);

                            // Method to send an email
                            let sendEmail = () => {
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
                                            "<p>The '" + item.Title + "' app test site has been created. Click the link below to access the site:</p>" +
                                            "<p>" + window.location.origin + web.ServerRelativeUrl + "</p>" +
                                            "<p>App Team</p>"
                                    }).execute(
                                        () => {
                                            // Close the dialog
                                            LoadingDialog.hide();

                                            // Notify the parent this process is complete
                                            onComplete();
                                        },
                                        ex => {
                                            // Log the error
                                            ErrorDialog.show("Sending Email", "There was an error sending the email notification.", ex);
                                        }
                                    );
                                } else {
                                    // Close the dialog
                                    LoadingDialog.hide();

                                    // Notify the parent this process is complete
                                    onComplete();
                                }
                            }

                            // Configure the test site
                            this.configureTestSite(web.ServerRelativeUrl, item).then(() => {
                                // Get the app
                                Web(web.ServerRelativeUrl, { requestDigest }).SiteCollectionAppCatalog().AvailableApps(item.AppProductID).execute(
                                    app => {
                                        // See if the app is already installed
                                        if (app.SkipDeploymentFeature) {
                                            // Log
                                            ErrorDialog.logInfo(`App has 'Skip Feature Deployment' flag set, skipping the install...`);

                                            // Send the email
                                            sendEmail();
                                        } else {
                                            // Log
                                            ErrorDialog.logInfo(`Installing the app '${item.Title}' with id ${item.AppProductID}...`);

                                            // Update the loading dialog
                                            LoadingDialog.setHeader("Installing the App");
                                            LoadingDialog.setBody("Installing the application to the test site.");

                                            // Install the application to the test site
                                            app.install().execute(
                                                // Success
                                                () => {
                                                    // Log
                                                    ErrorDialog.logInfo(`The app '${item.Title}' with id ${item.AppProductID} was installed successfully...`);

                                                    // Send the email
                                                    sendEmail();
                                                },

                                                // Error
                                                () => {
                                                    // Log the error
                                                    ErrorDialog.show("Error Deploying Application", "There was an error installing the application to the test site.");
                                                }
                                            );
                                        }
                                    },
                                    ex => {
                                        // Log the error
                                        ErrorDialog.show("Getting App", "There was an error getting the app item.", ex);
                                    }
                                );
                            });
                        },
                        // Error
                        () => {
                            // Log the error
                            ErrorDialog.show("Error Creating Test Site", "There was an error creating the test site.");
                        }
                    );
                },
                ex => {
                    // Log the error
                    ErrorDialog.show("Getting Context", "There was an error getting the web context", ex);
                }
            );
        });
    }

    // Method to delete the test site
    static deleteTestSite(item: IAppItem): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Log
            ErrorDialog.logInfo(`Deleting the test site for app '${item.Title}' with id ${item.AppProductID}...`);

            // Update the loading dialog
            LoadingDialog.setHeader("Deleting the Test Site");
            LoadingDialog.setBody("Removing the test site created for this app.");
            LoadingDialog.show();

            // Delete the test site
            DataSource.deleteTestSite(item).then(() => {
                // Close the dialog
                LoadingDialog.hide();

                // Resolve the request
                resolve();
            }, reject);
        });
    }

    // Deploys the solution to the app catalog
    static deploy(item: IAppItem, tenantFl: boolean, skipFeatureDeployment: boolean, onUpdate: () => void) {
        let appFile: Types.SP.File = null;

        // Log
        ErrorDialog.logInfo(`Deploying the SPFx package to the ${tenantFl ? "tenant" : "site"} app catalog...`);

        // Show a loading dialog
        LoadingDialog.setHeader("Uploading Package");
        LoadingDialog.setBody("Uploading the spfx package to the app catalog.");
        LoadingDialog.show();

        // Log
        ErrorDialog.logInfo(`Getting the SPFx package...`);

        // Find the package file
        for (let i = 0; i < DataSource.DocSetFolder.Files.results.length; i++) {
            let file = DataSource.DocSetFolder.Files.results[i];

            // See if this is the package
            if (file.Name.toLowerCase().endsWith(".sppkg")) {
                // Set the file
                appFile = file;
                break;
            }
        }

        // Ensure a file exists
        if (appFile == null) {
            // Log
            ErrorDialog.logInfo(`The SPFx package was not found...`);

            // This shouldn't happen
            LoadingDialog.hide();
            return;
        }

        // Log
        ErrorDialog.logInfo(`Getting the SPFx package's file content from ${appFile.ServerRelativeUrl}...`);

        // Get the package file contents
        Web(Strings.SourceUrl).getFileByServerRelativeUrl(appFile.ServerRelativeUrl).content().execute(
            content => {
                let catalogUrl = tenantFl ? AppConfig.Configuration.tenantAppCatalogUrl : AppConfig.Configuration.appCatalogUrl;

                // Log
                ErrorDialog.logInfo(`Getting the web context for ${catalogUrl}...`);

                // Load the context of the app catalog
                ContextInfo.getWeb(catalogUrl).execute(
                    context => {
                        let requestDigest = context.GetContextWebInformation.FormDigestValue;
                        let web = Web(catalogUrl, { requestDigest });

                        // Retract the app
                        this.retract(item, tenantFl, true, () => {
                            // Log
                            ErrorDialog.logInfo(`Uploading the app '${item.Title}' with id ${item.AppProductID} to the catalog...`);

                            // Update the dialog
                            LoadingDialog.setHeader("Uploading the Package");
                            LoadingDialog.setBody("This will close after the app is deployed.");
                            LoadingDialog.show();

                            // Upload the file to the app catalog
                            (tenantFl ? web.TenantAppCatalog() : web.SiteCollectionAppCatalog()).add(appFile.Name, true, content).execute(
                                file => {
                                    // Log
                                    ErrorDialog.logInfo(`The app '${item.Title}' with id ${item.AppProductID} was added to the catalog successfully...`);

                                    // Update the dialog
                                    LoadingDialog.setHeader("Deploying the Package");
                                    LoadingDialog.setBody("This will close after the app is deployed.");

                                    // Get the app item
                                    file.ListItemAllFields().execute(
                                        appItem => {
                                            // Update the metadata
                                            this.updateAppMetadata(item, appItem.Id, tenantFl, catalogUrl, requestDigest).then(() => {
                                                // Get the app catalog
                                                let web = Web(catalogUrl, { requestDigest });
                                                let appCatalog = (tenantFl ? web.TenantAppCatalog() : web.SiteCollectionAppCatalog());

                                                // Log
                                                ErrorDialog.logInfo(`Deploying the app '${item.Title}' with id ${item.AppProductID}...`);

                                                // Deploy the app
                                                appCatalog.AvailableApps(item.AppProductID).deploy(skipFeatureDeployment).execute(app => {
                                                    // Log
                                                    ErrorDialog.logInfo(`The app '${item.Title}' with id ${item.AppProductID} was deployed successfully...`);

                                                    // See if this is the tenant app
                                                    if (tenantFl) {
                                                        // Log
                                                        ErrorDialog.logInfo(`Setting the app '${item.Title}' with id ${item.AppProductID} tenant deployed flag...`);

                                                        // Update the tenant deployed flag
                                                        item.update({
                                                            AppIsTenantDeployed: true
                                                        }).execute(() => {
                                                            // Log
                                                            ErrorDialog.logInfo(`The app '${item.Title}' with id ${item.AppProductID} tenant deployed flag was set to true...`);

                                                            // Hide the dialog
                                                            LoadingDialog.hide();

                                                            // Call the update event
                                                            onUpdate();
                                                        }, () => {
                                                            // Log the error
                                                            ErrorDialog.show("Updating App", "There was an error setting the tenant deployed flag.");
                                                        });
                                                    } else {
                                                        // Hide the dialog
                                                        LoadingDialog.hide();

                                                        // Call the update event
                                                        onUpdate();
                                                    }
                                                }, () => {
                                                    // Log the error
                                                    ErrorDialog.show("Getting Apps", "There was an error getting the available apps from the app catalog.");

                                                    // Error deploying the app
                                                    // TODO - Show an error
                                                    // Call the update event
                                                    onUpdate();
                                                });
                                            });
                                        },
                                        ex => {
                                            // Log the error
                                            ErrorDialog.show("Loading App", "There was an error loading the app item.", ex);
                                        }
                                    );
                                },
                                ex => {
                                    // Log the error
                                    ErrorDialog.show("Adding App", "There was an error adding the app to the app catalog.", ex);
                                }
                            );
                        });
                    },
                    ex => {
                        // Log the error
                        ErrorDialog.show("Getting Context", "There was an error getting the web context", ex);
                    }
                );
            },
            ex => {
                // Log the error
                ErrorDialog.show("Reading File", "There was an error reading the app package file.", ex);
            }
        );
    }

    // Deploys the app to a site collection app catalog
    static deployAppToSite(item: IAppItem, siteUrl: string, updateSitesFl: boolean, onUpdate: () => void) {
        let appFile: Types.SP.File = null;

        // Log
        ErrorDialog.logInfo(`Getting the SPFx app package...`);

        // Show a loading dialog
        LoadingDialog.setHeader("Uploading Package");
        LoadingDialog.setBody("Uploading the spfx package to the site collection app catalog.");
        LoadingDialog.show();

        // Find the package file
        for (let i = 0; i < DataSource.DocSetFolder.Files.results.length; i++) {
            let file = DataSource.DocSetFolder.Files.results[i];

            // See if this is the package
            if (file.Name.toLowerCase().endsWith(".sppkg")) {
                // Set the file
                appFile = file;
                break;
            }
        }

        // Ensure a file exists
        if (appFile == null) {
            // Log
            ErrorDialog.logInfo(`Unable to find the SPFx app package...`);

            // This shouldn't happen
            LoadingDialog.hide();
            return;
        }

        // Log
        ErrorDialog.logInfo(`Getting the SPFx app package contents...`);

        // Upload the package file
        Web(Strings.SourceUrl).getFileByServerRelativeUrl(appFile.ServerRelativeUrl).content().execute(
            content => {
                // Log
                ErrorDialog.logInfo(`Getting the web context for: ${siteUrl}`);

                // Load the context of the app catalog
                ContextInfo.getWeb(siteUrl).execute(
                    context => {
                        let requestDigest = context.GetContextWebInformation.FormDigestValue;

                        // Retract the app
                        this.retract(item, false, true, () => {
                            // Log
                            ErrorDialog.logInfo(`Uploading the SPFx app package '${appFile.Name}' to the app catalog...`);

                            // Update the dialog
                            LoadingDialog.setHeader("Uploading the Package");
                            LoadingDialog.setBody("This will close after the app is uploaded.");
                            LoadingDialog.show();

                            // Upload the file to the app catalog
                            Web(siteUrl, { requestDigest }).SiteCollectionAppCatalog().add(appFile.Name, true, content).execute(
                                file => {
                                    // Log
                                    ErrorDialog.logInfo(`Uploading the SPFx app package '${appFile.Name}' to the app catalog...`);

                                    // Update the dialog
                                    LoadingDialog.setHeader("Deploying the Package");
                                    LoadingDialog.setBody("This will close after the app is deployed.");

                                    // Get the app item
                                    file.ListItemAllFields().execute(
                                        appItem => {
                                            // Update the metadata
                                            this.updateAppMetadata(item, appItem.Id, false, siteUrl, requestDigest).then(() => {
                                                // Log
                                                ErrorDialog.logInfo(`Deploying the SPFx app '${item.Title}' with id: ${item.AppProductID}`);

                                                // Deploy the app
                                                Web(siteUrl, { requestDigest }).SiteCollectionAppCatalog().AvailableApps(item.AppProductID).deploy(item.AppSkipFeatureDeployment).execute(app => {
                                                    // Send the notifications
                                                    AppNotifications.sendAppDeployedEmail(item, context.GetContextWebInformation.WebFullUrl).then(() => {
                                                        // See if we are updating the metadata
                                                        if (updateSitesFl) {
                                                            // Append the url to the list of sites the solution has been deployed to
                                                            let sites = (item.AppSiteDeployments || "").trim();

                                                            // Log
                                                            ErrorDialog.logInfo(`Updating the site deployments metadata...`);
                                                            ErrorDialog.logInfo(`Current Value: ${item.AppProductID}`);

                                                            // Ensure it doesn't contain the url already
                                                            if (sites.indexOf(context.GetContextWebInformation.WebFullUrl) < 0) {
                                                                // Append the url
                                                                sites += (sites.length > 0 ? "\r\n" : "") + context.GetContextWebInformation.WebFullUrl;
                                                            }

                                                            // Log
                                                            ErrorDialog.logInfo(`New Value: ${sites}`);

                                                            // Update the metadata
                                                            item.update({
                                                                AppSiteDeployments: sites
                                                            }).execute(() => {
                                                                // Log
                                                                ErrorDialog.logInfo(`Site deployments metadata field updated successfully...`);

                                                                // Hide the dialog
                                                                LoadingDialog.hide();

                                                                // Call the update event
                                                                onUpdate();
                                                            });
                                                        } else {
                                                            // Log
                                                            ErrorDialog.logInfo(`Error updating the site deployments metadata field...`);

                                                            // Hide the dialog
                                                            LoadingDialog.hide();

                                                            // Call the update event
                                                            onUpdate();
                                                        }
                                                    });
                                                }, ex => {
                                                    // Log the error
                                                    ErrorDialog.show("Deploying App", "There was an error deploying the app.", ex);

                                                    // Error deploying the app
                                                    // TODO - Show an error
                                                    // Call the update event
                                                    onUpdate();
                                                });
                                            });
                                        },
                                        ex => {
                                            // Log the error
                                            ErrorDialog.show("Reading Fields", "There was an error reading the file metadata.", ex);
                                        }
                                    );
                                },
                                ex => {
                                    // Log the error
                                    ErrorDialog.show("Uploading File", "There was an error uploading the app package file.", ex);
                                }
                            );
                        });
                    },
                    ex => {
                        // Log the error
                        ErrorDialog.show("Getting Context", "There was an error getting the web context", ex);
                    }
                );
            },
            ex => {
                // Log the error
                ErrorDialog.show("Reading File", "There was an error reading the app package file.", ex);
            }
        );
    }

    // Deploys the solution to the app catalog
    static deployToSite(item: IAppItem, siteUrl: string, onUpdate: () => void) {
        // Deploy the app to the site
        this.deployAppToSite(item, siteUrl, true, () => {
            // Call the update event
            onUpdate();
        });
    }

    // Deploys the solution to teams
    static deployToTeams(item: IAppItem, onComplete: () => void) {
        // Log
        ErrorDialog.logInfo(`Deploying app '${item.Title}' with id ${item.AppProductID} to Teams...`);

        // Show a loading dialog
        LoadingDialog.setHeader("Deploying to Teams");
        LoadingDialog.setBody("Syncing the app to Teams.");
        LoadingDialog.show();

        // Log
        ErrorDialog.logInfo(`Getting the tenant app catalog...`);

        // Load the context of the app catalog
        ContextInfo.getWeb(AppConfig.Configuration.tenantAppCatalogUrl).execute(context => {
            let requestDigest = context.GetContextWebInformation.FormDigestValue;

            // Log
            ErrorDialog.logInfo(`Syncing app '${item.Title}' with id ${item.AppProductID} to Teams...`);

            // Sync the app with Teams
            Web(AppConfig.Configuration.tenantAppCatalogUrl, { requestDigest }).TenantAppCatalog().syncSolutionToTeams(item.Id).execute(
                // Success
                () => {
                    // Log
                    ErrorDialog.logInfo(`App '${item.Title}' with id ${item.AppProductID} was successfully synced to Teams...`);

                    // Hide the dialog
                    LoadingDialog.hide();

                    // Notify the parent this process is complete
                    onComplete();
                },

                // Error
                ex => {
                    // Log the error
                    ErrorDialog.show("Deploy to Teams", "There was an error deploying the app to the tenant app catalog.", ex);
                }
            );
        }, ex => {
            // Log
            ErrorDialog.logInfo(`Getting the tenant app catalog...`);
        });
    }

    // Reads an app package file
    static readPackage(data): PromiseLike<{ item: IAppItem; image: JSZip.JSZipObject }> {
        // Return a promise
        return new Promise(resolve => {
            // Log
            ErrorDialog.logInfo(`Unzipping the SPFx package...`);

            // Unzip the package
            JSZip.loadAsync(data).then(files => {
                let metadata: IAppItem = {} as any;
                let image = null;

                // Log
                ErrorDialog.logInfo(`Parsing the SPFx package files...`);

                // Parse the files
                files.forEach((path, fileInfo) => {
                    /** What are we doing here? */
                    /** Should we convert this to use the doc set dashboard solution? */

                    // Get the file name
                    let fileName = fileInfo.name.toLowerCase();

                    // Log
                    ErrorDialog.logInfo(`Processing SPFx package file: ${fileName}`);

                    // See if this is an image
                    if (fileName.endsWith(".png") || fileName.endsWith(".jpg") || fileName.endsWith(".jpeg") || fileName.endsWith(".gif")) {
                        // Log
                        ErrorDialog.logInfo(`${fileName}: Image file found...`);

                        // Save a reference to the iamge
                        image = fileInfo;

                        // Check the next file
                        return;
                    }

                    // See if this is the app manifest
                    if (fileName != "appmanifest.xml") { return; }

                    // Log
                    ErrorDialog.logInfo(`${fileName}: App Manifest found. Processing the information...`);

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

                        // Log
                        ErrorDialog.logInfo(`App Information Extracted: \r\n${JSON.stringify(metadata)}`);

                        // Resolve the request
                        resolve({ item: metadata, image });
                    });
                });
            });
        });
    }

    // Retracts the solution from the app catalog
    static retract(item: IAppItem, tenantFl: boolean, removeFl: boolean, onUpdate: () => void) {
        // Log
        ErrorDialog.logInfo(`Retracting the SPFx package...`);

        // Show a loading dialog
        LoadingDialog.setHeader("Retracting the Package");
        LoadingDialog.setBody("Retracting the spfx package from the app catalog.");
        LoadingDialog.show();

        // Get the app catalog url
        let catalogUrl = tenantFl ? AppConfig.Configuration.tenantAppCatalogUrl : AppConfig.Configuration.appCatalogUrl;

        // Log
        ErrorDialog.logInfo(`Getting the web's context at ${catalogUrl}...`);

        // Load the context of the app catalog
        ContextInfo.getWeb(catalogUrl).execute(
            context => {
                // Log
                ErrorDialog.logInfo(`Loading the SPFx apps in the catalog...`);

                // Log
                ErrorDialog.logInfo(`Retracting the app '${item.Title}' with id ${item.AppProductID} from the app catalog...`);

                // Load the apps
                let web = Web(catalogUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue });
                (tenantFl ? web.TenantAppCatalog() : web.SiteCollectionAppCatalog()).AvailableApps(item.AppProductID).retract().execute(() => {
                    // Log
                    ErrorDialog.logInfo(`The app '${item.Title}' with id ${item.AppProductID} was retracted from the app catalog successfully...`);

                    // See if we are removing the app
                    if (removeFl) {
                        // Log
                        ErrorDialog.logInfo(`Removing the SPFx app from the catalog...`);

                        // Remove the app
                        (tenantFl ? web.TenantAppCatalog() : web.SiteCollectionAppCatalog()).AvailableApps(item.AppProductID).remove().execute(
                            () => {
                                // Log
                                ErrorDialog.logInfo(`The app '${item.Title}' with id ${item.AppProductID} was removed from the app catalog successfully...`);

                                // Close the dialog
                                LoadingDialog.hide();

                                // Call the update event
                                onUpdate();
                            },
                            ex => {
                                // Log the error
                                ErrorDialog.show("Removing App", "There was an error removing the app from the catalog.", ex);
                            }
                        );
                    } else {
                        // Close the dialog
                        LoadingDialog.hide();

                        // Call the update event
                        onUpdate();
                    }
                }, () => {
                    // Log
                    ErrorDialog.logInfo(`Error retracting the app '${item.Title} with id ${item.AppProductID} from the app catalog...`);

                    // Close the dialog
                    LoadingDialog.hide();

                    // Call the update event
                    onUpdate();
                });
            },
            ex => {
                // Log the error
                ErrorDialog.show("Getting Context", "There was an error getting the web context", ex);
            }
        );
    }

    // Updates the app
    static updateApp(item: IAppItem, siteUrl: string, isTestSite: boolean): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // Log
            ErrorDialog.logInfo(`Getting the web context information at ${siteUrl}...`);

            // Load the context of the app catalog
            ContextInfo.getWeb(siteUrl).execute(
                context => {
                    let requestDigest = context.GetContextWebInformation.FormDigestValue;

                    // Update the dialog
                    LoadingDialog.setHeader("Uninstalling the Solution");
                    LoadingDialog.setBody("Uninstalling in site: " + siteUrl + "<br/>This will close after the app is upgraded.");
                    LoadingDialog.show();

                    // Log
                    ErrorDialog.logInfo(`Uninstalling the app '${item.Title}' with id ${item.AppProductID} from the app catalog...`);

                    // Uninstall the app
                    Web(siteUrl, { requestDigest }).SiteCollectionAppCatalog().AvailableApps(item.AppProductID).uninstall().execute(
                        () => {
                            // Update the dialog
                            LoadingDialog.setHeader("Upgrading the Solution");

                            // Log
                            ErrorDialog.logInfo(`Upgrading the app '${item.Title}' with id ${item.AppProductID}...`);

                            // Deploy the solution
                            this.deploy(item, false, isTestSite ? false : item.AppSkipFeatureDeployment, () => {
                                // Update the dialog
                                LoadingDialog.setHeader("Upgrading the Solution");

                                // Log
                                ErrorDialog.logInfo(`The app '${item.Title}' with id ${item.AppProductID} was upgraded successfully...`);

                                // Log
                                ErrorDialog.logInfo(`Getting the app from the catalog...`);

                                // Get the app
                                Web(siteUrl, { requestDigest }).SiteCollectionAppCatalog().AvailableApps(item.AppProductID).execute(
                                    app => {
                                        // Log
                                        ErrorDialog.logInfo(`The app '${item.Title}' with id ${item.AppProductID} was found in the app catalog...`);

                                        // See if the app is already installed
                                        if (app.SkipDeploymentFeature) {
                                            // Resolve the requst
                                            resolve();
                                        } else {
                                            // Log
                                            ErrorDialog.logInfo(`Installing the app '${item.Title}' with id ${item.AppProductID}...`);

                                            // Update the dialog
                                            LoadingDialog.setHeader("Installing the Solution");
                                            LoadingDialog.setBody("Installing in site: " + siteUrl + "<br/>This will close after the app is upgraded.");
                                            LoadingDialog.show();

                                            // Install the app
                                            Web(siteUrl, { requestDigest }).SiteCollectionAppCatalog().AvailableApps(item.AppProductID).install().execute(
                                                () => {
                                                    // Log
                                                    ErrorDialog.logInfo(`The app '${item.Title}' with id ${item.AppProductID} was installed successfully...`);

                                                    // Close the dialog
                                                    LoadingDialog.hide();

                                                    // Resolve the requst
                                                    resolve();
                                                },
                                                ex => {
                                                    // Log the error
                                                    ErrorDialog.show("Installing App", "There was an error installing the app.", ex);
                                                }
                                            );
                                        }
                                    },
                                    ex => {
                                        // Log the error
                                        ErrorDialog.show("Getting App", "There was an error loading the app.", ex);
                                    }
                                );
                            });
                        },
                        ex => {
                            // Log the error
                            ErrorDialog.show("Uninstall App", "There was an error uninstalling the app.", ex);
                        }
                    );
                },
                ex => {
                    // Log the error
                    ErrorDialog.show("Getting Context", "There was an error getting the web context", ex);
                }
            );
        });
    }

    // Updates the app metadata
    static updateAppMetadata(item: IAppItem, appItemId: number, tenantFl: boolean, catalogUrl: string, requestDigest: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Log
            ErrorDialog.logInfo(`Getting the app catalog list from web: ${catalogUrl}`);

            // Get the app catalog
            Web(catalogUrl, { requestDigest }).Lists().query({
                Filter: "BaseTemplate eq " + (tenantFl ? SPTypes.ListTemplateType.TenantAppCatalog : SPTypes.ListTemplateType.AppCatalog)
            }).execute(
                lists => {
                    // Log
                    ErrorDialog.logInfo(`Updating the app metadata...`);

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
                        }).execute(
                            () => {
                                // Log
                                ErrorDialog.logInfo(`The app metadata was updated succesfully...`);

                                // Resolve the request
                                resolve();
                            },
                            ex => {
                                // Log the error
                                ErrorDialog.show("Updating Metadata", "There was an error updating the app's metadata.", ex);
                            }
                        );
                    } else {
                        // Log
                        ErrorDialog.logInfo(`The app catalog was not found on the site: ${catalogUrl}`);

                        // Resolve the request
                        resolve();
                    }
                },
                ex => {
                    // Log the error
                    ErrorDialog.show("Loading Lists", "There was an error loading the lists.", ex);
                }
            );
        });
    }

    // Upload form
    static upload(onComplete: (item: IAppItem) => void) {
        // Show the upload file dialog
        Helper.ListForm.showFileDialog().then(file => {
            // Log
            ErrorDialog.logInfo(`File uploaded: ${file.name}`);

            // Ensure this is an spfx package
            if (file.name.toLowerCase().endsWith(".sppkg")) {
                // Log
                ErrorDialog.logInfo(`File ${file.name} uploaded is a SPFx package...`);

                // Display a loading dialog
                LoadingDialog.setHeader("Reading Package");
                LoadingDialog.setBody("Validating the package...");
                LoadingDialog.show();

                // Extract the metadata from the package
                this.readPackage(file.data).then(pkgInfo => {
                    // Ensure the package doesn't exist
                    let existsFl = false;
                    for (let i = 0; i < DataSource.Items.length; i++) {
                        let item = DataSource.Items[i];

                        // See if it exists
                        if (item.AppProductID == pkgInfo.item.AppProductID) {
                            // Set the flag
                            existsFl = true;
                            break;
                        }
                    }

                    // See if the app exists
                    if (existsFl) {
                        // Log
                        ErrorDialog.logInfo(`SPFx ${file.name} package already exists in the app catalog tool...`);

                        // Close the loading dialog
                        LoadingDialog.hide();

                        // Display a modal
                        Modal.clear();
                        Modal.setHeader("Package Validation Error");
                        Modal.setBody("The spfx package already exists. Please update the existing application through it's details page.");
                        Modal.show();
                    }
                    // Else, validate the data extracted from the app
                    else if (pkgInfo.item.AppProductID && pkgInfo.item.AppVersion && pkgInfo.item.Title) {
                        // Log
                        ErrorDialog.logInfo(`Creating the document set folder for: ${file.name}...`);

                        // Update the loading dialog
                        LoadingDialog.setHeader("Creating App Folder");
                        LoadingDialog.setBody("Creating the app folder...");

                        // Create the document set folder
                        Helper.createDocSet(pkgInfo.item.Title, Strings.Lists.Apps).then(
                            // Success
                            item => {
                                // Log
                                ErrorDialog.logInfo(`Document set folder for ${pkgInfo.item.Title} was created succesfully...`);

                                // Update the loading dialog
                                LoadingDialog.setHeader("Updating Metadata");
                                LoadingDialog.setBody("Saving the package information...");

                                // Default the owner to the current user
                                pkgInfo.item.AppDevelopersId = { results: [ContextInfo.userId] } as any;

                                // Log
                                ErrorDialog.logInfo(`Updating the item metadata...`);

                                // Update the metadata
                                item.update(pkgInfo.item).execute(() => {
                                    // Log
                                    ErrorDialog.logInfo(`App's metadata was updated successfully...`);

                                    // Update the loading dialog
                                    LoadingDialog.setHeader("Uploading the Package");
                                    LoadingDialog.setBody("Uploading the app package...");

                                    // Log
                                    ErrorDialog.logInfo(`Uploading the SPFx package ${file.name} to the folder...`);

                                    // Upload the file
                                    item.Folder().Files().add(file.name, true, file.data).execute(
                                        // Success
                                        file => {
                                            // Log
                                            ErrorDialog.logInfo(`The SPFx package was uploaded successfully...`);

                                            // See if the image exists
                                            if (pkgInfo.image) {
                                                // Log
                                                ErrorDialog.logInfo(`Uploading the SPFx package icon to the folder...`);

                                                // Get image in different format for later uploading
                                                pkgInfo.image.async("arraybuffer").then(function (content) {
                                                    // Get the file extension
                                                    let fileExt: string | string[] = pkgInfo.image.name.split('.');
                                                    fileExt = fileExt[fileExt.length - 1];

                                                    // Upload the image to this folder
                                                    item.Folder().Files().add("AppIcon." + fileExt, true, content).execute(
                                                        file => {
                                                            // Log
                                                            ErrorDialog.logInfo(`The SPFx app icon '${file.Name}' was added successfully...`);

                                                            // Set the icon url
                                                            item.update({
                                                                AppThumbnailURL: {
                                                                    __metadata: { type: "SP.FieldUrlValue" },
                                                                    Description: file.ServerRelativeUrl,
                                                                    Url: file.ServerRelativeUrl
                                                                }
                                                            }).execute(() => {
                                                                // Log
                                                                ErrorDialog.logInfo(`The SPFx app icon url metadata was updated successfully...`);

                                                                // Close the loading dialog
                                                                LoadingDialog.hide();

                                                                // Execute the completed event
                                                                onComplete(item as any);
                                                            });
                                                        },
                                                        ex => {
                                                            // Log the error
                                                            ErrorDialog.show("Uploading Icon", "There was an error uploading the app icon file.", ex);
                                                        }
                                                    );
                                                });
                                            } else {
                                                // Close the loading dialog
                                                LoadingDialog.hide();

                                                // Execute the completed event
                                                onComplete(item as any);
                                            }
                                        },

                                        // Error
                                        ex => {
                                            // Log the error
                                            ErrorDialog.show("Uploading File", "There was an error uploading the app package file.", ex);
                                        }
                                    );
                                });
                            },

                            // Error
                            ex => {
                                // Log the error
                                ErrorDialog.show("Creating DocSet Item", "There was an error creating the document set item for this app.", ex);
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
                // Log
                ErrorDialog.logInfo(`File ${file.name} uploaded is not a SPFx package...`);

                // Display a modal
                Modal.clear();
                Modal.setHeader("Error Adding Package");
                Modal.setBody("The file must be a valid SPFx app package file.");
                Modal.show();
            }
        });
    }
}