import { LoadingDialog, Modal } from "dattatable";
import { ContextInfo, Helper, List, SPTypes, Types, Utility, Web } from "gd-sprest-bs";
import * as JSZip from "jszip";
import { AppConfig } from "./appCfg";
import { AppSecurity } from "./appSecurity";
import * as Common from "./common";
import { DataSource, IAppItem } from "./ds";
import { ErrorDialog } from "./errorDialog";
import Strings from "./strings";


/**
 * Application Actions
 * The actions taken for an app.
 */
export class AppActions {
    // Archives the current package file
    static archivePackage(onComplete: () => void) {
        let appFile: Types.SP.File = null;

        // Show a loading dialog
        LoadingDialog.setHeader("Archiving the Package");
        LoadingDialog.setBody("The current app package is being archived.");
        LoadingDialog.show();

        // Log
        ErrorDialog.logInfo("Archiving the package....");

        // Find the package file
        for (let i = 0; i < DataSource.AppFolder.Files.results.length; i++) {
            let file = DataSource.AppFolder.Files.results[i];

            // See if this is the package
            let fileName = file.Name.toLowerCase();
            if (fileName.endsWith(".sppkg")) {
                // See if this is a prod file and skip it
                if (fileName.indexOf("-prod.sppkg") > 0) { continue; }
                // Else, see if this is a test file
                else if (fileName.indexOf("-test.sppkg") > 0) { continue; }
                // Else set the file
                else { appFile = file; }
            }
        }

        // Ensure a file exists
        if (appFile == null) {
            // Log
            ErrorDialog.logInfo("The app file was not found...");

            // This shouldn't happen
            LoadingDialog.hide();

            // Call the event to continue
            onComplete();
            return;
        }

        // Create the archive folder
        this.createArchiveFolder(DataSource.AppFolder).then(archiveFolder => {
            // Log
            ErrorDialog.logInfo(`Getting the SPFx package's file content from ${appFile.ServerRelativeUrl}...`);

            // Get the package file contents
            Web(Strings.SourceUrl).getFileByServerRelativeUrl(appFile.ServerRelativeUrl).content().execute(content => {
                // Update the dialog
                LoadingDialog.setBody("Copying the file package.")

                // Get the name of the file
                let fileName = appFile.Name.toLowerCase().split(".sppkg")[0] + "_" + DataSource.AppItem.AppVersion + ".sppkg"

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

    // Configures the test site collection
    static configureTestSite(webUrl: string): PromiseLike<void> {
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
                    let developers = DataSource.AppItem.AppDevelopers.results || [];
                    for (let i = 0; i < developers.length; i++) {
                        // Ensure the user exists in this site collection
                        web.ensureUser(developers[i].Name).execute(user => {
                            // Update the developers user id
                            for (let j = 0; j < developers.length; j++) {
                                // See if this is the target user
                                if (developers[j].Name == user.LoginName) {
                                    // Log
                                    ErrorDialog.logInfo(`Adding developer ${developers[j].Name} as an owner...`);

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

                        // Add the app sponsor
                        DataSource.AppItem.AppSponsorId > 0 ? web.RoleAssignments().addRoleAssignment(DataSource.AppItem.AppSponsorId, role.Id).execute() : null;

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

    // Creates the client side assets folder
    static createClientSideAssetsFolder(rootFolder: Types.SP.IFolder): PromiseLike<Types.SP.Folder> {
        // Log
        ErrorDialog.logInfo("Getting the client side assets folder...");

        // Return a promise
        return new Promise(resolve => {
            // Set the folder name
            let folderName = DataSource.AppItem.AppProductID.toLowerCase();

            // Load the folders
            rootFolder.Folders().execute(folders => {
                // Find the archive folder
                for (let i = 0; i < folders.results.length; i++) {
                    let folder = folders.results[i];

                    // See if this is the archive folder
                    if (folder.Name.toLowerCase() == folderName) {
                        // Log
                        ErrorDialog.logInfo("ClientSideAssets folder already exists...");

                        // Resolve the request
                        resolve(folder);
                        return;
                    }
                }

                // Log
                ErrorDialog.logInfo("Creating the client side assets folder...");

                // Create the folder
                rootFolder.Folders().add(folderName).execute(
                    folder => {
                        // Log
                        ErrorDialog.logInfo("Created the client side assets folder successfully...");

                        // Resolve the request
                        resolve(folder);
                    },
                    ex => {
                        // Log the error
                        ErrorDialog.show("ClientSideAssets Folder", "There was an error creating the client side assets folder.", ex);
                    }
                );
            });
        });
    }

    // Creates the image folder
    static createImageFolder(rootFolder: Types.SP.IFolder): PromiseLike<Types.SP.Folder> {
        // Log
        ErrorDialog.logInfo("Getting the image folder...");

        // Return a promise
        return new Promise(resolve => {
            // Set the folder name
            let folderName = DataSource.AppItem.AppProductID.toLowerCase();

            // Load the folders
            rootFolder.Folders().execute(folders => {
                // Find the archive folder
                for (let i = 0; i < folders.results.length; i++) {
                    let folder = folders.results[i];

                    // See if this is the archive folder
                    if (folder.Name.toLowerCase() == folderName) {
                        // Log
                        ErrorDialog.logInfo("Image folder already exists...");

                        // Resolve the request
                        resolve(folder);
                        return;
                    }
                }

                // Log
                ErrorDialog.logInfo("Creating the image folder...");

                // Create the folder
                rootFolder.Folders().add(folderName).execute(
                    folder => {
                        // Log
                        ErrorDialog.logInfo("Created the image folder successfully...");

                        // Resolve the request
                        resolve(folder);
                    },
                    ex => {
                        // Log the error
                        ErrorDialog.show("Image Folder", "There was an error creating the image folder.", ex);
                    }
                );
            });
        });
    }

    // Creates the test site for the application
    static createTestSite(onComplete: (web?: Types.SP.WebInformation) => void) {
        // Log
        ErrorDialog.logInfo(`Creating the test site...`);

        // Show a loading dialog
        LoadingDialog.setHeader("Deploying the Application");
        LoadingDialog.setBody("Adding the application to the test site collection app catalog.");
        LoadingDialog.show();

        // Deploy the solution
        // Force the skip feature deployment to be false for a test site.
        this.deploy("test", false, false, onComplete, () => {
            // Update the loading dialog
            LoadingDialog.setHeader("Creating the Test Site");
            LoadingDialog.setBody("Getting the web context information...");
            LoadingDialog.show();

            // Log
            ErrorDialog.logInfo(`Getting the web context information for ${AppConfig.Configuration.appCatalogUrl}...`);

            // Load the context of the app catalog
            ContextInfo.getWeb(AppConfig.Configuration.appCatalogUrl).execute(
                context => {
                    let requestDigest = context.GetContextWebInformation.FormDigestValue;

                    // Update the loading dialog
                    LoadingDialog.setBody("Creating the sub-web for testing the application...");

                    // Log
                    ErrorDialog.logInfo(`Creating the test web for app: ${DataSource.AppItem.AppProductID}...`);

                    // Create the test site
                    Web(AppConfig.Configuration.appCatalogUrl, { requestDigest }).WebInfos().add({
                        Description: "The test site for the " + DataSource.AppItem.Title + " application.",
                        Title: DataSource.AppItem.Title,
                        Url: DataSource.AppItem.AppProductID,
                        WebTemplate: SPTypes.WebTemplateType.Site
                    }).execute(
                        // Success
                        web => {
                            // Log
                            ErrorDialog.logInfo(`Creating the test web '${web.ServerRelativeUrl}' successfully...`);

                            // Configure the test site
                            this.configureTestSite(web.ServerRelativeUrl).then(() => {
                                // Get the app
                                Web(web.ServerRelativeUrl, { requestDigest }).SiteCollectionAppCatalog().AvailableApps(DataSource.AppItem.AppProductID).execute(
                                    app => {
                                        // See if the app is already installed
                                        if (app.SkipDeploymentFeature) {
                                            // Log
                                            ErrorDialog.logInfo(`App has 'Skip Feature Deployment' flag set, skipping the install...`);

                                            // Close the dialog
                                            LoadingDialog.hide();

                                            // Notify the parent this process is complete
                                            onComplete(web);
                                        } else {
                                            // Log
                                            ErrorDialog.logInfo(`Installing the app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID}...`);

                                            // Update the loading dialog
                                            LoadingDialog.setHeader("Installing the App");
                                            LoadingDialog.setBody("Installing the application to the test site.");
                                            LoadingDialog.show();

                                            // Install the application to the test site
                                            app.install().execute(
                                                // Success
                                                () => {
                                                    // Log
                                                    ErrorDialog.logInfo(`The app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} was installed successfully...`);

                                                    // Close the dialog
                                                    LoadingDialog.hide();

                                                    // Notify the parent this process is complete
                                                    onComplete(web);
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
    static deleteTestSite(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Log
            ErrorDialog.logInfo(`Deleting the test site for app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID}...`);

            // Update the loading dialog
            LoadingDialog.setHeader("Deleting the Test Site");
            LoadingDialog.setBody("Removing the test site created for this app.");
            LoadingDialog.show();

            // Delete the test site
            DataSource.deleteTestSite(DataSource.AppItem).then(() => {
                // Close the dialog
                LoadingDialog.hide();

                // Resolve the request
                resolve();
            }, reject);
        });
    }

    // Deploys the solution to the app catalog
    static deploy(pkgType: "" | "test" | "prod", tenantFl: boolean, skipFeatureDeployment: boolean, onError: () => void, onUpdate: () => void) {
        let appFile: Types.SP.File = null;
        let appProdFile: Types.SP.File = null;
        let appTestFile: Types.SP.File = null;

        // Log
        ErrorDialog.logInfo(`Deploying the SPFx package to the ${tenantFl ? "tenant" : "site"} app catalog...`);

        // Show a loading dialog
        LoadingDialog.setHeader("Uploading Package");
        LoadingDialog.setBody("Uploading the spfx package to the app catalog.");
        LoadingDialog.show();

        // Log
        ErrorDialog.logInfo(`Getting the SPFx package...`);

        // Find the package file
        for (let i = 0; i < DataSource.AppFolder.Files.results.length; i++) {
            let file = DataSource.AppFolder.Files.results[i];

            // See if this is the package
            let fileName = file.Name.toLowerCase();
            if (fileName.endsWith(".sppkg")) {
                // See if this is a prod file
                if (fileName.indexOf("-prod.sppkg") > 0) {
                    // Set the file
                    appProdFile = file;
                }
                // Else, see if this is a test file
                else if (fileName.indexOf("-test.sppkg") > 0) {
                    // Set the file
                    appTestFile = file;
                } else {
                    // Set the file
                    appFile = file;
                }
            }
        }

        // Set the target app file to use for the deployment
        switch (pkgType) {
            case "test":
                // Set the target app file
                appFile = appTestFile ? appTestFile : appFile;
                break;
            case "prod":
                // Set the target app file
                appFile = appProdFile ? appProdFile : appFile;
                break;
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
                        this.retract(tenantFl, true, () => {
                            // Log
                            ErrorDialog.logInfo(`Uploading the app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} to the catalog...`);

                            // Update the dialog
                            LoadingDialog.setHeader("Uploading the Package");
                            LoadingDialog.setBody("This will close after the app is deployed.");
                            LoadingDialog.show();

                            // Upload the file to the app catalog
                            (tenantFl ? web.TenantAppCatalog() : web.SiteCollectionAppCatalog()).add(appFile.Name, true, content).execute(
                                file => {
                                    // Log
                                    ErrorDialog.logInfo(`The app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} was added to the catalog successfully...`);

                                    // Update the dialog
                                    LoadingDialog.setHeader("Updating the Metadata");
                                    LoadingDialog.setBody("Updating the metadata in the app catalog...");

                                    // Get the app item
                                    file.ListItemAllFields().execute(
                                        appItem => {
                                            // Update the metadata
                                            this.updateAppMetadata(appItem.Id, tenantFl, catalogUrl, requestDigest).then(() => {
                                                // Upload the client side assets
                                                this.uploadClientSideAssets(pkgType).then(() => {
                                                    // Upload the app images
                                                    this.uploadImages(pkgType).then(() => {
                                                        // Get the app catalog
                                                        let web = Web(catalogUrl, { requestDigest });
                                                        let appCatalog = (tenantFl ? web.TenantAppCatalog() : web.SiteCollectionAppCatalog());

                                                        // Update the dialog
                                                        LoadingDialog.setHeader("Deploying the App");
                                                        LoadingDialog.setBody("Deploying the solution in the app catalog...");
                                                        LoadingDialog.show();

                                                        // Log
                                                        ErrorDialog.logInfo(`Deploying the app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID}...`);

                                                        // Deploy the app
                                                        appCatalog.AvailableApps(DataSource.AppItem.AppProductID).deploy(skipFeatureDeployment).execute(app => {
                                                            // Log
                                                            ErrorDialog.logInfo(`The app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} was deployed successfully...`);

                                                            // See if this is the tenant app
                                                            if (tenantFl) {
                                                                // Log
                                                                ErrorDialog.logInfo(`Setting the app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} tenant deployed flag...`);

                                                                // Update the dialog
                                                                LoadingDialog.setHeader("Updating the App");
                                                                LoadingDialog.setBody("Updating the app item...");

                                                                // Update the tenant deployed flag
                                                                DataSource.AppItem.update({
                                                                    AppIsTenantDeployed: true
                                                                }).execute(() => {
                                                                    // Log
                                                                    ErrorDialog.logInfo(`The app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} tenant deployed flag was set to true...`);

                                                                    // Hide the dialog
                                                                    LoadingDialog.hide();

                                                                    // Call the update event
                                                                    onUpdate();
                                                                }, () => {
                                                                    // Log the error
                                                                    ErrorDialog.show("Updating App", "There was an error setting the tenant deployed flag.");

                                                                    // Call the event
                                                                    onError();
                                                                });
                                                            } else {
                                                                // Hide the dialog
                                                                LoadingDialog.hide();

                                                                // Call the update event
                                                                onUpdate();
                                                            }
                                                        }, () => {
                                                            // See if this isn't the tenant
                                                            if (!tenantFl) {
                                                                // Update the dialog
                                                                LoadingDialog.setHeader("Loading the App Catalog");
                                                                LoadingDialog.setBody("Error deploying the app, getting the information...");
                                                                LoadingDialog.show();

                                                                // Load the site collection app item
                                                                DataSource.loadSiteAppByName(appFile.Name).then(appItem => {
                                                                    // See if the item exists
                                                                    if (appItem) {
                                                                        // Log the error
                                                                        ErrorDialog.show("Deploy Error", "The app was added to the catalog successfully, but there was an error with it.");
                                                                    } else {
                                                                        // Log the error
                                                                        ErrorDialog.show("Getting Apps", "There was an error getting the available apps from the app catalog.");
                                                                    }

                                                                    // Call the event
                                                                    onError();
                                                                });
                                                            } else {
                                                                // Log the error
                                                                ErrorDialog.show("Getting Apps", "There was an error getting the available apps from the app catalog.");

                                                                // Call the event
                                                                onError();
                                                            }
                                                        });
                                                    });
                                                });
                                            });
                                        },
                                        ex => {
                                            // Log the error
                                            ErrorDialog.show("Loading App", "There was an error loading the app item.", ex);

                                            // Call the event
                                            onError();
                                        }
                                    );
                                },
                                ex => {
                                    // Log the error
                                    ErrorDialog.show("Adding App", "There was an error adding the app to the app catalog.", ex);

                                    // Call the event
                                    onError();
                                }
                            );
                        });
                    },
                    ex => {
                        // Log the error
                        ErrorDialog.show("Getting Context", "There was an error getting the web context", ex);

                        // Call the event
                        onError();
                    }
                );
            },
            ex => {
                // Log the error
                ErrorDialog.show("Reading File", "There was an error reading the app package file.", ex);

                // Call the event
                onError();
            }
        );
    }

    // Deploys the app to a site collection app catalog
    static deployAppToSite(siteUrl: string, updateSitesFl: boolean, onUpdate: () => void) {
        let appFile: Types.SP.File = null;
        let appProdFile: Types.SP.File = null;

        // Log
        ErrorDialog.logInfo(`Getting the SPFx app package...`);

        // Show a loading dialog
        LoadingDialog.setHeader("Uploading Package");
        LoadingDialog.setBody("Uploading the spfx package to the site collection app catalog.");
        LoadingDialog.show();

        // Find the package file
        for (let i = 0; i < DataSource.AppFolder.Files.results.length; i++) {
            let file = DataSource.AppFolder.Files.results[i];

            // See if this is the package
            let fileName = file.Name.toLowerCase();
            if (fileName.endsWith(".sppkg")) {
                // See if this is a prod file
                if (fileName.indexOf("-prod.sppkg") > 0) {
                    // Set the file
                    appProdFile = file;
                }
                // Else, see if this is a test file and skip it
                else if (fileName.indexOf("-test.sppkg") > 0) { continue; }
                // Else, set the file
                else { appFile = file; }
            }
        }

        // Set the target app file
        appFile = appProdFile ? appProdFile : appFile;

        // Ensure a file exists
        if (appFile == null) {
            // Log
            ErrorDialog.logInfo(`The SPFx package was not found...`);

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
                        this.retract(false, true, () => {
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
                                            this.updateAppMetadata(appItem.Id, false, siteUrl, requestDigest).then(() => {
                                                // Log
                                                ErrorDialog.logInfo(`Deploying the SPFx app '${DataSource.AppItem.Title}' with id: ${DataSource.AppItem.AppProductID}`);

                                                // Deploy the app
                                                Web(siteUrl, { requestDigest }).SiteCollectionAppCatalog().AvailableApps(DataSource.AppItem.AppProductID).deploy(DataSource.AppItem.AppSkipFeatureDeployment).execute(app => {
                                                    // See if we are updating the metadata
                                                    if (updateSitesFl) {
                                                        // Append the url to the list of sites the solution has been deployed to
                                                        let sites = (DataSource.AppItem.AppSiteDeployments || "").trim();

                                                        // Log
                                                        ErrorDialog.logInfo(`Updating the site deployments metadata...`);
                                                        ErrorDialog.logInfo(`Current Value: ${DataSource.AppItem.AppProductID}`);

                                                        // Ensure it doesn't contain the url already
                                                        if (sites.indexOf(context.GetContextWebInformation.WebFullUrl) < 0) {
                                                            // Append the url
                                                            sites += (sites.length > 0 ? "\r\n" : "") + context.GetContextWebInformation.WebFullUrl;
                                                        }

                                                        // Log
                                                        ErrorDialog.logInfo(`New Value: ${sites}`);

                                                        // Update the metadata
                                                        DataSource.AppItem.update({
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
    static deployToSite(siteUrl: string, onUpdate: () => void) {
        // Deploy the app to the site
        this.deployAppToSite(siteUrl, true, () => {
            // Call the update event
            onUpdate();
        });
    }

    // Deploys the solution to teams
    static deployToTeams(onComplete: () => void) {
        // Log
        ErrorDialog.logInfo(`Deploying app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} to Teams...`);

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
            ErrorDialog.logInfo(`Syncing app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} to Teams...`);

            // Get the tenant app catalog item
            Web(AppConfig.Configuration.tenantAppCatalogUrl, { requestDigest }).Lists("Apps for SharePoint").Items().query({
                Filter: "AppProductID eq '" + DataSource.AppItem.AppProductID + "'"
            }).execute(items => {
                // Ensure the item exists
                let item = items.results[0];
                if (item) {
                    // Sync the app with Teams
                    Web(AppConfig.Configuration.tenantAppCatalogUrl, { requestDigest }).TenantAppCatalog().syncSolutionToTeams(item.Id).execute(
                        // Success
                        () => {
                            // Log
                            ErrorDialog.logInfo(`App '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} was successfully synced to Teams...`);

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
                } else {
                    // Log the error
                    ErrorDialog.show("Tenant App", "The app was not found in the tenant app catalog.", DataSource.AppItem);
                }
            }, ex => {
                // Log the error
                ErrorDialog.show("Tenant App", "Unable to query the tenant app catalog for the item.", ex);
            });
        }, ex => {
            // Log
            ErrorDialog.logInfo(`Error getting the context information for the tenant app catalog...`, ex);
        });
    }

    // Reads an app package file
    static readPackage(data): PromiseLike<{ assets: JSZip.JSZipObject[]; item: IAppItem; image: JSZip.JSZipObject; spfxProd: Uint8Array; spfxTest: Uint8Array; }> {
        // Gets the app package metadata
        let readMetadata = (fileInfo: JSZip.JSZipObject): PromiseLike<IAppItem> => {
            // Return a promise
            return new Promise(resolve => {
                let metadata: IAppItem = {} as any;

                // Ensure the file information exists
                if (fileInfo) {
                    // Log
                    ErrorDialog.logInfo(`${fileInfo.name}: App Manifest found. Processing the information...`);

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

                        // Default the status
                        metadata.AppStatus = "New";

                        // See if the config has a default value for the support url
                        if (AppConfig.Configuration.supportUrl) {
                            metadata.AppSupportURL = {
                                Description: AppConfig.Configuration.supportUrl,
                                Url: AppConfig.Configuration.supportUrl
                            };
                        }

                        // Log
                        ErrorDialog.logInfo(`App Information Extracted: \r\n${JSON.stringify(metadata)}`);

                        // Resolve the request
                        resolve(metadata);
                    });
                } else {
                    // Log
                    ErrorDialog.logError(`The app manifest file was not found in the app package.`);

                    // Resolve the request
                    resolve(null);
                }
            });
        }

        // Generates the CDN package
        let generateCDNPackage = (appId: string, zipFile: JSZip, files: JSZip.JSZipObject[], isTest: boolean = true): PromiseLike<Uint8Array> => {
            // Return a promise
            return new Promise(resolve => {
                // Get the CDN url from the configuration and ensure it exists
                let cdn = isTest ? AppConfig.Configuration.cdnTest : AppConfig.Configuration.cdnProd;
                if (cdn) {
                    // Ensure the url has a trailing '/'
                    cdn = cdn.replace(/\/$/, '') + '/' + appId.toLowerCase() + '/';

                    // Parse the files
                    Helper.Executor(files, file => {
                        // Return a promise
                        return new Promise(resolve => {
                            // Read the file
                            file.async("string").then(content => {
                                // Ensure the default cdn exists
                                if (content.indexOf("HTTPS://SPCLIENTSIDEASSETLIBRARY/") > 0) {
                                    // Replace the content
                                    content = content.replace("HTTPS://SPCLIENTSIDEASSETLIBRARY/", cdn);

                                    // Update the file
                                    zipFile.file(file.name, content);
                                }

                                // Resolve the request
                                resolve(null);
                            });
                        });
                    }).then(() => {
                        // Get the file contents and resolve the request
                        zipFile.generateAsync({ type: "uint8array" }).then(resolve);
                    });
                } else {
                    // Resolve the request
                    resolve(null);
                }
            });
        }

        // Gets the component manifest content
        let getComponentManifestContent = (files: JSZip.JSZipObject[]): PromiseLike<{ content: string; files: JSZip.JSZipObject[] }> => {
            // Return a promise
            return new Promise(resolve => {
                let appContent = null;
                let appFiles = [];

                // Parse the files
                Helper.Executor(files, file => {
                    // Return a promise
                    return new Promise((resolve) => {
                        // Log
                        ErrorDialog.logInfo(`${file.name}: Reading XML Configuration File. Processing the information...`);

                        // Read the file
                        file.async("string").then(content => {
                            var i = content.indexOf("<ClientSideComponent");
                            if (i < 0) { resolve(false); return; }

                            // Read the configuration
                            var oParser = new DOMParser();
                            var oDOM = oParser.parseFromString(content.substring(i), "text/xml");

                            // Get the manifest information
                            let xmlAppManifest = oDOM.documentElement.attributes["ComponentManifest"];
                            if (xmlAppManifest) {
                                // Try to convert it
                                try {
                                    // Get the manifest information
                                    let manifest = JSON.parse(xmlAppManifest.value.replace(/&quot;/g, '"'));

                                    // Set the app manifest value
                                    appContent = JSON.stringify(manifest, null, 2);

                                    // Set the file
                                    appFiles.push(file);

                                    // Log
                                    ErrorDialog.logInfo(`App Manifest: \r\n${manifest.value}`);
                                } catch { }
                            }

                            // Check the next file
                            resolve(null);
                        });
                    });
                }).then(() => {
                    // Resolve the request
                    resolve({ content: appContent, files: appFiles });
                })
            });
        }

        // Return a promise
        return new Promise(resolve => {
            // Log
            ErrorDialog.logInfo(`Unzipping the SPFx package...`);

            // Unzip the package
            JSZip.loadAsync(data).then(zipFile => {
                let assets: JSZip.JSZipObject[] = [];
                let image = null;
                let xmlFiles: JSZip.JSZipObject[] = [];

                // Log
                ErrorDialog.logInfo(`Parsing the SPFx package files...`);

                // Read the app manifest
                readMetadata(zipFile.files["AppManifest.xml"]).then(item => {
                    // Parse the files
                    zipFile.forEach((path, fileInfo) => {
                        let fileName = fileInfo.name.toLowerCase();

                        // See if this is an image
                        if (fileName.endsWith(".png") || fileName.endsWith(".jpg") || fileName.endsWith(".jpeg") || fileName.endsWith(".gif")) {
                            // Log
                            ErrorDialog.logInfo(`${fileName}: Image file found...`);

                            // Save a reference to the iamge
                            image = fileInfo;

                            // Check the next file
                            return;
                        }

                        // See if this is a client side asset
                        if (path.toLowerCase().indexOf("clientsideassets") == 0) {
                            // Ensure this is a file
                            if (!fileInfo.name.endsWith('/')) {
                                // Add the asset information
                                assets.push(fileInfo);
                            }
                        }

                        // See if this is a component xml file
                        if (fileName.indexOf("/") > 0 && fileName.endsWith(".xml")) {
                            // Save a reference to this file
                            xmlFiles.push(fileInfo);
                        }
                    });

                    // Determine if this app supports teams
                    getComponentManifestContent(xmlFiles).then(info => {
                        // Set the app manifest value
                        item.AppManifest = info.content;

                        // Generate the test package
                        generateCDNPackage(item.AppProductID, zipFile, info.files, true).then(spfxTest => {
                            // Generate the production package
                            generateCDNPackage(item.AppProductID, zipFile, info.files, false).then(spfxProd => {
                                // Resolve the request
                                resolve({ assets, item, image, spfxProd, spfxTest });
                            });
                        });
                    });
                });
            });
        });
    }

    // Retracts the solution from the app catalog
    static retract(tenantFl: boolean, removeFl: boolean, onUpdate: () => void) {
        // Log
        ErrorDialog.logInfo(`Retracting the SPFx package...`);

        // See if this is for the tenant
        if (tenantFl) {
            // Ensure the app exists
            if (!AppSecurity.IsTenantAppCatalogOwner || DataSource.AppCatalogTenantItem == null) {
                // Call the update method and return
                onUpdate();
                return;
            }
        } else {
            // Ensure the app exists
            if (!AppSecurity.IsSiteAppCatalogOwner || DataSource.AppCatalogSiteItem == null) {
                // Call the update method and return
                onUpdate();
                return;
            }
        }

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
                ErrorDialog.logInfo(`Retracting the app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} from the app catalog...`);

                // Load the apps
                let web = Web(catalogUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue });
                (tenantFl ? web.TenantAppCatalog() : web.SiteCollectionAppCatalog()).AvailableApps(DataSource.AppItem.AppProductID).retract().execute(() => {
                    // Log
                    ErrorDialog.logInfo(`The app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} was retracted from the app catalog successfully...`);

                    // See if this is a tenant app
                    if (tenantFl) {
                        // Update the item
                        DataSource.AppItem.update({
                            AppIsTenantDeployed: false
                        }).execute(() => {
                            // Log
                            ErrorDialog.logInfo(`The app metadata flag for tenant deployment was cleared.`);
                        }, () => {
                            // Log
                            ErrorDialog.logError(`Failed to clear the app metadata flag for tenant deployment.`);
                        });
                    }

                    // See if we are removing the app
                    if (removeFl) {
                        // Log
                        ErrorDialog.logInfo(`Removing the SPFx app from the catalog...`);

                        // Remove the app
                        (tenantFl ? web.TenantAppCatalog() : web.SiteCollectionAppCatalog()).AvailableApps(DataSource.AppItem.AppProductID).remove().execute(
                            () => {
                                // Log
                                ErrorDialog.logInfo(`The app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} was removed from the app catalog successfully...`);

                                // Close the dialog
                                LoadingDialog.hide();

                                // Call the update event
                                onUpdate();
                            },
                            ex => {
                                // Log
                                ErrorDialog.logInfo(`There was an error removing the app '${DataSource.AppItem.Title}'.`);

                                // Close the dialog
                                LoadingDialog.hide();

                                // Call the update event
                                onUpdate();
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
                    ErrorDialog.logInfo(`Error retracting the app '${DataSource.AppItem.Title} with id ${DataSource.AppItem.AppProductID} from the app catalog...`);

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

    // Runs the flow for the item
    static runFlow(flowId: string) {
        // See if a flow exists and the user is licensed to run it
        if (flowId && DataSource.HasLicense) {
            // Run the flow
            List.runFlow({
                data: {
                    rows: [{
                        entity: {
                            ID: DataSource.AppItem.Id
                        }
                    }]
                },
                id: flowId,
                list: Strings.Lists.Apps,
                cloudEnv: AppConfig.Configuration.cloudEnv,
                webUrl: Strings.SourceUrl
            }).then(results => {
                // See if it didn't run
                if (!results.executed) {
                    // Log the error
                    ErrorDialog.logError(`Flow '${flowId} failed: ${results.errorMessage}`);
                }
            });
        }
    }

    // Uninstalls the app
    // This will delete the app and not place it in the recycle bin
    private static uninstallApp(siteUrl: string, requestDigest: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Update the dialog
            LoadingDialog.setHeader("Uninstalling the Solution");
            LoadingDialog.setBody("Uninstalling in site: " + siteUrl + "<br/>This will close after the app is upgraded.");
            LoadingDialog.show();

            // Log
            ErrorDialog.logInfo(`Uninstalling the app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} from the app catalog...`);

            // Set the complete method
            let onComplete = () => {
                // Close the dialog
                LoadingDialog.hide();

                // Resolve the request
                resolve();
            }

            // Uninstall the app
            Web(siteUrl, { requestDigest }).SiteCollectionAppCatalog().AvailableApps(DataSource.AppItem.AppProductID).uninstall().execute(onComplete, onComplete);
        });
    }

    // Updates the app in a site collection app catalog
    static updateSiteApp(siteUrl: string, isTestSite: boolean, onError: () => void): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // Log
            ErrorDialog.logInfo(`Getting the web context information at ${siteUrl}...`);

            // Show a loading dialog
            LoadingDialog.setHeader("Getting the web information");
            LoadingDialog.setBody("Loading the web information...");
            LoadingDialog.show();

            // Load the context of the app catalog
            ContextInfo.getWeb(siteUrl).execute(
                context => {
                    let requestDigest = context.GetContextWebInformation.FormDigestValue;

                    // Uninstall the app
                    this.uninstallApp(siteUrl, requestDigest).then(() => {
                        // Update the dialog
                        LoadingDialog.setHeader("Upgrading the Solution");
                        LoadingDialog.setBody("Deploying the app to the catalog...");
                        LoadingDialog.show();

                        // Log
                        ErrorDialog.logInfo(`Upgrading the app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID}...`);

                        // Deploy the solution
                        this.deploy("test", false, isTestSite ? false : DataSource.AppItem.AppSkipFeatureDeployment, onError, () => {
                            // Update the dialog
                            LoadingDialog.setHeader("Reading the App Catalog");
                            LoadingDialog.setBody("Getting the new app from the app catalog...");
                            LoadingDialog.show();

                            // Log
                            ErrorDialog.logInfo(`The app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} was upgraded successfully...`);

                            // Log
                            ErrorDialog.logInfo(`Getting the app from the catalog...`);

                            // Get the app
                            Web(siteUrl, { requestDigest }).SiteCollectionAppCatalog().AvailableApps(DataSource.AppItem.AppProductID).execute(
                                app => {
                                    // Log
                                    ErrorDialog.logInfo(`The app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} was found in the app catalog...`);

                                    // See if the app is already installed
                                    if (app.SkipDeploymentFeature) {
                                        // Hide the dialog
                                        LoadingDialog.hide();

                                        // Resolve the requst
                                        resolve();
                                    } else {
                                        // Log
                                        ErrorDialog.logInfo(`Installing the app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID}...`);

                                        // Update the dialog
                                        LoadingDialog.setHeader("Installing the Solution");
                                        LoadingDialog.setBody("Installing in site: " + siteUrl + "<br/>This will close after the app is upgraded.");
                                        LoadingDialog.show();

                                        // Install the app
                                        Web(siteUrl, { requestDigest }).SiteCollectionAppCatalog().AvailableApps(DataSource.AppItem.AppProductID).install().execute(
                                            () => {
                                                // Log
                                                ErrorDialog.logInfo(`The app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} was installed successfully...`);

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
                    });
                },
                ex => {
                    // Log the error
                    ErrorDialog.show("Getting Context", "There was an error getting the web context", ex);
                }
            );
        });
    }

    // Updates the app in a tenant app catalog
    static updateTenantApp(skipFeatureDeployment: boolean, onError: () => void): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // Log
            ErrorDialog.logInfo(`Getting the web context information at ${AppConfig.Configuration.tenantAppCatalogUrl}...`);

            // Show a loading dialog
            LoadingDialog.setHeader("Getting the web information");
            LoadingDialog.setBody("Loading the web information...");
            LoadingDialog.show();

            // Load the context of the app catalog
            ContextInfo.getWeb(AppConfig.Configuration.tenantAppCatalogUrl).execute(
                context => {
                    let requestDigest = context.GetContextWebInformation.FormDigestValue;

                    // Uninstall the app
                    this.uninstallApp(AppConfig.Configuration.tenantAppCatalogUrl, requestDigest).then(() => {
                        // Update the dialog
                        LoadingDialog.setHeader("Upgrading the Solution");
                        LoadingDialog.setBody("Deploying the app to the catalog...");
                        LoadingDialog.show();

                        // Log
                        ErrorDialog.logInfo(`Upgrading the app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID}...`);

                        // Deploy the solution
                        this.deploy("", true, skipFeatureDeployment, onError, () => {
                            // Update the dialog
                            LoadingDialog.setHeader("Reading the App Catalog");
                            LoadingDialog.setBody("Getting the new app from the app catalog...");
                            LoadingDialog.show();

                            // Log
                            ErrorDialog.logInfo(`The app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} was upgraded successfully...`);

                            // Log
                            ErrorDialog.logInfo(`Getting the app from the catalog...`);

                            // Get the app
                            Web(AppConfig.Configuration.tenantAppCatalogUrl, { requestDigest }).TenantAppCatalog().AvailableApps(DataSource.AppItem.AppProductID).execute(
                                app => {
                                    // Log
                                    ErrorDialog.logInfo(`The app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} was found in the app catalog...`);

                                    // See if the app is already installed
                                    if (app.SkipDeploymentFeature) {
                                        // Hide the dialog
                                        LoadingDialog.hide();

                                        // Resolve the requst
                                        resolve();
                                    } else {
                                        // Log
                                        ErrorDialog.logInfo(`Installing the app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID}...`);

                                        // Update the dialog
                                        LoadingDialog.setHeader("Installing the Solution");
                                        LoadingDialog.setBody("Installing in site: " + AppConfig.Configuration.tenantAppCatalogUrl + "<br/>This will close after the app is upgraded.");
                                        LoadingDialog.show();

                                        // Install the app
                                        Web(AppConfig.Configuration.tenantAppCatalogUrl, { requestDigest }).TenantAppCatalog().AvailableApps(DataSource.AppItem.AppProductID).install().execute(
                                            () => {
                                                // Log
                                                ErrorDialog.logInfo(`The app '${DataSource.AppItem.Title}' with id ${DataSource.AppItem.AppProductID} was installed successfully...`);

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
                    });
                },
                ex => {
                    // Log the error
                    ErrorDialog.show("Getting Context", "There was an error getting the web context", ex);
                }
            );
        });
    }

    // Updates the app metadata
    static updateAppMetadata(appItemId: number, tenantFl: boolean, catalogUrl: string, requestDigest: string): PromiseLike<void> {
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
                        // Set the metadata
                        let metadata = {
                            AppDescription: DataSource.AppItem.AppDescription,
                            AppImageURL1: DataSource.AppItem.AppImageURL1 ? { Description: DataSource.AppItem.AppImageURL1.Description, Url: DataSource.AppItem.AppImageURL1.Url } : null,
                            AppImageURL2: DataSource.AppItem.AppImageURL2 ? { Description: DataSource.AppItem.AppImageURL2.Description, Url: DataSource.AppItem.AppImageURL2.Url } : null,
                            AppImageURL3: DataSource.AppItem.AppImageURL3 ? { Description: DataSource.AppItem.AppImageURL3.Description, Url: DataSource.AppItem.AppImageURL3.Url } : null,
                            AppImageURL4: DataSource.AppItem.AppImageURL4 ? { Description: DataSource.AppItem.AppImageURL4.Description, Url: DataSource.AppItem.AppImageURL4.Url } : null,
                            AppImageURL5: DataSource.AppItem.AppImageURL5 ? { Description: DataSource.AppItem.AppImageURL5.Description, Url: DataSource.AppItem.AppImageURL5.Url } : null,
                            AppPublisher: DataSource.AppItem.AppPublisher,
                            AppShortDescription: DataSource.AppItem.AppShortDescription,
                            AppSupportURL: DataSource.AppItem.AppSupportURL ? { Description: DataSource.AppItem.AppSupportURL.Description, Url: DataSource.AppItem.AppSupportURL.Url } : null,
                            AppThumbnailURL: DataSource.AppItem.AppThumbnailURL ? { Description: DataSource.AppItem.AppThumbnailURL.Description, Url: DataSource.AppItem.AppThumbnailURL.Url } : null,
                            AppVideoURL: DataSource.AppItem.AppVideoURL ? { Description: DataSource.AppItem.AppVideoURL.Description, Url: DataSource.AppItem.AppVideoURL.Url } : null,
                        };

                        // See if there is a CDN for the images
                        if (AppConfig.Configuration.cdnImage) {
                            // Parse the fields
                            Helper.Executor(["AppThumbnailURL", "AppImageURL1", "AppImageURL2", "AppImageURL3", "AppImageURL4", "AppImageURL5"], fieldName => {
                                // Ensure a value exists
                                let fieldValue = DataSource.AppItem[fieldName] ? DataSource.AppItem[fieldName].Url : null;
                                if (fieldValue) {
                                    // Get the file name
                                    let urlInfo = fieldValue.split('/');
                                    let fileName = urlInfo[urlInfo.length - 1];

                                    // Set the url
                                    let imageUrl = `${AppConfig.Configuration.cdnImage}/${DataSource.AppItem.AppProductID}/${fileName}`;
                                    metadata[fieldName] = { Description: imageUrl, Url: imageUrl };

                                    // Log
                                    ErrorDialog.logInfo(`Setting the ${fieldName} to: ${fieldValue}`);
                                }
                            });
                        }

                        // Log
                        ErrorDialog.logInfo("Setting the metadata in the app catalog", metadata);

                        // Update the metadata
                        list.Items(appItemId).update(metadata).execute(
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
        Helper.ListForm.showFileDialog([".sppkg"]).then(file => {
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
                    for (let i = 0; i < DataSource.DocSetList.Items.length; i++) {
                        let item = DataSource.DocSetList.Items[i];

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
                        LoadingDialog.show();

                        // Create the document set folder
                        Helper.createDocSet(pkgInfo.item.Title, Strings.Lists.Apps).then(
                            // Success
                            (item: IAppItem) => {
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

                                    // Set the id
                                    pkgInfo.item.Id = item.Id;

                                    // Update the loading dialog
                                    LoadingDialog.setHeader("Uploading the Package");
                                    LoadingDialog.setBody("Uploading the app package...");

                                    // Log
                                    ErrorDialog.logInfo(`Uploading the SPFx package(s) ${file.name} to the folder...`);

                                    // Upload the packages
                                    this.uploadSPFxPackages(item.Folder(), file, pkgInfo.spfxTest, pkgInfo.spfxProd).then(() => {
                                        // Log
                                        ErrorDialog.logInfo(`The SPFx package was uploaded successfully...`);

                                        // See if the image exists
                                        if (pkgInfo.image) {
                                            // Update the loading dialog
                                            LoadingDialog.setHeader("Uploading the App Icon");
                                            LoadingDialog.setBody("Uploading the icon for this app'...");
                                            LoadingDialog.show();

                                            // Log
                                            ErrorDialog.logInfo(`Uploading the SPFx package icon to the folder...`);

                                            // Get image in different format for later uploading
                                            pkgInfo.image.async("arraybuffer").then(content => {
                                                // Get the file extension
                                                let fileExt: string | string[] = pkgInfo.image.name.split('.');
                                                fileExt = fileExt[fileExt.length - 1];

                                                // Upload the image to this folder
                                                item.Folder().Files().add("AppIcon." + fileExt, true, content).execute(
                                                    file => {
                                                        // Log
                                                        ErrorDialog.logInfo(`The SPFx app icon '${file.Name}' was added successfully...`);

                                                        // Convert the file to base64
                                                        let reader = new FileReader();
                                                        reader.onloadend = () => {
                                                            // Set the icon values
                                                            item.update({
                                                                AppThumbnailURLBase64: reader.result as string,
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
                                                                onComplete(pkgInfo.item);
                                                            });
                                                        }

                                                        // Read the array buffer content
                                                        reader.readAsDataURL(new Blob([content]));
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
                                            onComplete(pkgInfo.item);
                                        }
                                    }, (ex) => {
                                        // Log the error
                                        ErrorDialog.show("Uploading File", "There was an error uploading the app package file.", ex);
                                    });
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

    // Method to upload an existing app package
    static uploadAppPackage(fileInfo: Helper.IListFormAttachmentInfo): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Display a loading dialog
            LoadingDialog.setHeader("Uploading New App Package");
            LoadingDialog.setBody("Validating the package...");
            LoadingDialog.show();

            // Read the package
            AppActions.readPackage(fileInfo.data).then(pkgInfo => {
                let errorMessage = null;
                let appUpgraded = false;

                // Deny if the product id doesn't match
                if (pkgInfo.item.AppProductID != DataSource.AppItem.AppProductID) {
                    // Set the error message
                    errorMessage = "The app's product id doesn't match the current one.";
                }
                // Else, deny if the version is less than the current
                else if (pkgInfo.item.AppVersion != DataSource.AppItem.AppVersion) {
                    // Compare the values
                    let appVersion = DataSource.AppItem.AppVersion.split('.');
                    let newAppVersion = pkgInfo.item.AppVersion.split('.');

                    // See if the new version is greater
                    for (let i = 0; i < appVersion.length; i++) {
                        // See if this version is not greater than
                        if (newAppVersion[0] < appVersion[0]) {
                            // Set the error message
                            errorMessage = "The app's version is less than the current one."
                            break;
                        }
                    }

                    // Set the flag
                    appUpgraded = errorMessage ? false : true;
                }

                // See if an error message exists
                if (errorMessage) {
                    // Deny the file upload
                    reject();

                    // Log the error
                    ErrorDialog.show("Uploading Package", "<p>The app package being uploaded has been denied for the following reason:</p><br/>" + errorMessage);

                    // Show the modal
                    Modal.show();
                } else {
                    // Updates the status
                    let updateStatus = (): PromiseLike<IAppItem> => {
                        // Update the loading dialog
                        LoadingDialog.setBody("Updating the Status");

                        // Return a promise
                        return new Promise((resolve, reject) => {
                            let itemInfo = pkgInfo.item;

                            // Set the current status
                            itemInfo.AppStatus = DataSource.AppItem.AppStatus;

                            // See if the status is after the test case status
                            let status = AppConfig.Status[DataSource.AppItem.AppStatus];
                            if (status.stepNumber > AppConfig.Status[AppConfig.TestCasesStatus].stepNumber) {
                                // Revert the status back to the testing status
                                itemInfo.AppStatus = AppConfig.TestCasesStatus;
                            }

                            // Update the app information
                            DataSource.AppItem.update(itemInfo).execute(
                                () => {
                                    // Resolve the request
                                    resolve(itemInfo);
                                },
                                ex => {
                                    // Log the error
                                    ErrorDialog.show("Updating Package", "There was an error updating the package.", ex);

                                    // Reject the request
                                    reject();
                                }
                            );
                        });
                    }

                    // Update the status and set it back to the testing status
                    updateStatus().then(
                        (itemInfo) => {
                            // Log
                            ErrorDialog.logInfo(`Archiving the ${DataSource.AppItem.Title} app.`);

                            // Archive the package
                            this.archivePackage(() => {
                                // Log
                                ErrorDialog.logInfo(`Uploading the ${DataSource.AppItem.Title} app.`);

                                // Display a loading dialog
                                LoadingDialog.setHeader("Uploading New App Package");
                                LoadingDialog.setBody("Uploading the new package file...");
                                LoadingDialog.show();

                                // Upload the packages
                                this.uploadSPFxPackages(Web(Strings.SourceUrl).getFolderByServerRelativeUrl(DataSource.AppFolder.ServerRelativeUrl), fileInfo, pkgInfo.spfxTest, pkgInfo.spfxProd).then(() => {
                                    // Display a loading dialog
                                    LoadingDialog.setHeader("Refreshing the App");
                                    LoadingDialog.setBody("Refreshing the app details...");
                                    LoadingDialog.show();

                                    // See if the app was upgraded
                                    if (appUpgraded) {
                                        // See if there is a flow
                                        if (AppConfig.Configuration.appFlows && AppConfig.Configuration.appFlows.upgradeApp) {
                                            // Execute the flow
                                            AppActions.runFlow(AppConfig.Configuration.appFlows.upgradeApp);
                                        }

                                        // Log
                                        DataSource.logItem({
                                            LogUserId: ContextInfo.userId,
                                            ParentId: itemInfo.AppProductID || DataSource.AppItem.AppProductID,
                                            ParentListName: Strings.Lists.Apps,
                                            Title: DataSource.AuditLogStates.AppUpdated,
                                            LogComment: `A new version (${itemInfo.AppVersion}) of the app ${itemInfo.Title} was added.`
                                        }, Object.assign({ ...DataSource.AppItem, ...itemInfo }));

                                        // Resolve the request
                                        resolve();
                                    } else {
                                        // Resolve the request
                                        resolve();
                                    }
                                }, ex => {
                                    // Log the error
                                    ErrorDialog.show("Uploading Package", "There was an error uploading the new package.", ex);
                                });
                            });
                        },
                        ex => {
                            // Log the error
                            ErrorDialog.show("Updating Package", "There was an error updating the package.", ex);
                        }
                    );
                }
            });
        });
    }

    // Uploads the spfx client side assets
    static uploadClientSideAssets(pkgType: "" | "test" | "prod"): PromiseLike<void> {
        // Gets the client side assets
        let getClientSideAssets = (file: Types.SP.File): PromiseLike<JSZip.JSZipObject[]> => {
            let assets: JSZip.JSZipObject[] = [];

            // Return a promise
            return new Promise((resolve, reject) => {
                // Ensure the file exists
                if (file == null) { resolve(assets); return; }

                // Read the package contents
                Web(Strings.SourceUrl).getFileByServerRelativeUrl(file.ServerRelativeUrl).content().execute(
                    content => {
                        // Extract the files
                        JSZip.loadAsync(content).then(zipFile => {
                            // Parse the files
                            zipFile.forEach((path, fileInfo) => {
                                // See if this is a client side asset
                                if (path.toLowerCase().indexOf("clientsideassets") == 0) {
                                    // Ensure this is a file
                                    if (!fileInfo.name.endsWith('/')) {
                                        // Add the asset information
                                        assets.push(fileInfo);
                                    }
                                }
                            });

                            // Resolve the request
                            resolve(assets);
                        });
                    },
                    () => {
                        // Error getting the image
                        ErrorDialog.logError("Error getting the app image file: " + DataSource.AppItem.AppThumbnailURL.Url);

                        // Reject the request
                        reject();
                    }
                );
            });
        }

        // Gets the spfx package
        let getSPFxPackage = () => {
            // Parse the Files in the app folder
            for (let i = 0; i < DataSource.AppFolder.Files.results.length; i++) {
                let file = DataSource.AppFolder.Files.results[i];
                let fileName = file.Name.toLowerCase();
                let fileInfo = fileName.split('.');
                let fileExt = fileInfo[fileInfo.length - 1];

                // See if this is a .pkg file
                if (fileExt == "sppkg") {
                    // See if we are looking for the
                    if (fileName.indexOf("-" + pkgType) > 0) {
                        // Return the package
                        return file;
                    }
                }
            }

            // Not found
            return null;
        }

        // Return a promise
        return new Promise((resolve) => {
            let isTest = pkgType == "test";

            // Get the cdn url
            let cdn = isTest ? AppConfig.Configuration.cdnTest : AppConfig.Configuration.cdnProd;
            if (cdn) {
                // Show a loading dialog
                LoadingDialog.setHeader("Uploading Client Side Assets");
                LoadingDialog.setBody("Getting the web information...");
                LoadingDialog.show();

                // Get the package
                let pkgFile = getSPFxPackage();
                if (pkgFile == null) {
                    // Do nothing
                    resolve();
                    return;
                }

                // Get the web and list url information
                let webUrl = "";
                let urlInfo = cdn.split('/');
                let listUrl = urlInfo[urlInfo.length - 1];
                for (let i = 0; i < urlInfo.length - 1; i++) {
                    // Append the url info
                    webUrl += (i == 0 ? "" : "/") + urlInfo[i];
                }

                // Get the web context information
                ContextInfo.getWeb(webUrl).execute(context => {
                    // Set the web
                    let web = Web(webUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue });

                    // Update the loading dialog
                    LoadingDialog.setBody("Getting the client side assets...");

                    // Get the assets
                    getClientSideAssets(pkgFile).then(
                        assets => {
                            // Update the loading dialog
                            LoadingDialog.setBody("Creating/Getting the client side assets folder...");

                            // Ensure the folder is created
                            this.createClientSideAssetsFolder(web.Folders(listUrl)).then(folder => {
                                // Update the loading dialog
                                LoadingDialog.setBody("Uploading the solution assets to the app folder...");

                                // Parse the assets
                                Helper.Executor(assets, assetFile => {
                                    // Return a promise
                                    return new Promise((resolve) => {
                                        // Get the file information
                                        assetFile.async("arraybuffer").then(content => {
                                            // Upload the file
                                            folder.Files().add(assetFile.name.replace(/clientsideassets\//i, ''), true, content).execute(resolve, () => {
                                                // Error uploading the asset
                                                ErrorDialog.logError("Error uploading the client side asset file: " + assetFile.name);

                                                // Upload the next file
                                                resolve(null);
                                            });
                                        });
                                    });
                                }).then(() => {
                                    // Close the dialog and resolve the request
                                    LoadingDialog.hide();
                                    resolve();
                                });
                            }, resolve);
                        },
                        () => {
                            // Unable to get the assets
                            // Skip this and they will need to do this manually
                            resolve();
                        }
                    );
                });
            } else {
                // Resolve the request
                resolve();
            }
        });
    }

    // Uploads the app images
    static uploadImages(pkgType: string): PromiseLike<void> {
        // Get the app images
        let getImages = (): PromiseLike<{ file: Types.SP.File, content: any }[]> => {
            let files: { file: Types.SP.File, content: any }[] = [];

            // Return a promise
            return new Promise((resolve) => {
                // Parse the fields
                Helper.Executor(["AppThumbnailURL", "AppImageURL1", "AppImageURL2", "AppImageURL3", "AppImageURL4", "AppImageURL5"], fieldName => {
                    // Return a promise
                    return new Promise((resolve) => {
                        // Ensure the url exists
                        if (DataSource.AppItem[fieldName] && DataSource.AppItem[fieldName].Url) {
                            // Read the package contents
                            Web(Strings.SourceUrl).getFileByUrl(DataSource.AppItem[fieldName].Url).execute(
                                file => {
                                    // Get the content
                                    file.content().execute(
                                        content => {
                                            // Append the file
                                            files.push({ file, content });

                                            // Check the next file
                                            resolve(null);
                                        },
                                        () => {
                                            // Error getting the image
                                            ErrorDialog.logError("Error downloading the app image file: " + DataSource.AppItem.AppThumbnailURL.Url);

                                            // Check the next file
                                            resolve(null);
                                        }
                                    )
                                },
                                () => {
                                    // Error getting the image
                                    ErrorDialog.logError("Error getting the app image file: " + DataSource.AppItem.AppThumbnailURL.Url);

                                    // Check the next file
                                    resolve(null);
                                }
                            );
                        } else {
                            // Error getting the icon
                            ErrorDialog.logInfo("The app image file doesn't exist for field: " + fieldName);

                            // Check the next file
                            resolve(null);
                        }
                    });
                }).then(() => {
                    // Resolve the request
                    resolve(files);
                });
            });
        }

        // Return a promise
        return new Promise((resolve) => {
            // See if this is the test site
            if (pkgType == "test") {
                // Do nothing
                resolve();
                return;
            }

            // Ensure the url exists
            if (AppConfig.Configuration.cdnImage) {
                // Show a loading dialog
                LoadingDialog.setHeader("Uploading App Images");
                LoadingDialog.setBody("Getting the web information...");
                LoadingDialog.show();

                // Get the web and list url information
                let webUrl = "";
                let urlInfo = AppConfig.Configuration.cdnImage.split('/');
                let listUrl = urlInfo[urlInfo.length - 1];
                for (let i = 0; i < urlInfo.length - 1; i++) {
                    // Append the url info
                    webUrl += (i == 0 ? "" : "/") + urlInfo[i];
                }

                // Get the web context information
                ContextInfo.getWeb(webUrl).execute(context => {
                    // Set the web
                    let web = Web(webUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue });

                    // Update the loading dialog
                    LoadingDialog.setBody("Getting the app images...");

                    // Get the images
                    getImages().then(appImages => {
                        // Update the loading dialog
                        LoadingDialog.setBody("Creating/Getting the image folder...");

                        // Ensure the folder is created
                        this.createImageFolder(web.Folders(listUrl)).then(folder => {
                            // Update the loading dialog
                            LoadingDialog.setBody("Uploading the images to the app folder...");

                            // Parse the files
                            Helper.Executor(appImages, appImage => {
                                // Return a promise
                                return new Promise((resolve) => {
                                    // Update the loading dialog
                                    LoadingDialog.setBody("Uploading the image: " + appImage.file.Name);

                                    // Upload the file
                                    folder.Files().add(appImage.file.Name, true, appImage.content).execute(() => {
                                        // Close the dialog and resolve the request
                                        LoadingDialog.hide();

                                        // Upload the next file
                                        resolve(null);
                                    }, () => {
                                        // Error uploading the asset
                                        ErrorDialog.logError("Error uploading the app image file: " + appImage.file.Name);

                                        // Upload the next file
                                        resolve(null);
                                    });
                                });
                            }).then(() => { resolve(); });
                        }, () => { resolve(); });
                    }, () => { resolve(); });
                }, () => { resolve(); });
            } else {
                // Resolve the request
                resolve();
            }
        });
    }

    // Uploads the spfx packages to the folder
    private static uploadSPFxPackages(appFolder: Types.SP.IFolder, spfxPkg: Helper.IListFormAttachmentInfo, testPkg: Uint8Array, prodPkg: Uint8Array): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Update the loading dialog
            LoadingDialog.setHeader("Uploading the SPFx Packages");
            LoadingDialog.setBody("Uploading the test and prod generated packages...");
            LoadingDialog.show();

            // Log
            ErrorDialog.logInfo(`Uploading the SPFx package assets...`);

            // Get the file name without extension
            let pkgName = spfxPkg.name.replace(/.sppkg/i, '');

            // Upload the packages
            appFolder.Files().add(spfxPkg.name, true, spfxPkg.data).execute(() => {
                Helper.Executor([{ name: pkgName + "-test.sppkg", content: testPkg }, { name: pkgName + "-prod.sppkg", content: prodPkg }], pkg => {
                    // Return a promise
                    return new Promise(resolve => {
                        // See if the file contents exists
                        if (pkg.content) {
                            // Upload the file and process the next file afterwards
                            appFolder.Files().add(pkg.name, true, pkg.content).execute(resolve, reject);
                        } else {
                            // Process the next file
                            resolve(null);
                        }
                    });
                }).then(() => {
                    // Close the dialog and resolve the request
                    LoadingDialog.hide();
                    resolve();
                });
            }, reject);
        });
    }
}