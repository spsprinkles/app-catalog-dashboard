import { ItemForm, LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo, Helper, List, SPTypes, Types, Utility, Web } from "gd-sprest-bs";
import { loadAsync } from "jszip";
import { AppConfig, IStatus } from "./appCfg";
import { DataSource, IAppItem, IAssessmentItem } from "./ds";
import Strings from "./strings";


/**
 * Application Actions
 * The actions taken for an app.
 */
export class AppActions {
    // Creates the test site for the application
    static createTestSite(item: IAppItem, onComplete: () => void) {
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
            ContextInfo.getWeb(AppConfig.Configuration.appCatalogUrl).execute(context => {
                let requestDigest = context.GetContextWebInformation.FormDigestValue;

                // Create the test site
                Web(AppConfig.Configuration.appCatalogUrl, { requestDigest }).WebInfos().add({
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

                                        // Notify the parent this process is complete
                                        onComplete();
                                    });
                                } else {
                                    // Close the dialog
                                    LoadingDialog.hide();

                                    // Notify the parent this process is complete
                                    onComplete();
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

    // Method to delete the test site
    static deleteTestSite(item: IAppItem): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
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
    static deploy(item: IAppItem, tenantFl: boolean, onUpdate: () => void) {
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
                let catalogUrl = tenantFl ? AppConfig.Configuration.tenantAppCatalogUrl : AppConfig.Configuration.appCatalogUrl;

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
    static deployToTeams(item: IAppItem, onComplete: () => void) {
        // Show a loading dialog
        LoadingDialog.setHeader("Deploying to Teams");
        LoadingDialog.setBody("Syncing the app to Teams.");
        LoadingDialog.show();

        // Load the context of the app catalog
        ContextInfo.getWeb(AppConfig.Configuration.tenantAppCatalogUrl).execute(context => {
            let requestDigest = context.GetContextWebInformation.FormDigestValue;

            // Sync the app with Teams
            Web(AppConfig.Configuration.tenantAppCatalogUrl, { requestDigest }).TenantAppCatalog().syncSolutionToTeams(item.Id).execute(
                // Success
                () => {
                    // Hide the dialog
                    LoadingDialog.hide();

                    // Notify the parent this process is complete
                    onComplete();
                },

                // Error
                () => {
                    // Hide the dialog
                    LoadingDialog.hide();

                    // Error deploying the app
                    // TODO - Show an error

                    // Notify the parent this process is complete
                    onComplete();
                }
            );
        });
    }

    // Reads an app package file
    private static readPackage(data): PromiseLike<IAppItem> {
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

    // Updates the app metadata
    static updateApp(item: IAppItem, appItemId: number, tenantFl: boolean, catalogUrl: string, requestDigest: string): PromiseLike<void> {
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
    static upload(onComplete: (item: IAppItem) => void) {
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
                                    // Update the loading dialog
                                    LoadingDialog.setHeader("Uploading the Package");
                                    LoadingDialog.setBody("Uploading the app package...");

                                    // Upload the file
                                    item.Folder().Files().add(file.name, true, file.data).execute(
                                        // Success
                                        file => {
                                            // Close the loading dialog
                                            LoadingDialog.hide();

                                            // Execute the completed event
                                            onComplete(item as any);
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