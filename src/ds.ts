import { InstallationRequired, LoadingDialog } from "dattatable";
import { Components, ContextInfo, Helper, Site, Types, Web } from "gd-sprest-bs";
import { AppConfig } from "./appCfg";
import { AppSecurity } from "./appSecurity";
import { Configuration, createSecurityGroups } from "./cfg";
import * as Common from "./common";
import Strings from "./strings";

// App Item
export interface IAppItem extends Types.SP.ListItem {
    AppAPIPermissions?: string;
    AppComments?: string;
    AppDescription: string;
    AppDevelopers: { results: { Id: number; EMail: string; Title: string; }[] };
    AppDevelopersId: { results: number[] };
    AppIsClientSideSolution?: boolean;
    AppIsDomainIsolated?: boolean;
    AppIsRejected: boolean;
    AppIsTenantDeployed: boolean;
    AppJustification: string;
    AppPermissionsJustification: string;
    AppProductID: string;
    AppPublisher: string;
    AppSharePointMinVersion?: boolean;
    AppSiteDeployments?: string;
    AppSkipFeatureDeployment?: boolean;
    AppSponsor: { Id: number; EMail: string; Title: string; };
    AppSponsorId: number;
    AppStatus: string;
    AppVersion: string;
    AuthorId: number;
    AppPackageErrorMessage?: string;
    FileLeafRef: string;

    AppImageURL1: Types.SP.FieldUrlValue;
    AppImageURL2: Types.SP.FieldUrlValue;
    AppImageURL3: Types.SP.FieldUrlValue;
    AppImageURL4: Types.SP.FieldUrlValue;
    AppImageURL5: Types.SP.FieldUrlValue;
    AppShortDescription: string;
    AppSupportURL: Types.SP.FieldUrlValue;
    AppThumbnailURL: Types.SP.FieldUrlValue;
    AppVideoURL: Types.SP.FieldUrlValue;
    CheckoutUser: { Id: number; Title: string; };
    ContentTypeId: string;
    IsDefaultAppMetadataLocale: boolean;
    IsAppPackageEnabled: boolean;
}

// Assessment Item
export interface IAssessmentItem extends Types.SP.ListItem {
    Completed?: string;
    RelatedAppId: number;
}

/**
 * Data Source
 */
export class DataSource {
    // Loads the document set item
    private static _docSetInfo: Helper.IListFormResult = null;
    static get DocSetFolder(): Types.SP.FolderOData { return this.DocSetItem ? this.DocSetItem.Folder as any : null; }
    static get DocSetInfo(): Helper.IListFormResult { return this._docSetInfo; }
    static get DocSetItem(): IAppItem { return this.DocSetInfo ? this.DocSetInfo.item as any : null; }
    static get DocSetItemId(): number { return this.DocSetItem ? this.DocSetItem.Id : 0; }
    private static _docSetSCApp: Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata = null;
    static get DocSetSCApp(): Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata { return this._docSetSCApp; }
    private static _docSetSCAppItem: IAppItem = null;
    static get DocSetSCAppItem(): IAppItem { return this._docSetSCAppItem; }
    private static _docSetTenantApp: Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata = null;
    static get DocSetTenantApp(): Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata { return this._docSetTenantApp; }
    static loadDocSetFromQS(): number {
        let itemId = null;

        // Parse the query string values
        let qs = document.location.search.split('?');
        qs = qs.length > 1 ? qs[1].split('&') : [];
        for (let i = 0; i < qs.length; i++) {
            let qsItem = qs[i].split('=');
            let key = qsItem[0].toLowerCase();
            let value = qsItem[1];

            // See if this is the "id" key
            if (key == "id" || key == "app-id") {
                // Return the item
                itemId = parseInt(value);
                break;
            }
        }

        // Return the doc set item id
        return itemId;
    }
    static loadDocSet(id?: number): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the list information
            Components.ListForm.create({
                listName: Strings.Lists.Apps,
                itemId: id || this.DocSetItemId,
                contentType: "App",
                webUrl: Strings.SourceUrl,
                query: {
                    Expand: [
                        "AppDevelopers", "AppSponsor", "CheckoutUser",
                        "Folder/Files", "Folder/Folders"
                    ],
                    Select: [
                        "*", "Id", "FileLeafRef", "ContentTypeId",
                        "AppDevelopers/Id", "AppDevelopers/EMail", "AppDevelopers/Title",
                        "AppSponsor/Id", "AppSponsor/EMail", "AppSponsor/Title",
                        "CheckoutUser/Id", "CheckoutUser/EMail", "CheckoutUser/Title"
                    ]
                }
            }).then(info => {
                // Save the info
                this._docSetInfo = info;

                // Ensure items exist
                if (this._items) {
                    // Parse the items
                    for (let i = 0; i < this._items.length; i++) {
                        let item = this._items[i];

                        // See if the id match
                        if (item.Id == info.item.Id) {
                            // Update the item
                            this._items[i] = info.item as any;
                            break;
                        }
                    }
                }

                // Load the site collection apps
                this.loadSiteCollectionApp(this.DocSetItem.AppProductID).then(() => {
                    // Load the tenant apps
                    this.loadTenantApp(this.DocSetItem.AppProductID).then(() => {
                        // Resolve the request
                        resolve();
                    }, reject);
                }, reject);
            }, reject);
        });
    }

    // Deletes the app test site
    static deleteTestSite(item: IAppItem): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // Get the url to the test site
            let url = [AppConfig.Configuration.appCatalogUrl, item.AppProductID].join('/');

            // Get the web context
            ContextInfo.getWeb(url).execute(context => {
                // Delete the web
                Web(url, { requestDigest: context.GetContextWebInformation.FormDigestValue }).delete().execute(() => {
                    // Resolve the request
                    resolve();
                }, () => {
                    // Web doesn't exist
                    resolve();
                });
            }, () => {
                // Error getting the context
                resolve();
            });
        });
    }

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

    // Loads the app test site
    static loadTestSite(item: IAppItem): PromiseLike<{ app: Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata, web: Types.SP.Web }> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the url to the test site
            let url = [AppConfig.Configuration.appCatalogUrl, item.AppProductID].join('/');

            // Get the web
            Web(url).execute(
                // Success
                web => {
                    // Get the app
                    web.SiteCollectionAppCatalog().AvailableApps(item.AppProductID).execute(
                        // App exists
                        app => {
                            // Resolve the request
                            resolve({ app, web });
                        },
                        // App doesn't exist
                        () => {
                            // Reject the request
                            reject();
                        }
                    );
                },

                // Error
                () => {
                    // Site doesn't exist
                    reject();
                }
            );
        });
    }

    // Initializes the application
    static init(cfgWebUrl?: string, cfgUrl?: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the configuration
            AppConfig.loadConfiguration(cfgWebUrl, cfgUrl).then(() => {
                // Load the security information
                AppSecurity.init(AppConfig.Configuration.appCatalogUrl, AppConfig.Configuration.tenantAppCatalogUrl).then(() => {
                    // Call the refresh method to load the data
                    this.refresh(this.loadDocSetFromQS()).then(resolve, reject);
                }, reject);
            }, reject);
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
                        content: "App configuration doesn't exist. Edit the webpart and set the configuration property.",
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
                if (AppSecurity.ApproverGroup == null || AppSecurity.DevGroup == null || AppSecurity.SponsorGroup == null) {
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

    // Loads the list data
    private static _items: IAppItem[] = null;
    static get Items(): IAppItem[] { return this._items; }
    static load(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the list items
            Web(Strings.SourceUrl).Lists(Strings.Lists.Apps).Items().query({
                Expand: ["AppDevelopers", "AppSponsor", "CheckoutUser", "Folder"],
                Filter: "ContentType eq 'App'",
                Select: [
                    "*", "Id", "FileLeafRef", "ContentTypeId",
                    "AppDevelopers/Id", "AppDevelopers/EMail", "AppDevelopers/Title",
                    "AppSponsor/Id", "AppSponsor/EMail", "AppSponsor/Title",
                    "CheckoutUser/Id", "CheckoutUser/EMail", "CheckoutUser/Title"
                ]
            }).execute(items => {
                // Set the items
                this._items = items.results as any;

                // Resolve the promise
                resolve();
            }, reject);
        });
    }

    // Status Filters
    private static _statusFilters: Components.ICheckboxGroupItem[] = null;
    static get StatusFilters(): Components.ICheckboxGroupItem[] { return this._statusFilters; }
    static loadStatusFilters(): PromiseLike<Components.ICheckboxGroupItem[]> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the status field
            Web(Strings.SourceUrl).Lists(Strings.Lists.Apps).Fields("AppStatus").execute((fld: Types.SP.FieldChoice) => {
                let items: Components.ICheckboxGroupItem[] = [];

                // Parse the choices
                for (let i = 0; i < fld.Choices.results.length; i++) {
                    // Add an item
                    items.push({
                        label: fld.Choices.results[i],
                        type: Components.CheckboxGroupTypes.Switch
                    });
                }

                // Set the filters and resolve the promise
                this._statusFilters = items;
                resolve(items);
            }, reject);
        });
    }

    // Site Collection Apps
    private static _siteCollectionApps: Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata[] = null;
    static get SiteCollectionAppCatalogExists(): boolean { return this._siteCollectionApps != null; }
    static get SiteCollectionApps(): Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata[] { return this._siteCollectionApps; }
    static getSiteCollectionAppById(appId: string): Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata {
        // Parse the apps
        for (let i = 0; i < this._siteCollectionApps.length; i++) {
            let app = this._siteCollectionApps[i];

            // See if this is the target app
            if (app.ProductId == appId) { return app; }
        }

        // App not found
        return null;
    }
    static loadSiteCollectionApp(id: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // See if the app catalog is defined
            if (AppConfig.Configuration.appCatalogUrl) {
                // Load the available apps
                Web(AppConfig.Configuration.appCatalogUrl).SiteCollectionAppCatalog().AvailableApps(id).execute(app => {
                    // Set the app
                    this._docSetSCApp = app;

                    // Resolve the request
                    resolve();
                }, () => {
                    // Query the app catalog and try to find it by file name
                    Web(AppConfig.Configuration.appCatalogUrl).Lists("Apps for SharePoint").Items().query({
                        Expand: ["File"]
                    }).execute(
                        // Success
                        items => {
                            // Parse the docset items
                            for (let i = 0; i < this.DocSetFolder.Files.results.length; i++) {
                                let file = this.DocSetFolder.Files.results[i];

                                // Ensure this is the package file
                                if (!file.Name.endsWith(".sppkg")) { continue; }

                                // Parse the items
                                for (let j = 0; j < items.results.length; j++) {
                                    let item = items.results[i];

                                    // See if the item matches
                                    if (item.File.Name == file.Name) {
                                        // Item found
                                        this._docSetSCAppItem = item as any;
                                        break;
                                    }
                                }

                                // Break from the loop
                                break;
                            }

                            // Resolve the request
                            resolve();
                        },
                        // Error
                        () => {
                            // App not found, resolve the request
                            resolve();
                        }
                    );
                    // App not found, resolve the request
                    resolve();
                });
            }
        });
    }
    static loadSiteCollectionApps(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // See if the app catalog is defined
            if (AppConfig.Configuration.appCatalogUrl) {
                // Load the available apps
                Web(AppConfig.Configuration.appCatalogUrl).SiteCollectionAppCatalog().AvailableApps().execute(apps => {
                    // Set the apps
                    this._siteCollectionApps = apps.results;

                    // Resolve the request
                    resolve();
                }, () => {
                    // No access to the site collection app catalog
                    this._siteCollectionApps = [];

                    // Resolve the request
                    resolve();
                });
            } else {
                // Default the site collection apps
                this._siteCollectionApps = [];

                // Resolve the request
                resolve();
            }
        });
    }

    // Tenant Apps
    private static _tenantApps: Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata[] = null;
    static get TenantApps(): Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata[] { return this._tenantApps; }
    static getTenantAppById(appId: string): Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata {
        // Parse the apps
        for (let i = 0; i < this._tenantApps.length; i++) {
            let app = this._tenantApps[i];

            // See if this is the target app
            if (app.ProductId == appId) { return app; }
        }

        // App not found
        return null;
    }
    static loadTenantApp(id: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // See if the app catalog is defined
            if (AppConfig.Configuration.tenantAppCatalogUrl) {
                // Load the available apps
                Web(AppConfig.Configuration.tenantAppCatalogUrl).TenantAppCatalog().AvailableApps(id).execute(app => {
                    // Set the app
                    this._docSetTenantApp = app;

                    // Resolve the request
                    resolve();
                }, () => {
                    // App not found, resolve the request
                    resolve();
                });
            }
        });
    }
    static loadTenantApps(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if the tenant app catalog is defined
            if (AppConfig.Configuration.tenantAppCatalogUrl) {
                // Load the available apps
                Web(AppConfig.Configuration.tenantAppCatalogUrl).TenantAppCatalog().AvailableApps().execute(apps => {
                    // Set the apps
                    this._tenantApps = apps.results;

                    // Resolve the request
                    resolve();
                }, () => {
                    // No access to the tenant app catalog
                    this._tenantApps = [];

                    // Resolve the request
                    resolve();
                });
            } else {
                // Default the tenant apps
                this._tenantApps = [];

                // Resolve the request
                resolve();
            }
        });
    }

    // Method to refresh the data source
    static refresh(docSetId?: number): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the site collection apps
            this.loadSiteCollectionApps().then(() => {
                // Load the tenant apps
                this.loadTenantApps().then(() => {
                    // See if this is a document set id is set
                    if (docSetId > 0) {
                        // Load the document set item
                        this.loadDocSet(docSetId).then(() => {
                            // Resolve the request
                            resolve();
                        }, reject);
                    } else {
                        // Load all the items
                        this.load().then(() => {
                            // Load the status filters
                            this.loadStatusFilters().then(() => {
                                // Resolve the request
                                resolve();
                            }, reject);
                        }, reject);
                    }
                }, reject);
            }, reject);
        });
    }
}