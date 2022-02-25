import { InstallationRequired, LoadingDialog } from "dattatable";
import { Components, ContextInfo, Helper, List, SPTypes, Types, Web } from "gd-sprest-bs";
import { Configuration, createSecurityGroups } from "./cfg";
import Strings from "./strings";

// App Catalog Item
export interface IAppCatalogItem extends Types.SP.ListItem {
    AppDescription: string;
    AppImageURL1: Types.SP.FieldUrlValue;
    AppImageURL2: Types.SP.FieldUrlValue;
    AppImageURL3: Types.SP.FieldUrlValue;
    AppImageURL4: Types.SP.FieldUrlValue;
    AppImageURL5: Types.SP.FieldUrlValue;
    AppProductID: string;
    AppPublisher: string;
    AppShortDescription: string;
    AppSupportURL: Types.SP.FieldUrlValue;
    AppThumbnailURL: Types.SP.FieldUrlValue;
    AppVersion: string;
    AppVideoURL: Types.SP.FieldUrlValue;
    AuthorId: number;
    CheckoutUser: { Id: number; Title: string; };
    FileLeafRef: string;
    IsDefaultAppMetadataLocale: boolean;
    IsAppPackageEnabled: boolean;
}

// App Item
export interface IAppItem extends Types.SP.ListItem {
    AppAPIPermissions?: string;
    AppComments?: string;
    AppDescription: string;
    AppDevelopers: { results: { Id: number; EMail: string; }[] };
    AppDevelopersId: { results: number[] };
    AppIsClientSideSolution?: boolean;
    AppIsDomainIsolated?: boolean;
    AppJustification: string;
    AppPermissionsJustification: string;
    AppProductID: string;
    AppPublisher: string;
    AppSharePointMinVersion?: boolean;
    AppSkipFeatureDeployment?: boolean;
    AppStatus: string;
    AppVersion: string;
    AuthorId: number;
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
    SharePointAppCategory: string;
}

// Assessment Item
export interface IAssessmentItem extends Types.SP.ListItem {
    Completed?: string;
    RelatedAppId: number;
}

// Configuration
export interface IConfiguration {
    appCatalogAdminEmailGroup?: string;
    appCatalogUrl?: string;
    approvalChecklist?: string[];
    helpPageUrl?: string;
    submitChecklist?: string[];
    templatesLibraryUrl?: string;
    tenantAppCatalogUrl?: string;
    approvals: { [key: string]: string[] }
}

/**
 * Data Source
 */
export class DataSource {
    // Configuration
    private static _cfg: IConfiguration = null;
    static get Configuration(): IConfiguration { return this._cfg; }
    static loadConfiguration(): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Get the current web
            Web(Strings.SourceUrl).getFileByServerRelativeUrl(Strings.ConfigUrl).content().execute(
                // Success
                file => {
                    // Convert the string to a json object
                    let cfg = null;
                    try { cfg = JSON.parse(String.fromCharCode.apply(null, new Uint8Array(file))); }
                    catch { cfg = {}; }

                    // Set the configuration
                    this._cfg = cfg;

                    // Updates the url
                    let updateUrl = (url: string) => {
                        return url.replace("~site/", ContextInfo.webServerRelativeUrl + "/")
                            .replace("~sitecollection/", ContextInfo.siteServerRelativeUrl + "/");
                    }

                    // Replace the urls
                    this._cfg.helpPageUrl = this._cfg.helpPageUrl ? updateUrl(this._cfg.helpPageUrl) : this._cfg.helpPageUrl;
                    this._cfg.templatesLibraryUrl = this._cfg.templatesLibraryUrl ? updateUrl(this._cfg.templatesLibraryUrl) : this._cfg.templatesLibraryUrl;

                    // Resolve the request
                    resolve();
                },

                // Error
                () => {
                    // Set the configuration to nothing
                    this._cfg = {} as any;

                    // Resolve the request
                    resolve();
                }
            );
        });
    }

    // Loads the document set item
    private static _docSetInfo: Helper.IListFormResult = null;
    static get DocSetInfo(): Helper.IListFormResult { return this._docSetInfo; }
    static get DocSetItem(): IAppItem { return this.DocSetInfo ? this.DocSetInfo.item as any : null; }
    static get DocSetItemId(): number { return this.DocSetItem ? this.DocSetItem.Id : 0; }
    private static _docSetSCApp: Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata = null;
    static get DocSetSCApp(): Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata { return this._docSetSCApp; }
    private static _docSetTenantApp: Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata = null;
    static get DocSetTenantApp(): Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata { return this._docSetTenantApp; }
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
                    Expand: ["AppDevelopers", "CheckoutUser"],
                    Select: [
                        "*", "Id", "FileLeafRef", "ContentTypeId", "AppDevelopers/Id", "AppDevelopers/EMail", "CheckoutUser/Title"
                    ]
                }
            }).then(info => {
                // Save the info
                this._docSetInfo = info;

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

    // Approver Security Group
    private static _approverGroup: Types.SP.GroupOData = null;
    static get ApproverGroup(): Types.SP.GroupOData { return this._approverGroup; }
    static get ApproverUrl(): string { return ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + this._approverGroup.Id; }
    static get IsApprover(): boolean {
        // See if the group doesn't exist
        if (this.ApproverGroup == null) { return false; }

        // Parse the group
        for (let i = 0; i < this.ApproverGroup.Users.results.length; i++) {
            // See if this is the current user
            if (this.ApproverGroup.Users.results[i].Id == ContextInfo.userId) {
                // Found
                return true;
            }
        }

        // Return false by default
        return false;
    }
    private static loadApproverGroup(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the security group
            Web(Strings.SourceUrl).SiteGroups().getByName(Strings.Groups.Approvers).query({
                Expand: ["Users"]
            }).execute(group => {
                // Set the group
                this._approverGroup = group;

                // Resolve the request
                resolve();
            }, reject);
        });
    }

    // Developer Security Group
    private static _devGroup: Types.SP.GroupOData = null;
    static get DevGroup(): Types.SP.GroupOData { return this._devGroup; }
    static get DevUrl(): string { return ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + this._devGroup.Id; }
    static get IsDeveloper(): boolean {
        // See if the group doesn't exist
        if (this.DevGroup == null) { return false; }

        // Parse the group
        for (let i = 0; i < this.DevGroup.Users.results.length; i++) {
            // See if this is the current user
            if (this.DevGroup.Users.results[i].Id == ContextInfo.userId) {
                // Found
                return true;
            }
        }

        // Return false by default
        return false;
    }
    private static loadDevGroup(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the security group
            Web(Strings.SourceUrl).SiteGroups().getByName(Strings.Groups.Developers).query({
                Expand: ["Users"]
            }).execute(group => {
                // Set the group
                this._devGroup = group;

                // Resolve the request
                resolve();
            }, reject);
        });
    }

    // Deletes the app test site
    static deleteTestSite(item: IAppItem): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // Get the url to the test site
            let url = [this.Configuration.appCatalogUrl, item.AppProductID].join('/');

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

    // Loads the app test site
    static loadTestSite(item: IAppItem): PromiseLike<Types.SP.Web> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the url to the test site
            let url = [this.Configuration.appCatalogUrl, item.AppProductID].join('/');

            // Get the web
            Web(url).execute(
                // Success
                web => {
                    // Resolve the request
                    resolve(web);
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
    static init(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the configuration
            this.loadConfiguration().then(() => {
                // Load the approvers group
                this.loadApproverGroup().then(() => {
                    // Load the developers group
                    this.loadDevGroup().then(() => {
                        // Load the owners group
                        this.loadOwnerGroup().then(() => {
                            // Load the site collection apps
                            this.loadSiteCollectionApps().then(() => {
                                // Load the tenant apps
                                this.loadTenantApps().then(() => {
                                    // See if this is a document set item and not teams
                                    if (!Strings.IsTeams && this.DocSetItemId > 0) {
                                        // Load the document set item
                                        this.loadDocSet().then(() => {
                                            // Resolve the request
                                            resolve();
                                        }, reject);
                                    } else {
                                        // Load the data
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
                        }, reject);
                    }, reject);
                }, reject);
            }, reject);
        });
    }

    // Sees if an install is required and displays a dialog
    static InstallRequired(showFl: boolean = false) {
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
            if ((installFl || errors.length > 0) || showFl) {
                // Show the installation dialog
                InstallationRequired.showDialog({
                    errors,
                    onFooterRendered: el => {
                        // See if a custom error exists
                        if (errors.length > 0 || showFl) {
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

    // Loads the list data
    private static _items: IAppItem[] = null;
    static get Items(): IAppItem[] { return this._items; }
    static load(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the list items
            Web(Strings.SourceUrl).Lists(Strings.Lists.Apps).Items().query({
                Expand: ["AppDevelopers", "CheckoutUser", "Folder"],
                Filter: "ContentType eq 'App'",
                Select: [
                    "*", "Id", "FileLeafRef", "ContentTypeId", "AppDevelopers/Id", "AppDevelopers/EMail", "CheckoutUser/Title"
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
    private static _isSiteAppCatalogOwner = false;
    static get IsSiteAppCatalogOwner(): boolean { return this._isSiteAppCatalogOwner; }
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
            if (this.Configuration.appCatalogUrl) {
                // Load the available apps
                Web(this.Configuration.appCatalogUrl).SiteCollectionAppCatalog().AvailableApps(id).execute(app => {
                    // Set the app
                    this._docSetSCApp = app;

                    // Resolve the request
                    resolve();
                }, () => {
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
            if (this.Configuration.appCatalogUrl) {
                // Ensure the user is an owner of the site
                Web(this.Configuration.appCatalogUrl).AssociatedOwnerGroup().Users().getByEmail(ContextInfo.userEmail).execute(user => {
                    // Ensure the user is an owner
                    if (user && user.Id > 0) {
                        // Set the flag
                        this._isSiteAppCatalogOwner = true;

                        // Load the available apps
                        Web(this.Configuration.appCatalogUrl).SiteCollectionAppCatalog().AvailableApps().execute(apps => {
                            // Set the apps
                            this._siteCollectionApps = apps.results;

                            // Resolve the request
                            resolve();
                        }, () => {
                            // No access to the tenant app catalog
                            this._siteCollectionApps = [];

                            // Resolve the request
                            resolve();
                        });
                    }
                }, () => {
                    // Default the tenant apps
                    this._siteCollectionApps = [];

                    // Resolve the request
                    resolve();
                });
            } else {
                // Default the tenant apps
                this._siteCollectionApps = [];

                // Resolve the request
                resolve();
            }
        });
    }

    // Tenant Apps
    private static _isTenantAppCatalogOwner = false;
    static get IsTenantAppCatalogOwner(): boolean { return this._isTenantAppCatalogOwner; }
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
            if (this.Configuration.tenantAppCatalogUrl) {
                // Load the available apps
                Web(this.Configuration.tenantAppCatalogUrl).TenantAppCatalog().AvailableApps(id).execute(app => {
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
            if (this.Configuration.tenantAppCatalogUrl) {
                // Ensure the user is an owner of the site
                Web(this.Configuration.tenantAppCatalogUrl).AssociatedOwnerGroup().Users().getByEmail(ContextInfo.userEmail).execute(user => {
                    // Ensure the user is an owner
                    if (user && user.Id > 0) {
                        // Set the flag
                        this._isTenantAppCatalogOwner = true;

                        // Load the available apps
                        Web(this.Configuration.tenantAppCatalogUrl).TenantAppCatalog().AvailableApps().execute(apps => {
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
                    }
                }, () => {
                    // Default the tenant apps
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

    // User Agreement
    private static _userAgreement: string = null;
    static get UserAgreement(): string { return this._userAgreement; }
    static loadUserAgreement(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // Get the current web
            Web(Strings.SourceUrl).getFileByServerRelativeUrl(Strings.UserAgreementUrl).content().execute(
                // Success
                file => {
                    // Set the html
                    this._userAgreement = String.fromCharCode.apply(null, new Uint8Array(file));

                    // Resolve the request
                    resolve();
                },

                // Error
                () => {
                    // Set the configuration to nothing
                    this._cfg = {} as any;

                    // Resolve the request
                    resolve();
                }
            );
        });
    }

    // Owner Security Group
    private static _ownerGroup: Types.SP.GroupOData = null;
    static get OwnerGroup(): Types.SP.GroupOData { return this._ownerGroup; }
    static get IsOwner(): boolean {
        // See if the group doesn't exist
        if (this.OwnerGroup == null) { return false; }

        // Parse the group
        for (let i = 0; i < this.OwnerGroup.Users.results.length; i++) {
            // See if this is the current user
            if (this.OwnerGroup.Users.results[i].Id == ContextInfo.userId) {
                // Found
                return true;
            }
        }

        // Return false by default
        return false;
    }
    private static loadOwnerGroup(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // Load the security group
            Web(Strings.SourceUrl).AssociatedOwnerGroup().query({
                Expand: ["Users"]
            }).execute(group => {
                // Set the group
                this._ownerGroup = group;

                // Resolve the request
                resolve();
            }, resolve as any);
        });
    }
}