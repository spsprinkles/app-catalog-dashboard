import { AuditLog, IAuditLogItemCreation, List } from "dattatable";
import { Components, ContextInfo, Helper, Types, Web } from "gd-sprest-bs";
import { AppConfig } from "./appCfg";
import { AppSecurity } from "./appSecurity";
import { ErrorDialog } from "./errorDialog";
import Strings from "./strings";

// App Item
export interface IAppItem extends Types.SP.ListItem {
    AppAPIPermissions?: string;
    AppComments?: string;
    AppDevelopers: { results: { Id: number; EMail: string; Title: string }[] };
    AppDevelopersId: { results: number[] };
    AppIsClientSideSolution?: boolean;
    AppIsDomainIsolated?: boolean;
    AppIsRejected: boolean;
    AppIsTenant: boolean;
    AppIsTenantDeployed: boolean;
    AppJustification: string;
    AppManifest: string;
    AppPermissionsJustification: string;
    AppProductID: string;
    AppPublisher: string;
    AppSharePointMinVersion?: boolean;
    AppSiteDeployments?: string;
    AppSkipFeatureDeployment?: boolean;
    AppSourceControl: Types.SP.FieldUrlValue;
    AppSponsor: { Id: number; EMail: string; Title: string };
    AppSponsorId: number;
    AppStatus: string;
    AppVersion: string;
    AuthorId: number;
    AppPackageErrorMessage?: string;
    FileLeafRef: string;

    AppDescription: string;
    AppImageURL1: Types.SP.FieldUrlValue;
    AppImageURL2: Types.SP.FieldUrlValue;
    AppImageURL3: Types.SP.FieldUrlValue;
    AppImageURL4: Types.SP.FieldUrlValue;
    AppImageURL5: Types.SP.FieldUrlValue;
    AppShortDescription: string;
    AppSupportURL: Types.SP.FieldUrlValue;
    AppThumbnailURL: Types.SP.FieldUrlValue;
    AppVideoURL: Types.SP.FieldUrlValue;
    CheckoutUser: { Id: number; Title: string };
    ContentTypeId: string;
    IsDefaultAppMetadataLocale: boolean;
    IsAppPackageEnabled: boolean;
}

// App Catalog Item
export interface IAppCatalogItem extends Types.SP.ListItemOData {
    AppCatalog_IsNewVersionAvailable: boolean;
    ApPDescription: string;
    AppImageURL1: string;
    AppImageURL2: string;
    AppImageURL3: string;
    AppImageURL4: string;
    AppImageURL5: string;
    AppMetadataLocale: string;
    AppPackageErrorMessage: string;
    AppProductID: string;
    AppPublisher: string;
    AppShortDescription: string;
    AppSupportURL: string;
    AppThumbnailURL: string;
    AppVersion: string;
    AppVideoURL: string;
    AssetID: string;
    GUID: string;
    IsAppPackageEnabled: boolean;
    IsClientSideSolution: boolean;
    IsClientSideSolutionCurrentVersionDeployed: boolean;
    IsClientSideSolutionDeployed: boolean;
    IsDefaultAppMetadataLocalte: boolean;
    IsFeaturedApp: boolean;
    IsValidAppPackage: boolean;
    SkipFeatureDeployment: boolean;
    StoreAssetId: string;
    SupportsTeamsTabs: boolean;
}

// App Catalog Request Item
export interface IAppCatalogRequestItem extends Types.SP.ListItem {
    RequestersId: { results: number[] };
    RequestNotes: string;
    RequestStatus: string;
    SiteCollectionUrl: Types.SP.FieldUrlValue;
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
    /**
     * App Assessments
     */

    // App Assessements
    private static _appAssessments: List<IAssessmentItem> = null;
    static get AppAssessments(): List<IAssessmentItem> {
        return this._appAssessments;
    }
    static initAppAssessments(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Set the requests list
            this._appAssessments = new List<IAssessmentItem>({
                listName: Strings.Lists.Assessments,
                itemQuery: { Filter: "RelatedAppId eq 0" },
                onInitError: reject,
                onInitialized: resolve
            });
        });
    }
    static loadAppAssessments(itemId: number, testFl: boolean) {
        // Query the assessments
        return this._appAssessments.refresh({
            Filter:
                "RelatedAppId eq " +
                itemId +
                " and ContentType eq '" +
                (testFl ? "TestCases" : "Item") +
                "'",
            OrderBy: ["Completed desc"],
        });
    }

    /**
     * App Catalog Requests
     */

    // App Catalog Requests
    private static _appCatalogRequests: List<IAppCatalogRequestItem> = null;
    static get AppCatalogRequests(): List<IAppCatalogRequestItem> {
        return this._appCatalogRequests;
    }
    static get HasAppCatalogRequests(): boolean {
        return this.AppCatalogRequests
            ? this.AppCatalogRequests.Items.length > 0
            : false;
    }
    static initAppCatalogRequests(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Set the requests list
            this._appCatalogRequests = new List<IAppCatalogRequestItem>({
                listName: Strings.Lists.AppCatalogRequests,
                itemQuery: {
                    Filter: "Requesters/Id eq " + ContextInfo.userId,
                },
                onInitError: reject,
                onInitialized: resolve,
            });
        });
    }

    /**
     * App Dashboard Information
     */

    // Doc Set List
    private static _docSetList: List<IAppItem> = null;
    static get DocSetList(): List<IAppItem> {
        return this._docSetList;
    }
    static initDocSetList(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Set the app dashboard list
            this._docSetList = new List<IAppItem>({
                listName: Strings.Lists.Apps,
                itemQuery: {
                    Filter: "ContentType eq 'App'",
                    GetAllItems: true,
                    Top: 5000,
                    Expand: [
                        "AppDevelopers",
                        "AppSponsor",
                        "CheckoutUser",
                        "Folder/Files",
                        "Folder/Folders",
                    ],
                    Select: [
                        "*",
                        "Id",
                        "FileLeafRef",
                        "ContentTypeId",
                        "AppDevelopers/Id",
                        "AppDevelopers/EMail",
                        "AppDevelopers/Title",
                        "AppSponsor/Id",
                        "AppSponsor/EMail",
                        "AppSponsor/Title",
                        "CheckoutUser/Id",
                        "CheckoutUser/EMail",
                        "CheckoutUser/Title",
                    ],
                },
                onInitError: reject,
                onInitialized: resolve,
            });
        });
    }

    /**
     * App Information
     * Set when viewing the details of an app
     */

    // App Catalog List Item
    private static _appCatalogItem: IAppCatalogItem = null;
    static get AppCatalogItem(): IAppCatalogItem {
        return this._appCatalogItem;
    }

    // Site Collection App Catalog Item
    private static _appCatalogSiteItem: Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata =
        null;
    static get AppCatalogSiteItem(): Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata {
        return this._appCatalogSiteItem;
    }

    // Tenant App Catalog Item
    private static _appCatalogTenantItem: Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata =
        null;
    static get AppCatalogTenantItem(): Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata {
        return this._appCatalogTenantItem;
    }

    // App Form Information
    private static _appFormInfo: Helper.IListFormResult = null;
    static get AppFormInfo(): Helper.IListFormResult { return this._appFormInfo; }

    // App Document Set Item
    private static _appItem: IAppItem = null;
    static get AppFolder(): Types.SP.FolderOData {
        return this._appItem ? (this._appItem.Folder as any) : null;
    }
    static get AppItem(): IAppItem {
        return this._appItem;
    }
    static set AppItem(newItem: IAppItem) {
        // Parse the items
        for (let i = 0; i < this.DocSetList.Items.length; i++) {
            let item = this.DocSetList.Items[i];

            // See if the id match
            if (item.Id == newItem.Id) {
                // Update the item
                this.DocSetList.Items[i] = newItem;
                return;
            }
        }

        // Append the item
        this.DocSetList.Items.push(newItem);
    }

    // Status Filters
    private static _statusFilters: Components.ICheckboxGroupItem[] = null;
    static get StatusFilters(): Components.ICheckboxGroupItem[] {
        return this._statusFilters;
    }
    static initStatusFilters(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the status field
            Web(Strings.SourceUrl)
                .Lists(Strings.Lists.Apps)
                .Fields("AppStatus")
                .execute((fld: Types.SP.FieldChoice) => {
                    let items: Components.ICheckboxGroupItem[] = [];

                    // Parse the choices
                    for (let i = 0; i < fld.Choices.results.length; i++) {
                        // Add an item
                        items.push({
                            label: fld.Choices.results[i],
                            type: Components.CheckboxGroupTypes.Switch,
                        });
                    }

                    // Set the filters and resolve the promise
                    this._statusFilters = items.sort((a, b) => {
                        if (a.label < b.label) {
                            return -1;
                        }
                        if (a.label > b.label) {
                            return 1;
                        }
                        return 0;
                    });

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
            let url = [AppConfig.Configuration.appCatalogUrl, item.AppProductID].join(
                "/"
            );

            // Log
            ErrorDialog.logInfo(`Getting test site: ${url}`);

            // Get the web context
            ContextInfo.getWeb(url).execute(
                (context) => {
                    // Log
                    ErrorDialog.logInfo(`Deleting test site: ${url}`);

                    // Delete the web
                    Web(url, {
                        requestDigest: context.GetContextWebInformation.FormDigestValue,
                    })
                        .delete()
                        .execute(
                            () => {
                                // Log
                                ErrorDialog.logInfo(
                                    `The test site '${url}' was deleted successfully...`
                                );

                                // Resolve the request
                                resolve();
                            },
                            () => {
                                // Log
                                ErrorDialog.logInfo(`The test site '${url}' doesn't exist...`);

                                // Web doesn't exist
                                resolve();
                            }
                        );
                },
                () => {
                    // Log
                    ErrorDialog.logInfo(`Error deleting the test site: ${url}`);

                    // Error getting the context
                    resolve();
                }
            );
        });
    }

    // Returns the app id from the query string
    static getAppIdFromQS(): number {
        let itemId = null;

        // Parse the query string values
        let qs = document.location.search.split("?");
        qs = qs.length > 1 ? qs[1].split("&") : [];
        for (let i = 0; i < qs.length; i++) {
            let qsItem = qs[i].split("=");
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

    // Initializes the application
    static init(appConfiguration?: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Initialize the audit log
            this.initAuditLog().then(() => {
                // Load the configuration
                AppConfig.loadConfiguration(appConfiguration).then(() => {
                    // Load the security information
                    AppSecurity.init().then(() => {
                        // Wait for the components to initialize
                        Promise.all([
                            // Initialize the app assessments
                            this.initAppAssessments(),
                            // Initialize the app catalog requests
                            this.initAppCatalogRequests(),
                            // Initialize the document set list
                            this.initDocSetList(),
                            // Load the status filters
                            this.initStatusFilters(),
                        ]).then(() => {
                            // See if an app id was defined in the query string
                            let itemId = this.getAppIdFromQS();
                            if (itemId > 0) {
                                // Load the app information
                                this.loadAppDashboard(itemId).then(resolve, resolve);
                            } else {
                                // Resolve the request
                                resolve();
                            }
                        }, reject);
                    });
                }, reject);
            }, reject);
        });
    }

    // Loads an app selected from the dashboard
    static loadAppDashboard(itemId: number): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Parse the items
            for (let i = 0; i < this.DocSetList.Items.length; i++) {
                let item = this.DocSetList.Items[i];

                // See if the id match
                if (item.Id == itemId) {
                    // Set the item
                    this._appItem = item;
                    break;
                }
            }

            // Reject the request if we don't have the item
            if (this._appItem == null) {
                reject();
                return;
            }

            // Wait for the requests to complete
            Promise.all([
                // Load the site collection app catalog item
                this.loadSiteCollectionApp(),
                // Load the tenant app catalog item
                this.loadTenantApp(),
            ]).then(() => {
                // Load the form information
                Components.ListForm.create({
                    itemId,
                    contentType: "App",
                    listName: this.DocSetList.ListName,
                    webUrl: this.DocSetList.WebUrl,
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
                    // Set the form information
                    this._appFormInfo = info;

                    // Resolve the request
                    resolve();
                }, reject);
            }, reject);
        });
    }

    // Loads the app item information from the app catalog
    private static loadSiteCollectionApp(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            let web = Web(AppConfig.Configuration.appCatalogUrl);

            // Clear the references
            this._appCatalogItem = null;
            this._appCatalogSiteItem = null;

            // Load the available apps
            web.SiteCollectionAppCatalog().AvailableApps(this.AppItem.AppProductID).execute((app) => {
                // Set the app catalog item
                this._appCatalogSiteItem = app;
            });

            // Query the app catalog and try to find it by file name
            web.Lists(Strings.Lists.AppCatalog).Items().query({ Expand: ["File"] }).execute((items) => {
                // Parse the docset items
                for (let i = 0; i < this.AppFolder.Files.results.length; i++) {
                    let file = this.AppFolder.Files.results[i];

                    // Ensure this is the package file
                    if (!file.Name.endsWith(".sppkg")) {
                        continue;
                    }

                    // Parse the items
                    for (let j = 0; j < items.results.length; j++) {
                        let item = items.results[j];

                        // See if the item matches
                        if (item.File.Name == file.Name) {
                            // Item found
                            this._appCatalogItem = item as any;
                            break;
                        }
                    }

                    // Break from the loop
                    break;
                }
            });

            // Wait for the requests to complete
            web.done(resolve);
        });
    }

    // Loads a site collection app by file name
    static loadSiteAppByName(name: string): PromiseLike<IAppCatalogItem> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the available apps
            Web(AppConfig.Configuration.appCatalogUrl)
                .Lists(Strings.Lists.AppCatalog)
                .Items()
                .query({
                    Expand: ["File"],
                    GetAllItems: true,
                    Top: 5000,
                })
                .execute(
                    (items) => {
                        // Parse the items
                        for (let i = 0; i < items.results.length; i++) {
                            let item = items.results[i];

                            // See if this is the target app
                            if (item.File.Name == name) {
                                // Resolve the request
                                resolve(item as any);
                                return;
                            }
                        }

                        // Resolve the request
                        resolve(null);
                    },
                    () => {
                        // Resolve the request
                        resolve(null);
                    }
                );
        });
    }

    // Loads the app information from the tenant app catalog
    private static loadTenantApp(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // Clear the references
            this._appCatalogTenantItem = null;

            // See if the app catalog is defined and the user is an tenant admin
            if (AppConfig.Configuration.tenantAppCatalogUrl && AppSecurity.IsTenantAppCatalogOwner) {
                // Load the available apps
                Web(AppConfig.Configuration.tenantAppCatalogUrl)
                    .TenantAppCatalog()
                    .AvailableApps(this.AppItem.AppProductID)
                    .execute(
                        (app) => {
                            // Set the app
                            this._appCatalogTenantItem = app;

                            // Resolve the request
                            resolve();
                        },
                        () => {
                            // App not found, resolve the request
                            resolve();
                        }
                    );
            } else {
                // User doesn't have access
                resolve();
            }
        });
    }

    // Loads the app test site
    static loadTestSite(item: IAppItem): PromiseLike<Types.SP.Web> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the url to the test site
            let url = [AppConfig.Configuration.appCatalogUrl, item.AppProductID].join(
                "/"
            );

            // Get the web
            Web(url).execute(
                // Success
                (web) => {
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

    // Method to refresh the data source
    static refresh(itemId?: number): PromiseLike<any> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if we need to refresh a specific app
            if (itemId > 0) {
                // Load all the items
                this.DocSetList.refreshItem(itemId).then(() => {
                    // Load the app dashboard
                    this.loadAppDashboard(itemId).then(resolve, reject);
                }, reject);
            } else {
                // Load all the items
                this.DocSetList.refresh().then(resolve, reject);
            }
        });
    }

    /**
     * Audit Log
     */

    // Audit Log States
    static AuditLogStates = {
        AppAdded: "App Added",
        AppApproved: "App Approved",
        AppCatalogRequest: "App Catalog Request",
        AppDeployed: "App Deployed",
        AppRejected: "App Rejected",
        AppRetracted: "App Retracted",
        AppResubmitted: "App Resubmitted",
        AppSubmitted: "App Submitted",
        AppTenantDeployed: "App Tenant Deployed",
        AppUpdated: "App Updated",
        AppUpgraded: "App Upgraded",
        CreateTestSite: "Create Test Site",
        DeleteApp: "Delete App",
        DeleteTestSite: "Delete Test Site"
    }

    // Audit Log
    private static _auditLog: AuditLog = null;
    static get AuditLog(): AuditLog {
        return this._auditLog;
    }
    private static initAuditLog(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Intialize the audit log
            this._auditLog = new AuditLog({
                listName: Strings.Lists.AuditLog,
                onInitError: reject,
                onInitialized: resolve,
            });
        });
    }
    static logItem(values: IAuditLogItemCreation, item: IAppItem) {
        // Set the log data
        values.LogData = Helper.stringify(item);

        // Log the item
        this.AuditLog.logItem(values);
    }
}