import { Components, ContextInfo, Helper, List, Types, Web } from "gd-sprest-bs";
import Strings from "./strings";

// App Item
export interface IAppItem extends Types.SP.ListItem {
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
    ContentTypeId: string;
    DevAppStatus: string;
    FileLeafRef: string;
    IsDefaultAppMetadataLocale: boolean;
    IsAppPackageEnabled: boolean;
    Owners: { results: { Id: number; EMail: string; }[] };
    OwnersId: { results: number[] };
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
    helpPageUrl?: string;
    templatesLibraryUrl?: string;
    tenantAppCatalogUrl?: string;
}

/**
 * Data Source
 */
export class DataSource {
    // Apps
    private static _apps: Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata[] = null;
    static get Apps(): Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata[] { return this._apps; }
    static getAppById(appId: string): Types.Microsoft.SharePoint.Marketplace.CorporateCuratedGallery.CorporateCatalogAppMetadata {
        // Parse the apps
        for (let i = 0; i < this._apps.length; i++) {
            let app = this._apps[i];

            // See if this is the target app
            if (app.ProductId == appId) { return app; }
        }

        // App not found
        return null;
    }
    static loadApps(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the available apps
            Web().TenantAppCatalog().AvailableApps().execute(apps => {
                // Set the apps
                this._apps = apps.results;

                // Resolve the request
                resolve();
            }, () => {
                // No access to the tenant app catalog
                this._apps = [];

                // Resolve the request
                resolve();
            });
        });
    }

    // Configuration
    private static _cfg: IConfiguration = null;
    static get Configuration(): IConfiguration { return this._cfg; }
    static loadConfiguration(): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Get the current web
            Web().getFileByServerRelativeUrl(Strings.ConfigUrl).content().execute(
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

    // Gets the document set item id from the query string
    private static _docSetItemId: number = 0;
    static get IsDocSet(): boolean { return this.DocSetItemId > 0; }
    static get DocSetItemId(): number {
        // Return a valid value
        if (this._docSetItemId <= 0) {
            // Get the id from the querystring
            let qs = document.location.search.split('?');
            qs = qs.length > 1 ? qs[1].split('&') : [];
            for (let i = 0; i < qs.length; i++) {
                let qsItem = qs[i].split('=');
                let key = qsItem[0].toLowerCase();
                let value = qsItem[1];

                // See if this is the "id" key
                if (key == "id") {
                    // Return the item
                    this._docSetItemId = parseInt(value);
                }
            }
        }

        // Return the workspace item id
        return this._docSetItemId;
    }

    // Loads the document set item
    private static _docSetInfo: Helper.IListFormResult = null;
    static get DocSetInfo(): Helper.IListFormResult { return this._docSetInfo; }
    static get DocSetItem(): IAppItem { return this._docSetInfo.item as any; }
    static loadDocSet(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the list information
            Components.ListForm.create({
                listName: Strings.Lists.Apps,
                itemId: this.DocSetItemId,
                contentType: "App",
                query: {
                    Expand: ["CheckoutUser", "Owners"],
                    Select: [
                        "Id", "FileLeafRef", "CheckoutUser/Title", "AppThumbnailURL", "AuthorId",
                        "OwnersId", "DevAppStatus", 'Title', "AppProductID", "AppVersion", "AppPublisher", "Owners",
                        "IsAppPackageEnabled", "Owners/Id", "Owners/EMail"
                    ]
                }
            }).then(info => {
                // Save the info
                this._docSetInfo = info;

                // Resolve the request
                resolve();
            }, reject);
        });
    }

    // Approver Security Group
    private static _approverGroup: Types.SP.GroupOData = null;
    static get ApproverGroup(): Types.SP.GroupOData { return this._approverGroup; }
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
            Web().SiteGroups().getByName(Strings.Groups.Approvers).query({
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
            Web().SiteGroups().getByName(Strings.Groups.Developers).query({
                Expand: ["Users"]
            }).execute(group => {
                // Set the group
                this._devGroup = group;

                // Resolve the request
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
            List(Strings.Lists.Apps).Fields("DevAppStatus").execute((fld: Types.SP.FieldChoice) => {
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
                        // Load the apps
                        this.loadApps().then(() => {
                            // See if this is a document set item
                            if (this.DocSetItemId > 0) {
                                // Load the templates
                                this.loadTemplates().then(() => {
                                    // Load the document set item
                                    this.loadDocSet().then(() => {
                                        // Resolve the request
                                        resolve();
                                    }, reject);
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
        });
    }

    // Loads the list data
    private static _items: IAppItem[] = null;
    static get Items(): IAppItem[] { return this._items; }
    static load(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the list items
            List(Strings.Lists.Apps).Items().query({
                Expand: ["CheckoutUser", "Folder", "Owners"],
                Filter: "ContentType eq 'App'",
                Select: [
                    "Id", "FileLeafRef", "ContentTypeId", "CheckoutUser/Title", "AppThumbnailURL", "AuthorId",
                    "OwnersId", "DevAppStatus", 'Title', "AppProductID", "AppVersion", "AppPublisher", "Owners",
                    "IsAppPackageEnabled", "Owners/Id", "Owners/EMail"
                ]
            }).execute(items => {
                // Set the items
                this._items = items.results as any;

                // Resolve the promise
                resolve();
            }, reject);
        });
    }

    // Loads the templates
    private static _templates: Types.SP.File[] = null;
    static get Templates(): Types.SP.File[] { return this._templates; }
    static loadTemplates(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if the templates url exists
            if (this.Configuration.templatesLibraryUrl) {
                // Load the library
                Web().getFolderByServerRelativeUrl(this.Configuration.templatesLibraryUrl).Files().execute(files => {
                    // Save the files
                    this._templates = files.results.sort((a, b) => {
                        if (a.Name < b.Name) { return -1; }
                        if (a.Name > b.Name) { return 1; }
                        return 0;
                    });

                    // Resolve the request
                    resolve();
                }, reject);
            } else {
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
            Web().getFileByServerRelativeUrl(Strings.UserAgreementUrl).content().execute(
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
}