import { Components, List, Types, Web } from "gd-sprest-bs";
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
    DevAppStatus: string;
    IsDefaultAppMetadataLocale: boolean;
    IsAppPackageEnabled: boolean;
    Owners: { results: { Id: number; EMail: string; }[] }
    SharePointAppCategory: string;
}

// Assessment Item
export interface IAssessmentItem extends Types.SP.ListItem {
}

// Configuration
export interface IConfiguration {
    appCatalogAdminEmailGroup?: string;
    tenantAppCatalog?: string;
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
            Web().getFileByServerRelativeUrl(Strings.ConfigUrl).content().execute(
                // Success
                file => {
                    // Convert the string to a json object
                    let cfg = null;
                    try { cfg = JSON.parse(String.fromCharCode.apply(null, new Uint8Array(file))); }
                    catch { cfg = {}; }

                    // Set the configuration
                    this._cfg = cfg;

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

    // Gets the item id from the query string
    static getItemIdFromQS() {
        // Get the id from the querystring
        let qs = document.location.search.split('?');
        qs = qs.length > 1 ? qs[1].split('&') : [];
        for (let i = 0; i < qs.length; i++) {
            let qsItem = qs[i].split('=');
            let key = qsItem[0];
            let value = qsItem[1];

            // See if this is the "id" key
            if (key == "ID") {
                // Return the item
                return parseInt(value);
            }
        }
    }

    // Initializes the application
    static init(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the configuration
            this.loadConfiguration().then(() => {
                // Load the data
                this.load().then(() => {
                    // Load the status filters
                    this.loadStatusFilters().then(() => {
                        // Resolve the request
                        resolve();
                    }, reject);
                }, reject)
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
                Expand: ["CheckoutUser", "Owners"],
                Select: [
                    "Id", "FileLeafRef", "CheckoutUser/Title", "AppThumbnailURL", "AuthorId",
                    "OwnersId", "DevAppStatus", 'Title', "AppVersion", "AppPublisher", "Owners",
                    "IsAppPackageEnabled", "Owners/Id", "Owners/EMail"
                ]
            }).execute(items => {
                // Set the items
                this._items = items.results as any;

                // Resolve the promise
                resolve();
            });
        });
    }
}