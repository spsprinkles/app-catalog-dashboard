import { ContextInfo, Types, Web } from "gd-sprest-bs";
import Strings from "./strings";

// Configuration
export interface IConfiguration {
    appCatalogUrl?: string;
    helpPageUrl?: string;
    templatesLibraryUrl?: string;
    tenantAppCatalogUrl?: string;
    status: { [key: string]: IStatus };
    validation: {
        techReview: { [key: string]: string[]; }
        testCases: { [key: string]: string[]; }
    }
}

// Email
export interface IEmail {
    content: string;
    to: string[];
    cc: string[];
    subject: string;
}

// Status
export interface IStatus {
    actions?: string[];
    approval?: string[];
    checklists?: string[];
    lastStep: boolean;
    name: string;
    notification?: IEmail[];
    nextStep: string;
    prevStep: string;
    requiresTechReview?: boolean;
    requiresTestCase?: boolean;
    stepNumber: number;
}

/**
 * Application Configuration
 * Reads the configuration json file.
 */
export class AppConfig {
    // Status
    private static _status: { [key: string]: IStatus } = null;
    static get Status(): { [key: string]: IStatus } { return this._status; }

    // Configuration
    private static _cfg: IConfiguration = null;
    static get Configuration(): IConfiguration { return this._cfg; }

    // Load the configuration file
    static loadConfiguration(cfg?: any): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let completeRequest = () => {
                // Ensure the configuration exists
                if (this._cfg) {
                    // Updates the url
                    let updateUrl = (url: string) => {
                        return url.replace("~site/", ContextInfo.webServerRelativeUrl + "/")
                            .replace("~sitecollection/", ContextInfo.siteServerRelativeUrl + "/");
                    }

                    // Replace the urls
                    this._cfg.helpPageUrl = this._cfg.helpPageUrl ? updateUrl(this._cfg.helpPageUrl) : this._cfg.helpPageUrl;
                    this._cfg.templatesLibraryUrl = this._cfg.templatesLibraryUrl ? updateUrl(this._cfg.templatesLibraryUrl) : this._cfg.templatesLibraryUrl;

                    // Set the status and resolve the request
                    this.setStatus().then(() => { resolve(); });
                } else {
                    // Reject the request
                    reject();
                }
            }

            // See if the configuration is being provided
            if (cfg) {
                try { this._cfg = JSON.parse(cfg); }
                catch { this._cfg = null; }

                // Complete the request
                completeRequest();
            } else {
                // Get the current web
                Web(Strings.SourceUrl).getFileByServerRelativeUrl(Strings.ConfigUrl).content().execute(
                    // Success
                    file => {
                        // Convert the string to a json object
                        let cfg = null;
                        try { this._cfg = JSON.parse(String.fromCharCode.apply(null, new Uint8Array(file))); }
                        catch { this._cfg = null; }

                        // Complete the request
                        completeRequest();
                    },

                    // Error
                    () => {
                        // Clear the configuration
                        this._cfg = null;

                        // Complete the request
                        completeRequest();
                    }
                );
            }
        });
    }

    private static setStatus(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Clear the status information
            this._status = {};

            // Read the status field
            Web(Strings.SourceUrl).Lists(Strings.Lists.Apps).Fields("AppStatus").execute((field: Types.SP.FieldChoice) => {
                // Parse the choices
                for (let i = 0; i < field.Choices.results.length; i++) {
                    let choice = field.Choices.results[i];
                    let nextStep = field.Choices.results[i + 1];
                    let prevStep = i > 0 ? field.Choices.results[i - 1] : null;

                    // Set the status
                    this._status[choice] = this.Configuration.status[choice] || {} as any;
                    this._status[choice].lastStep = nextStep ? false : true;
                    this._status[choice].name = choice;
                    this._status[choice].nextStep = nextStep;
                    this._status[choice].prevStep = prevStep;
                    this._status[choice].stepNumber = i;
                };

                // Resolve the request
                resolve();
            }, reject);
        });
    }
}