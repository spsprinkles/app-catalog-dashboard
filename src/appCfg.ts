import { Types, Web } from "gd-sprest-bs";
import * as Common from "./common";
import Strings from "./strings";

// Configuration
export interface IConfiguration {
    appCatalogUrl?: string;
    dashboardUrl?: string;
    helpPageUrl?: string;
    templatesLibraryUrl?: string;
    tenantAppCatalogUrl?: string;
    status: { [key: string]: IStatus };
    userAgreement?: string;
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
    createTestSite?: boolean;
    lastStep: boolean;
    name: string;
    notification?: IEmail[];
    nextStep: string;
    prevStep: string;
    requiresTechReview?: boolean;
    requiresTestCases?: boolean;
    stepNumber: number;
}

// User Types
export const UserTypes = {
    ApproversGroup: "ApproversGroup",
    Developers: "Developers",
    DevelopersGroup: "DevelopersGroup",
    Sponsor: "Sponsor",
    SponsorsGroup: "SponsorsGroup"
}

/**
 * Application Configuration
 * Reads the configuration json file.
 */
export class AppConfig {
    // Configuration
    private static _cfg: IConfiguration = null;
    static get Configuration(): IConfiguration { return this._cfg; }

    // Status
    private static _status: { [key: string]: IStatus } = null;
    static get Status(): { [key: string]: IStatus } { return this._status; }

    // Approved Status
    private static _statusApproved: string = null;
    static get ApprovedStatus(): string { return this._statusApproved; }

    // Tech Review Status
    private static _statusTechReview: string = null;
    static get TechReviewStatus(): string { return this._statusTechReview; }

    // Test Cases Status
    private static _statusTestCases: string = null;
    static get TestCasesStatus(): string { return this._statusTestCases; }

    // Validation
    static get IsValid(): boolean {
        let isValid = true;

        // Check the required properties
        isValid = this._cfg.status ? true : false;
        isValid = this._cfg.validation ? true : false;

        // Return the flag
        return isValid;
    }

    // Load the configuration file
    static loadConfiguration(cfgWebUrl?: string, cfgUrl?: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let completeRequest = () => {
                // Ensure the configuration exists
                if (this._cfg) {
                    // Replace the urls
                    this._cfg.appCatalogUrl = this.Configuration.appCatalogUrl ? Common.updateUrl(this._cfg.appCatalogUrl) : Strings.SourceUrl;
                    this._cfg.dashboardUrl = this.Configuration.dashboardUrl ? Common.updateUrl(this._cfg.dashboardUrl) : Strings.DashboardUrl;
                    this._cfg.helpPageUrl = this._cfg.helpPageUrl ? Common.updateUrl(this._cfg.helpPageUrl) : this._cfg.helpPageUrl;
                    this._cfg.templatesLibraryUrl = this._cfg.templatesLibraryUrl ? Common.updateUrl(this._cfg.templatesLibraryUrl) : this._cfg.templatesLibraryUrl;
                    this._cfg.tenantAppCatalogUrl = this.Configuration.tenantAppCatalogUrl ? Common.updateUrl(this._cfg.tenantAppCatalogUrl) : this._cfg.tenantAppCatalogUrl;

                    // Ensure the configuration is valid
                    if (this.IsValid) {
                        // Set the status and resolve the request
                        this.setStatus().then(() => { resolve(); }, reject);
                    } else {
                        // Reject the request
                        reject();
                    }
                } else {
                    // Reject the request
                    reject();
                }
            }

            // Update the configuration web url
            cfgWebUrl = '/' + (cfgWebUrl || Strings.SourceUrl).replace(/^\//, '');

            // Update the configuration file url
            cfgUrl = (cfgUrl ? Common.updateUrl(cfgUrl) : Strings.ConfigUrl).replace(/^\//, '');

            // Get the current web
            Web(cfgWebUrl).getFileByUrl(cfgUrl).content().execute(
                // Success
                file => {
                    // Convert the string to a json object
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

                    // See if this is the last step
                    if (this._status[choice].lastStep) {
                        // Set the status
                        this._statusApproved = choice;
                    }
                    // Else, see if this is the tech review status
                    else if (this._status[choice].requiresTechReview) {
                        // Set the status
                        this._statusTechReview = choice;
                    }
                    // Else, see if this is the test cases status
                    else if (this._status[choice].requiresTestCases) {
                        // Set the status
                        this._statusTestCases = choice;
                    }
                };

                // Resolve the request
                resolve();
            }, reject);
        });
    }
}