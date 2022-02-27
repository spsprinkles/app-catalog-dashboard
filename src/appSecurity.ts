import { ContextInfo, Types, Web } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * Application Security
 * The security group/user references for the application.
 */
export class AppSecurity {
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

    // Site Collection App Catalog Owner
    private static _isSiteAppCatalogOwner = false;
    static get IsSiteAppCatalogOwner(): boolean { return this._isSiteAppCatalogOwner; }
    private static initSCOwner(appCatalogUrl: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // See if the app catalog is defined
            if (appCatalogUrl) {
                // Ensure the user is an owner of the site
                Web(appCatalogUrl).AssociatedOwnerGroup().Users().getByEmail(ContextInfo.userEmail).execute(user => {
                    // Ensure the user is an owner
                    if (user && user.Id > 0) {
                        // Set the flag
                        this._isSiteAppCatalogOwner = true;
                    }

                    // resolve the request
                    resolve();
                }, () => {
                    // Resolve the request
                    resolve();
                });
            }
        });
    }

    // Tenant App Catalog Owner
    private static _isTenantAppCatalogOwner = false;
    static get IsTenantAppCatalogOwner(): boolean { return this._isTenantAppCatalogOwner; }
    private static initTenantOwner(tenantAppCatalogUrl: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // See if the tenant app catalog is defined
            if (tenantAppCatalogUrl) {
                // Ensure the user is an owner of the site
                Web(tenantAppCatalogUrl).AssociatedOwnerGroup().Users().getByEmail(ContextInfo.userEmail).execute(user => {
                    // Ensure the user is an owner
                    if (user && user.Id > 0) {
                        // Set the flag
                        this._isTenantAppCatalogOwner = true;
                    }

                    // Resolve the request
                    resolve();
                }, () => {
                    // Resolve the request
                    resolve();
                });
            }
        });
    }

    // Initialization
    static init(appCatalogUrl: string, tenantAppCatalogUrl: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Initialize the SC Owner
            this.initSCOwner(appCatalogUrl).then(() => {
                // Initialize the tenant owner
                this.initTenantOwner(tenantAppCatalogUrl).then(() => {
                    // Load the approver's group
                    this.loadApproverGroup().then(() => {
                        // Load the developer's group
                        this.loadDevGroup().then(() => {
                            // Load the owner's group
                            this.loadOwnerGroup().then(() => {
                                // Resolve the request
                                resolve();
                            });
                        });
                    });
                });
            });
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