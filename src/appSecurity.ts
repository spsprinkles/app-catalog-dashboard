import { ContextInfo, Types, Web } from "gd-sprest-bs";
import { ErrorDialog } from "./errorDialog";
import Strings from "./strings";

/**
 * Application Security
 * The security group/user references for the application.
 */
export class AppSecurity {
    // Administrator
    static get IsAdmin(): boolean { return ContextInfo.isSiteAdmin; }

    // Approver Security Group
    private static _approverGroup: Types.SP.GroupOData = null;
    static get ApproverGroup(): Types.SP.GroupOData { return this._approverGroup; }
    static get ApproverUrl(): string { return ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + this._approverGroup.Id; }
    static get ApproverEmails(): string[] {
        let emails = [];

        // Parse the approvers
        for (let i = 0; i < this.ApproverGroup.Users.results.length; i++) {
            let email = this.ApproverGroup.Users.results[i].Email;

            // Append the email
            email ? emails.push(email) : null;
        }

        // Return the emails
        return emails;
    }
    static get IsApprover(): boolean {
        // See if the user is a site admin
        if (this.IsAdmin) { return true; }

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
    static get DeveloperEmails(): string[] {
        let emails = [];

        // Parse the users
        for (let i = 0; i < this.DevGroup.Users.results.length; i++) {
            let email = this.DevGroup.Users.results[i].Email;

            // Append the email
            email ? emails.push(email) : null;
        }

        // Return the emails
        return emails;
    }
    static get IsDeveloper(): boolean {
        // See if the group doesn't exist
        if (this.DevGroup == null) { return false; }

        // Parse the users
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
    // Add the user to the developer group
    static addDeveloper(userId: number): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Add the user
            this.DevGroup.Users.addUserById(userId).execute(
                user => {
                    // Append the user
                    this.DevGroup.Users.results.push(user);

                    // Resolve the request
                    resolve();
                },
                ex => {
                    // Log the error
                    ErrorDialog.show("Adding Developer", "There was an error adding the developer.", ex);
                }
            );
        });
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

    // Initialization
    static init(appCatalogUrl: string, tenantAppCatalogUrl: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
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
                                // Load the sponsor's group
                                this.loadSponsorGroup().then(() => {
                                    // Resolve the request
                                    resolve();
                                }, reject);
                            }, reject);
                        }, reject);
                    }, reject);
                }, reject);
            }, reject);
        });
    }

    // Owner Security Group
    private static _ownerGroup: Types.SP.GroupOData = null;
    static get OwnerGroup(): Types.SP.GroupOData { return this._ownerGroup; }
    static get IsOwner(): boolean {
        // See if the user is a site admin
        if (this.IsAdmin) { return true; }

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
                    // Get the user
                    Web(appCatalogUrl).SiteUsers().getByEmail(ContextInfo.userEmail).execute(user => {
                        // See if they are an admin
                        if (user.IsSiteAdmin) {
                            // Set the flag
                            this._isSiteAppCatalogOwner = true;
                        }

                        // Resolve the request
                        resolve();
                    }, () => {
                        // Resolve the request
                        resolve();
                    });
                });
            } else {
                // Resolve the request
                resolve();
            }
        });
    }

    // Sponsor Security Group
    private static _sponsorGroup: Types.SP.GroupOData = null;
    static get SponsorGroup(): Types.SP.GroupOData { return this._sponsorGroup; }
    static get SponsorUrl(): string { return ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + this._sponsorGroup.Id; }
    static get SponsorEmails(): string[] {
        let emails = [];

        // Parse the users
        for (let i = 0; i < this.SponsorGroup.Users.results.length; i++) {
            let email = this.SponsorGroup.Users.results[i].Email;

            // Append the email
            email ? emails.push(email) : null;
        }

        // Return the emails
        return emails;
    }
    // Add the user to the sponsor group
    static addSponsor(userId: number): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Add the user
            this.SponsorGroup.Users.addUserById(userId).execute(user => {
                // Append the user
                this.SponsorGroup.Users.results.push(user);

                // Resolve the request
                resolve();
            }, reject);
        });
    }
    static getSponsor(userId: number): Types.SP.User {
        // Parse the users
        for (let i = 0; i < this.SponsorGroup.Users.results.length; i++) {
            let user = this.SponsorGroup.Users.results[i];

            // Return the user if we found them
            if (user.Id == userId) { return user; }
        }

        // User not found
        return null;
    }
    static get IsSponsor(): boolean {
        // See if the group doesn't exist
        if (this.SponsorGroup == null) { return false; }

        // Parse the users
        for (let i = 0; i < this.SponsorGroup.Users.results.length; i++) {
            // See if this is the current user
            if (this.SponsorGroup.Users.results[i].Id == ContextInfo.userId) {
                // Found
                return true;
            }
        }

        // Return false by default
        return false;
    }
    private static loadSponsorGroup(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the security group
            Web(Strings.SourceUrl).SiteGroups().getByName(Strings.Groups.Sponsors).query({
                Expand: ["Users"]
            }).execute(group => {
                // Set the group
                this._sponsorGroup = group;

                // Resolve the request
                resolve();
            }, reject);
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
                    // Get the user
                    Web(tenantAppCatalogUrl).SiteUsers().getByEmail(ContextInfo.userEmail).execute(user => {
                        // See if they are an admin
                        if (user.IsSiteAdmin) {
                            // Set the flag
                            this._isSiteAppCatalogOwner = true;
                        }

                        // Resolve the request
                        resolve();
                    }, () => {
                        // Resolve the request
                        resolve();
                    });
                });
            } else {
                // Resolve the request
                resolve();
            }
        });
    }
}