import { ContextInfo, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { AppConfig } from "./appCfg";
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
            }, () => {
                // Log
                ErrorDialog.logError("[App Security] Error loading the approver group.");

                // Reject the request
                reject();
            });
        });
    }

    // Current User
    private static _currentUser: Types.SP.User = null;
    static get CurrentUser(): Types.SP.User { return this._currentUser; }
    static loadCurrentUser(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the current user
            Web(Strings.SourceUrl).CurrentUser().execute(user => {
                // Set the user
                this._currentUser = user;

                // Resolve the request
                resolve();
            }, () => {
                // Log
                ErrorDialog.logError("[App Security] Error loading the current user.");

                // Reject the request
                reject();
            });
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
            }, () => {
                // Log
                ErrorDialog.logError("[App Security] Error loading the developer group.");

                // Reject the request
                reject();
            });
        });
    }

    // Approver Security Group
    private static _finalApproverGroup: Types.SP.GroupOData = null;
    static get FinalApproverGroup(): Types.SP.GroupOData { return this._finalApproverGroup; }
    static get FinalApproverUrl(): string { return ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + this._finalApproverGroup.Id; }
    static get FinalApproverEmails(): string[] {
        let emails = [];

        // Parse the approvers
        for (let i = 0; i < this.FinalApproverGroup.Users.results.length; i++) {
            let email = this.FinalApproverGroup.Users.results[i].Email;

            // Append the email
            email ? emails.push(email) : null;
        }

        // Return the emails
        return emails;
    }
    static get IsFinalApprover(): boolean {
        // See if the user is a site admin
        if (this.IsAdmin) { return true; }

        // See if the group doesn't exist
        if (this.FinalApproverGroup == null) { return false; }

        // Parse the group
        for (let i = 0; i < this.FinalApproverGroup.Users.results.length; i++) {
            // See if this is the current user
            if (this.FinalApproverGroup.Users.results[i].Id == ContextInfo.userId) {
                // Found
                return true;
            }
        }

        // Return false by default
        return false;
    }
    private static loadFinalApproverGroup(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the security group
            Web(Strings.SourceUrl).SiteGroups().getByName(Strings.Groups.FinalApprovers).query({
                Expand: ["Users"]
            }).execute(group => {
                // Set the group
                this._finalApproverGroup = group;

                // Resolve the request
                resolve();
            }, () => {
                // Log
                ErrorDialog.logError("[App Security] Error loading the final approver group.");

                // Reject the request
                reject();
            });
        });
    }

    // Approver Security Group
    private static _helpdeskGroup: Types.SP.GroupOData = null;
    static get HelpdeskGroup(): Types.SP.GroupOData { return this._helpdeskGroup; }
    static get HelpdeskUrl(): string { return ContextInfo.webServerRelativeUrl + "/_layouts/15/people.aspx?MembershipGroupId=" + this._helpdeskGroup.Id; }
    static get HelpdeskEmails(): string[] {
        let emails = [];

        // Parse the approvers
        for (let i = 0; i < this.HelpdeskGroup.Users.results.length; i++) {
            let email = this.HelpdeskGroup.Users.results[i].Email;

            // Append the email
            email ? emails.push(email) : null;
        }

        // Return the emails
        return emails;
    }
    private static loadHelpdeskGroup(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Ensure the group name exists
            if (AppConfig.Configuration.helpdeskGroupName) {
                // Load the security group
                Web(Strings.SourceUrl).SiteGroups().getByName(AppConfig.Configuration.helpdeskGroupName).query({
                    Expand: ["Users"]
                }).execute(group => {
                    // Set the group
                    this._finalApproverGroup = group;

                    // Resolve the request
                    resolve();
                }, () => {
                    // Log
                    ErrorDialog.logError("[App Security] Error loading the helpdesk group.");

                    // Reject the request
                    reject();
                });
            }
            else {
                // Resolve the request
                resolve();
            }
        });
    }

    // Initialization
    static init(appCatalogUrl: string, tenantAppCatalogUrl: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Execute the requests
            Promise.all([
                // Load the current user
                this.loadCurrentUser(),
                // Initialize the SC Owner
                this.initSCAppCatalogOwner(appCatalogUrl),
                // Initialize the tenant owner
                this.initTenantAppCatalogOwner(tenantAppCatalogUrl),
                // Load the approver's group
                this.loadApproverGroup(),
                // Load the final approver's group
                this.loadFinalApproverGroup(),
                // Load the helpdesk group
                this.loadHelpdeskGroup(),
                // Load the developer's group
                this.loadDevGroup(),
                // Load the owner's group
                this.loadOwnerGroup(),
                // Load the sponsor's group
                this.loadSponsorGroup()
            ]).then(() => {
                // Resolve the request
                resolve();
            }, () => {
                // Log
                ErrorDialog.logError("[App Security] Error initializing the App Security component.");

                // Reject the request
                reject();
            });
        });
    }

    // Determines if the user is an admin of a site collection
    static isOwner(url: string): PromiseLike<{ url: string; isOwner: boolean; isRoot: boolean; }> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let isOwner = false;
            let isRoot = false;
            let userId = null;
            let web = Web(url);
            let webUrl: string = null;

            // Get the web
            web.query({ Expand: ["CurrentUser", "ParentWeb"], Select: ["Id", "Url"] }).execute(web => {
                // Set the web url
                webUrl = web.Url;

                // Set the user id
                userId = web.CurrentUser.Id;

                // Set the flag
                isRoot = web.ParentWeb.Id == null || web.ParentWeb.Id == web.Id;

                // See if the user is a SCA
                if (web.CurrentUser.IsSiteAdmin) {
                    // Set the flag
                    isOwner = true;
                }
            }, reject);

            // Get the owners group
            web.AssociatedOwnerGroup().Users().execute(users => {
                // Parse the users
                for (let i = 0; i < users.results.length; i++) {
                    // See if the user is in the group
                    if (users.results[i].Id == userId) {
                        // Set the flag and break from the loop
                        isOwner = true;
                        break;
                    }
                }
            }, reject, true);

            // Wait for the requests to complete
            web.done(() => {
                // Ensure the web url was set and resolve the request
                // Otherwise, the url entered is not valid
                webUrl ? resolve({ url: webUrl, isOwner, isRoot }) : null;
            });
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
            }, () => {
                // Log
                ErrorDialog.logError("[App Security] Error loading the owner group.");

                // Resolve the request
                resolve();
            });
        });
    }

    // Site Collection App Catalog Owner
    private static _isSiteAppCatalogOwner = false;
    static get IsSiteAppCatalogOwner(): boolean { return this._isSiteAppCatalogOwner; }
    private static initSCAppCatalogOwner(appCatalogUrl: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // See if the app catalog is defined
            if (appCatalogUrl) {
                // Get the user's permissions to the app catalog list
                Web(appCatalogUrl).Lists(Strings.Lists.AppCatalog).query({
                    Expand: ["EffectiveBasePermissions"]
                }).execute(list => {
                    // See if the user has right access
                    this._isSiteAppCatalogOwner = Helper.hasPermissions(list.EffectiveBasePermissions, [
                        SPTypes.BasePermissionTypes.FullMask
                    ]);

                    // Resolve the request
                    resolve();
                }, () => {
                    // Resolve the request
                    resolve();
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
        // Log
        ErrorDialog.logInfo(`Adding the user with id '${userId}' to the App Sponsor site group...`);

        // Return a promise
        return new Promise((resolve, reject) => {
            // Add the user
            this.SponsorGroup.Users.addUserById(userId).execute(user => {
                // Append the user
                this.SponsorGroup.Users.results.push(user);

                // Resolve the request
                resolve();
            },
                // Error
                ex => {
                    // Log
                    ErrorDialog.logError(`Error adding the app sponsor not added to the group.`);
                    ErrorDialog.logError(`Error Message: ${ex.response}`)

                    // Reject the request
                    reject();
                });
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
            }, () => {
                // Log
                ErrorDialog.logError("[App Security] Error loading the sponsor group.");

                // Reject the request
                reject();
            });
        });
    }

    // Tenant App Catalog Owner
    private static _isTenantAppCatalogOwner = false;
    static get IsTenantAppCatalogOwner(): boolean { return this._isTenantAppCatalogOwner; }
    private static initTenantAppCatalogOwner(tenantAppCatalogUrl: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // See if the tenant app catalog is defined
            if (tenantAppCatalogUrl) {
                // Get the user's permissions to the app catalog list
                Web(tenantAppCatalogUrl).Lists(Strings.Lists.AppCatalog).query({
                    Expand: ["EffectiveBasePermissions"]
                }).execute(list => {
                    // See if the user has right access
                    this._isTenantAppCatalogOwner = Helper.hasPermissions(list.EffectiveBasePermissions, [
                        SPTypes.BasePermissionTypes.FullMask
                    ]);

                    // Resolve the request
                    resolve();
                }, () => {
                    // Resolve the request
                    resolve();
                });
            } else {
                // Resolve the request
                resolve();
            }
        });
    }
}

/**
 * Create Security Groups
 */
export class CreateSecurityGroups {
    // Creates the security groups
    static run(): PromiseLike<any> {
        return Promise.all([
            this.createAppCatalogSecurityGroups(),
            this.createAppSecurityGroups()
        ]);
    }

    // Creates the security groups for the app site
    private static createAppSecurityGroups(webUrl: string = Strings.SourceUrl): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // Get the site groups
            this.getWebSiteGroups(webUrl, ContextInfo.formDigestValue).then(webGroups => {
                // Create the site groups
                this.createSiteGroups(webUrl, ContextInfo.formDigestValue, Strings.Groups).then(groups => {
                    // Get the definitions
                    this.getPermissionTypes(webUrl, ContextInfo.formDigestValue).then(permissions => {
                        // Reset group settings
                        Promise.all([
                            // Set the dev group owner
                            Helper.setGroupOwner(webGroups[Strings.Groups.Developers], webGroups[Strings.Groups.Approvers], Strings.SourceUrl),
                            // Set the sponsor group owner
                            Helper.setGroupOwner(webGroups[Strings.Groups.Sponsors], webGroups[Strings.Groups.Developers], Strings.SourceUrl),
                            // Reset the list permissions
                            this.resetListPermissions([Strings.Lists.Apps, Strings.Lists.Assessments, Strings.Lists.AppCatalogRequests])
                        ]).then(() => {
                            // Get the lists to update
                            let listApps = Web().Lists(Strings.Lists.Apps);
                            let listAssessments = Web().Lists(Strings.Lists.Assessments);
                            let listRequests = Web().Lists(Strings.Lists.AppCatalogRequests);

                            // Ensure the approver group exists
                            if (groups[Strings.Groups.Approvers]) {
                                // Set the list permissions
                                listApps.RoleAssignments().addRoleAssignment(groups[Strings.Groups.Approvers].Id, permissions[SPTypes.RoleType.WebDesigner]).execute(() => {
                                    // Log
                                    console.log("[Apps List] The approver permission was added successfully.");
                                });
                                listAssessments.RoleAssignments().addRoleAssignment(groups[Strings.Groups.Approvers].Id, permissions[SPTypes.RoleType.WebDesigner]).execute(() => {
                                    // Log
                                    console.log("[Assessments List] The approver permission was added successfully.");
                                });
                                listRequests.RoleAssignments().addRoleAssignment(groups[Strings.Groups.Approvers].Id, permissions[SPTypes.RoleType.WebDesigner]).execute(() => {
                                    // Log
                                    console.log("[App Catalog Requests List] The approver permission was added successfully.");
                                });
                            }

                            // Ensure the dev group exists
                            if (groups[Strings.Groups.Developers]) {
                                // Set the list permissions
                                listApps.RoleAssignments().addRoleAssignment(groups[Strings.Groups.Developers].Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                                    // Log
                                    console.log("[Apps List] The dev permission was added successfully.");
                                });
                                listAssessments.RoleAssignments().addRoleAssignment(groups[Strings.Groups.Developers].Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                                    // Log
                                    console.log("[Assessments List] The dev permission was added successfully.");
                                });
                                listRequests.RoleAssignments().addRoleAssignment(groups[Strings.Groups.Developers].Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                                    // Log
                                    console.log("[App Catalog Requests List] The dev permission was added successfully.");
                                });
                            }

                            // Ensure the final approver group exists
                            if (groups[Strings.Groups.FinalApprovers]) {
                                // Set the list permissions
                                listApps.RoleAssignments().addRoleAssignment(groups[Strings.Groups.FinalApprovers].Id, permissions[SPTypes.RoleType.WebDesigner]).execute(() => {
                                    // Log
                                    console.log("[Apps List] The approver permission was added successfully.");
                                });
                                listAssessments.RoleAssignments().addRoleAssignment(groups[Strings.Groups.FinalApprovers].Id, permissions[SPTypes.RoleType.WebDesigner]).execute(() => {
                                    // Log
                                    console.log("[Assessments List] The approver permission was added successfully.");
                                });
                                listRequests.RoleAssignments().addRoleAssignment(groups[Strings.Groups.FinalApprovers].Id, permissions[SPTypes.RoleType.WebDesigner]).execute(() => {
                                    // Log
                                    console.log("[App Catalog Requests List] The approver permission was added successfully.");
                                });
                            }

                            // Ensure the dev group exists
                            if (groups[Strings.Groups.Sponsors]) {
                                // Set the list permissions
                                listApps.RoleAssignments().addRoleAssignment(groups[Strings.Groups.Sponsors].Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                                    // Log
                                    console.log("[Apps List] The sponsor permission was added successfully.");
                                });
                                listAssessments.RoleAssignments().addRoleAssignment(groups[Strings.Groups.Sponsors].Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                                    // Log
                                    console.log("[Assessments List] The sponsor permission was added successfully.");
                                });
                                listRequests.RoleAssignments().addRoleAssignment(groups[Strings.Groups.Sponsors].Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                                    // Log
                                    console.log("[App Catalog Requests List] The sponsor permission was added successfully.");
                                });
                            }

                            // Ensure the default members group exists
                            if (webGroups.members) {
                                // Set the list permissions
                                listApps.RoleAssignments().addRoleAssignment(webGroups.members.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                                    // Log
                                    console.log("[Apps List] The default members permission was added successfully.");
                                });
                                listAssessments.RoleAssignments().addRoleAssignment(webGroups.members.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                                    // Log
                                    console.log("[Assessments List] The default members permission was added successfully.");
                                });
                                listRequests.RoleAssignments().addRoleAssignment(webGroups.members.Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                                    // Log
                                    console.log("[App Catalog Requests List] The default members permission was added successfully.");
                                });
                            }

                            // Ensure the default owners group exists
                            if (webGroups.owners) {
                                // Set the list permissions
                                listApps.RoleAssignments().addRoleAssignment(webGroups.owners.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                                    // Log
                                    console.log("[Apps List] The default owners permission was added successfully.");
                                });
                                listAssessments.RoleAssignments().addRoleAssignment(webGroups.owners.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                                    // Log
                                    console.log("[Assessments List] The default owners permission was added successfully.");
                                });
                                listRequests.RoleAssignments().addRoleAssignment(webGroups.owners.Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                                    // Log
                                    console.log("[App Catalog Requests List] The default owners permission was added successfully.");
                                });
                            }

                            // Ensure the default members group exists
                            if (webGroups.visitors) {
                                // Set the list permissions
                                listApps.RoleAssignments().addRoleAssignment(webGroups.visitors.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                                    // Log
                                    console.log("[Apps List] The default visitors permission was added successfully.");
                                });
                                listAssessments.RoleAssignments().addRoleAssignment(webGroups.visitors.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                                    // Log
                                    console.log("[Assessments List] The default visitors permission was added successfully.");
                                });
                                listRequests.RoleAssignments().addRoleAssignment(webGroups.visitors.Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                                    // Log
                                    console.log("[App Catalog Requests List] The default visitors permission was added successfully.");
                                });
                            }

                            // Wait for the app list updates to complete
                            listApps.done(() => {
                                // Wait for the assessment list updates to complete
                                listAssessments.done(() => {
                                    // Wait for the requests list updates to complete
                                    listRequests.done(() => {
                                        // Resolve the request
                                        resolve();
                                    });
                                });
                            });
                        });
                    });
                });
            });
        });
    }

    // Creates the security groups for the app catalog site
    private static createAppCatalogSecurityGroups(webUrl: string = AppConfig.Configuration.appCatalogUrl): PromiseLike<any> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the context information
            ContextInfo.getWeb(webUrl).execute(context => {
                // Execute the requests
                Promise.all([
                    // Get the site groups
                    this.getWebSiteGroups(webUrl, context.GetContextWebInformation.FormDigestValue),
                    // Create the site groups
                    this.createSiteGroups(webUrl, ContextInfo.formDigestValue, Strings.Groups),
                    // Get the definitions
                    this.getPermissionTypes(webUrl, ContextInfo.formDigestValue, true)
                ]).then(resolve, reject);
            }, reject);
        });
    }

    // Creates the site groups
    private static createSiteGroups(url: string, requestDigest: string, groups: { [key: string]: string }): PromiseLike<{ [key: string]: Types.SP.Group }> {
        let siteGroups: { [key: string]: Types.SP.Group } = {};

        // Return a promise
        return new Promise((resolve) => {
            // Get the group names
            let groupNames: string[] = [];
            for (let groupName in groups) {
                // Add the name
                groupNames.push(groupName);
            }

            // Parse the groups to create
            Helper.Executor(groupNames, groupName => {
                // Return a promise
                return new Promise((resolve, reject) => {
                    // Get the group
                    Web(url, { requestDigest }).SiteGroups().getByName(groupName).execute(
                        // Exists
                        group => {
                            // Save the group
                            siteGroups[groupName] = group;

                            // Check the next group
                            resolve(null);
                        },

                        // Doesn't exist
                        () => {
                            let isDevGroup = groupName == Strings.Groups.Developers;
                            let isSponsorGroup = groupName == Strings.Groups.Sponsors;

                            // Create the group
                            Web().SiteGroups().add({
                                AllowMembersEditMembership: true,
                                AllowRequestToJoinLeave: isDevGroup || isSponsorGroup,
                                AutoAcceptRequestToJoinLeave: isDevGroup || isSponsorGroup,
                                Title: groupName,
                                Description: "Group contains the '" + groupName + "' users.",
                                OnlyAllowMembersViewMembership: false
                            }).execute(
                                // Successful
                                group => {
                                    // Save the group
                                    siteGroups[groupName] = group;

                                    // Check the next group
                                    resolve(null);
                                },

                                // Error
                                ex => {
                                    // Log the error
                                    ErrorDialog.show("Security Group", "There was an error creating the security group.", ex);

                                    // Reject the request
                                    reject();
                                }
                            );
                        }
                    );
                });
            }).then(() => {
                // Resolve the request
                resolve(siteGroups);
            });
        });
    }

    // Gets the custom permission level
    static CustomPermissionLevel: string = "Contribute and Manage Subwebs";
    private static getCustomPermissionLevel(url: string, roles: { [key: number]: number }): PromiseLike<{ [key: number]: number }> {

        // Return a promise
        return new Promise(resolve => {
            // See if the roles contain the custom permission
            if (roles[this.CustomPermissionLevel]) {
                // Resolve the request
                resolve(roles);
            } else {
                // Create the custom permission
                Helper.copyPermissionLevel({
                    BasePermission: "Contribute",
                    Name: this.CustomPermissionLevel,
                    Description: "Extends the contribute permission level and adds the ability to create a subweb.",
                    AddPermissions: [SPTypes.BasePermissionTypes.ManageSubwebs],
                    WebUrl: url
                }).then(
                    role => {
                        // Update the mapper
                        roles[this.CustomPermissionLevel] = role.Id;

                        // Resolve the request
                        resolve(roles);
                    },
                    ex => {
                        // Log the error
                        ErrorDialog.show("Permission Level", "There was an error creating the contribute and manage subwebs custom permission level.", ex);
                    }
                );
            }
        });

    }

    // Gets the role definitions for the permission types
    private static getPermissionTypes(url: string, requestDigest: string, customPermission: boolean = false): PromiseLike<{ [key: number]: number }> {
        // Return a promise
        return new Promise(resolve => {
            // Get the definitions
            Web(url, { requestDigest }).RoleDefinitions().execute(roleDefs => {
                let roles = {};

                // Parse the role definitions
                for (let i = 0; i < roleDefs.results.length; i++) {
                    let roleDef = roleDefs.results[i];

                    // Add the role by type
                    roles[roleDef.RoleTypeKind > 0 ? roleDef.RoleTypeKind : roleDef.Name] = roleDef.Id;
                }

                // See if we are including the custom permission
                if (customPermission) {
                    // Get the custom permission level
                    this.getCustomPermissionLevel(url, roles).then(roles => {
                        // Resolve the request
                        resolve(roles);
                    });
                } else {
                    // Resolve the request
                    resolve(roles);
                }
            });
        });
    }

    // Resets the list permissions
    private static resetListPermissions(lists: string[]): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            Helper.Executor(lists, listName => {
                // Return a promise
                return new Promise(resolve => {
                    // Get the list
                    let list = Web().Lists(listName);

                    // Reset the permissions
                    list.resetRoleInheritance().execute();

                    // Clear the permissions
                    list.breakRoleInheritance(false, true).execute(true);

                    // Wait for the requests to complete
                    list.done(resolve);
                });
            }).then(resolve);
        });
    }


    // Gets the web's default site groups
    private static getWebSiteGroups(url: string, requestDigest: string): PromiseLike<{ members: Types.SP.Group, owners: Types.SP.Group, visitors: Types.SP.Group }> {
        let members: Types.SP.Group = null;
        let owners: Types.SP.Group = null;
        let visitors: Types.SP.Group = null;

        // Return a promise
        return new Promise((resolve) => {
            let web = Web(url, { requestDigest });

            // Get the default members group
            web.AssociatedMemberGroup().execute(group => {
                // Set the members group
                members = group;
            });

            // Get the default owners group
            web.AssociatedOwnerGroup().execute(group => {
                // Set the owners group
                owners = group;
            });

            // Get the default visitors group
            web.AssociatedVisitorGroup().execute(group => {
                // Set the visitors group
                visitors = group;
            });

            // Wait for the requests to complete
            web.done(() => {
                // Resolve the request
                resolve({ members, owners, visitors });
            });
        });
    }
}