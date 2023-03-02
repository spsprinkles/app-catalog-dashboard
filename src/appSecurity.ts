import { ContextInfo, Helper, SPTypes, Types, Web } from "gd-sprest-bs";
import { AppConfig } from "./appCfg";
import { ErrorDialog } from "./errorDialog";
import Strings, { setContext } from "./strings";

// App Groups
interface IAppGroups {
    Approvers: Types.SP.GroupOData;
    Developers: Types.SP.GroupOData;
    FinalApprovers: Types.SP.GroupOData;
    HelpDesk?: Types.SP.GroupOData;
    Sponsors: Types.SP.GroupOData;
}

// Web Groups
interface IWebGroups {
    members: Types.SP.GroupOData;
    owners: Types.SP.GroupOData;
    visitors: Types.SP.GroupOData;
}

/**
 * Application Security Web
 * The site groups for a web.
 */
class AppSecurityWeb {
    // Constructor
    constructor(webUrl: string, useCustomPermission: boolean = false) {
        // Set the properties
        this._customPermissionInUse = useCustomPermission;
        this._url = webUrl;

        // Initialize the class
        this.init();
    }

    /** Public Methods */

    // Add the user to the developer group
    addUserToGroup(groupName: string, userId: number): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the group
            let group = this.getGroup(groupName);
            if (group) {
                // Add the user
                group.Users.addUserById(userId).execute(
                    user => {
                        // Append the user
                        group.Users.results.push(user);

                        // Resolve the request
                        resolve();
                    },
                    ex => {
                        // Log the error
                        ErrorDialog.logError(`There was an error adding the user with id ${userId} to the ${groupName} security group.`);

                        // Reject the request
                        reject(ex);
                    }
                );
            }
        });
    }

    // Creates the app groups
    create(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let web = Web(this._url, { requestDigest: this._context.FormDigestValue });

            // Parse the group names
            let groupNames = this.getAppGroupNames();
            for (let i = 0; i < groupNames.length; i++) {
                let groupName = groupNames[i];

                // See if this group exists
                let group = this.getGroup(groupName);
                if (group) { continue; }

                // Create the group
                this.createAppGroup(web, groupName);
            }

            // Wait for the requests to complete
            web.done(() => {
                // Get the role types
                this.createCustomPermissionLevel().then(resolve, reject);
            });
        });
    }

    // Current user
    get CurrentUser(): Types.SP.User { return this._user; }

    // Returns the emails for a group
    getEmailsForGroup(groupName: string): string[] {
        let emails = [];

        // Get the group
        let group = this.getGroup(groupName);
        if (group) {
            // Parse the approvers
            for (let i = 0; i < group.Users.results.length; i++) {
                let email = group.Users.results[i].Email;

                // Append the email
                email ? emails.push(email) : null;
            }
        }

        // Return the emails
        return emails;
    }

    // Returns the group
    getGroup(groupName: string): Types.SP.GroupOData { return this._appGroups[groupName]; }

    // Returns a user from a group
    getUserForGroup(groupName: string, userId: number): Types.SP.User {
        // Get the group
        let group = this.getGroup(groupName);
        if (group) {
            // Parse the approvers
            for (let i = 0; i < group.Users.results.length; i++) {
                let user = group.Users.results[i];

                // Return the user, if it matches
                if (user.Id == userId) { return user; }
            }
        }
    }

    // Returns the url of the security group
    getUrlForGroup(groupName: string) {
        let group = this.getGroup(groupName);
        if (group) {
            // Return the url
            return this._url + "/_layouts/15/people.aspx?MembershipGroupId=" + group.Id;
        }
    }

    // Gets the role definitions for the permission types
    getPermissionTypes(): PromiseLike<{ [key: number]: number }> {
        // Return a promise
        return new Promise(resolve => {
            // Get the definitions
            Web(this._url, { requestDigest: this._context.FormDigestValue }).RoleDefinitions().execute(roleDefs => {
                let roles = {};

                // Parse the role definitions
                for (let i = 0; i < roleDefs.results.length; i++) {
                    let roleDef = roleDefs.results[i];

                    // Add the role by type
                    roles[roleDef.RoleTypeKind > 0 ? roleDef.RoleTypeKind : roleDef.Name] = roleDef.Id;
                }

                // Resolve the request
                resolve(roles);
            });
        });
    }

    // Returns true if the user is a SCA
    get IsAdmin(): boolean {
        // See if the user is an admin
        return this._user.IsSiteAdmin;
    }

    // Returns true if the user is in the group
    isInGroup(groupName: string): boolean {
        let userInGroup = false;

        // Get the group
        let group = this.getGroup(groupName);
        if (group) {
            // Parse the users
            for (let i = 0; i < group.Users.results.length; i++) {
                let user = group.Users.results[i];

                // See if this is the user
                if (user.Id == this._user.Id) {
                    // Set the flag and break from the loop
                    userInGroup = true;
                    break;
                }
            }
        }

        // Return the flag
        return userInGroup;
    }

    // Returns true if the user has full rights to the site
    get IsOwner(): boolean {
        // See if the user is an owner
        return Helper.hasPermissions(this._userPermissions, [SPTypes.BasePermissionTypes.FullMask]);
    }

    // Waits for the class to be initialized
    isInitialized(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let counter = 0;
            let maxTries = 100;

            // Wait 50 ms
            let loopId = setInterval(() => {
                // See if the class is inititalized
                if (this._initFl) {
                    // Stop the loop
                    clearInterval(loopId);

                    // Resolve the request
                    resolve();
                }

                // See if there was an error or we have maxed out attempts
                if (this._errorFl || ++counter > maxTries) {
                    // Stop the loop
                    clearInterval(loopId);

                    // Reject the request
                    reject();
                }
            }, 50);
        });
    }

    // Returns the web member's group
    get webMembers(): Types.SP.GroupOData { return this._webGroups.members; }

    // Returns the web owner's group
    get webOwners(): Types.SP.GroupOData { return this._webGroups.owners; }

    // Returns the web visitor's group
    get webVisitors(): Types.SP.GroupOData { return this._webGroups.visitors; }

    /** Global Variables */

    private _appGroups: IAppGroups = null;
    private _context: Types.SP.ContextWebInformation = null;
    private _customPermissionInUse: boolean = null;
    private _customPermissionLevel: string = "Contribute and Manage Subwebs";
    private _errorFl: boolean = null;
    private _initFl: boolean = null;
    private _url: string = null;
    private _user: Types.SP.User = null;
    private _userPermissions: Types.SP.BasePermissions = null;
    private _webGroups: IWebGroups = null;

    /** Private Methods */

    // Gets an app group
    private createAppGroup(web: Types.SP.IWeb, groupName: string): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Get the site group
            web.SiteGroups().getByName(groupName).query({ Expand: ["Users"] }).execute(
                // Exists
                group => {
                    // Set the group
                    this._appGroups[groupName] = group;

                    // Resolve the request
                    resolve();
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
                            // Get the group
                            web.SiteGroups(group.Id).query({ Expand: ["Users"] }).execute(group => {
                                // Set the group
                                this._appGroups[groupName] = group;

                                // Check the next group
                                resolve(null);
                            });
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
    }

    // Returns the app's site group names
    private getAppGroupNames(): string[] {
        let groupNames: string[] = [];

        // See if the help desk group is being used
        if (AppConfig.Configuration.helpdeskGroupName) {
            // Add the name
            groupNames.push(Strings.Groups[AppConfig.Configuration.helpdeskGroupName]);

            // Clear the group
            this._appGroups[AppConfig.Configuration.helpdeskGroupName] = null;
        }

        // Parse the configuration's group names
        for (let groupName in Strings.Groups) {
            // Add the name
            groupNames.push(Strings.Groups[groupName]);

            // Clear the group
            this._appGroups[groupName] = null;
        }

        // Return the group names
        return groupNames;
    }

    // Initialize the class
    private init() {
        // Defaults the flags
        this._errorFl = false;
        this._initFl = false;

        // Get the context
        ContextInfo.getWeb(this._url).execute(context => {
            // Set the context
            this._context = context.GetContextWebInformation;

            // Load the web information
            let web = Web(this._url, { requestDigest: this._context.FormDigestValue });

            // Query the base permissions
            web.query({ Expand: ["EffectiveBasePermissions"] }).execute(web => {
                // Set the base permissions
                this._userPermissions = web.EffectiveBasePermissions;
            })

            // Get the current user
            web.CurrentUser().execute(user => {
                // Set the user
                this._user = user;
            });

            // Initialize the app and web groups
            this.initAppGroups(web);
            this.initWebGroups(web);

            // Wait for the requests to complete
            web.done(() => {
                // Set the flag
                this._initFl = true;
            });
        }, () => {
            // Set the flag
            this._errorFl = true;
        });
    }

    // Initializes the app groups
    private initAppGroups(web: Types.SP.IWeb) {
        // Reset the groups
        this._appGroups = {} as any;

        // Parse the group names
        let groupNames = this.getAppGroupNames();
        for (let i = 0; i < groupNames.length; i++) {
            let groupName = groupNames[i];

            // Get the site group
            web.SiteGroups().getByName(groupName).query({ Expand: ["Users"] }).execute(group => {
                // Set the group
                this._appGroups[groupName] = group;
            }, () => {
                // Set the flag
                this._errorFl = true;
            });
        }
    }

    // Gets the custom permission level
    private createCustomPermissionLevel(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            // See if we are using a custom permission
            if (this._customPermissionInUse) {
                // Get the permissions
                this.getPermissionTypes().then(permissions => {
                    // See if the permissions already contains
                    if (permissions[this._customPermissionLevel]) {
                        // Resolve the request
                        resolve();
                    } else {
                        // Create the custom permission
                        Helper.copyPermissionLevel({
                            BasePermission: "Contribute",
                            Name: this._customPermissionLevel,
                            Description: "Extends the contribute permission level and adds the ability to create a subweb.",
                            AddPermissions: [SPTypes.BasePermissionTypes.ManageSubwebs],
                            WebUrl: this._url
                        }).then(
                            role => {
                                // Resolve the request
                                resolve();
                            },
                            ex => {
                                // Log the error
                                ErrorDialog.show("Permission Level", "There was an error creating the contribute and manage subwebs custom permission level.", ex);

                                // Reject the request
                                reject();
                            }
                        );
                    }
                });
            } else {
                // Resolve the request
                resolve();
            }
        });
    }

    // Initializes the web's default site groups
    private initWebGroups(web: Types.SP.IWeb) {
        // Clear the current value
        this._webGroups = {
            members: null,
            owners: null,
            visitors: null
        };

        // Get the default members group
        web.AssociatedMemberGroup().query({ Expand: ["Users"] }).execute(group => {
            // Set the members group
            this._webGroups.members = group;
        });

        // Get the default owners group
        web.AssociatedOwnerGroup().query({ Expand: ["Users"] }).execute(group => {
            // Set the owners group
            this._webGroups.owners = group;
        });

        // Get the default visitors group
        web.AssociatedVisitorGroup().query({ Expand: ["Users"] }).execute(group => {
            // Set the visitors group
            this._webGroups.visitors = group;
        });
    }
}

/**
 * Application Security
 * The security group/user references for the application.
 */
export class AppSecurity {
    /** Public Variables */

    // App Web
    private static _appWeb: AppSecurityWeb = null;
    static get AppWeb(): AppSecurityWeb { return this._appWeb; }

    // App Catalog Web
    private static _appCatalogWeb: AppSecurityWeb = null;
    static get AppCatalogWeb(): AppSecurityWeb { return this._appCatalogWeb; }

    // App Catalog Owner
    static get IsSiteAppCatalogOwner(): boolean { return this.AppCatalogWeb.IsOwner; }

    // Tenant App Catalog Owner
    private static _isTenantAppCatalogOwner = false;
    static get IsTenantAppCatalogOwner(): boolean { return this._isTenantAppCatalogOwner; }

    /** Private Interface */

    // Creates the security groups for the app site
    private static createAppSecurityGroups(webUrl: string = Strings.SourceUrl): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve) => {
            // Get the permission types
            this.AppWeb.getPermissionTypes().then(permissions => {
                // Reset group settings
                Promise.all([
                    // Set the dev group owner
                    Helper.setGroupOwner(Strings.Groups.Developers, Strings.Groups.Approvers, webUrl),
                    // Set the sponsor group owner
                    Helper.setGroupOwner(Strings.Groups.Sponsors, Strings.Groups.Developers, webUrl),
                    // Reset the list permissions
                    this.resetListPermissions([Strings.Lists.Apps, Strings.Lists.Assessments, Strings.Lists.AppCatalogRequests])
                ]).then(() => {
                    // Get the lists to update
                    let listApps = Web().Lists(Strings.Lists.Apps);
                    let listAssessments = Web().Lists(Strings.Lists.Assessments);
                    let listRequests = Web().Lists(Strings.Lists.AppCatalogRequests);

                    // Ensure the approver group exists
                    let group = this._appCatalogWeb.getGroup[Strings.Groups.Approvers];
                    if (group) {
                        // Set the list permissions
                        listApps.RoleAssignments().addRoleAssignment(group.Id, permissions[SPTypes.RoleType.WebDesigner]).execute(() => {
                            // Log
                            console.log("[Apps List] The approver permission was added successfully.");
                        });
                        listAssessments.RoleAssignments().addRoleAssignment(group.Id, permissions[SPTypes.RoleType.WebDesigner]).execute(() => {
                            // Log
                            console.log("[Assessments List] The approver permission was added successfully.");
                        });
                        listRequests.RoleAssignments().addRoleAssignment(group.Id, permissions[SPTypes.RoleType.WebDesigner]).execute(() => {
                            // Log
                            console.log("[App Catalog Requests List] The approver permission was added successfully.");
                        });
                    }

                    // Ensure the dev group exists
                    group = this._appCatalogWeb.getGroup[Strings.Groups.Developers];
                    if (group) {
                        // Set the list permissions
                        listApps.RoleAssignments().addRoleAssignment(group.Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                            // Log
                            console.log("[Apps List] The dev permission was added successfully.");
                        });
                        listAssessments.RoleAssignments().addRoleAssignment(group.Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                            // Log
                            console.log("[Assessments List] The dev permission was added successfully.");
                        });
                        listRequests.RoleAssignments().addRoleAssignment(group.Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                            // Log
                            console.log("[App Catalog Requests List] The dev permission was added successfully.");
                        });
                    }

                    // Ensure the final approver group exists
                    group = this._appCatalogWeb.getGroup[Strings.Groups.FinalApprovers];
                    if (group) {
                        // Set the list permissions
                        listApps.RoleAssignments().addRoleAssignment(group.Id, permissions[SPTypes.RoleType.WebDesigner]).execute(() => {
                            // Log
                            console.log("[Apps List] The approver permission was added successfully.");
                        });
                        listAssessments.RoleAssignments().addRoleAssignment(group.Id, permissions[SPTypes.RoleType.WebDesigner]).execute(() => {
                            // Log
                            console.log("[Assessments List] The approver permission was added successfully.");
                        });
                        listRequests.RoleAssignments().addRoleAssignment(group.Id, permissions[SPTypes.RoleType.WebDesigner]).execute(() => {
                            // Log
                            console.log("[App Catalog Requests List] The approver permission was added successfully.");
                        });
                    }

                    // Ensure the dev group exists
                    group = this._appCatalogWeb.getGroup[Strings.Groups.Sponsors];
                    if (group) {
                        // Set the list permissions
                        listApps.RoleAssignments().addRoleAssignment(group.Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                            // Log
                            console.log("[Apps List] The sponsor permission was added successfully.");
                        });
                        listAssessments.RoleAssignments().addRoleAssignment(group.Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                            // Log
                            console.log("[Assessments List] The sponsor permission was added successfully.");
                        });
                        listRequests.RoleAssignments().addRoleAssignment(group.Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                            // Log
                            console.log("[App Catalog Requests List] The sponsor permission was added successfully.");
                        });
                    }

                    // Ensure the default members group exists
                    if (this._appCatalogWeb.webMembers) {
                        // Set the list permissions
                        listApps.RoleAssignments().addRoleAssignment(this._appCatalogWeb.webMembers.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                            // Log
                            console.log("[Apps List] The default members permission was added successfully.");
                        });
                        listAssessments.RoleAssignments().addRoleAssignment(this._appCatalogWeb.webMembers.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                            // Log
                            console.log("[Assessments List] The default members permission was added successfully.");
                        });
                        listRequests.RoleAssignments().addRoleAssignment(this._appCatalogWeb.webMembers.Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                            // Log
                            console.log("[App Catalog Requests List] The default members permission was added successfully.");
                        });
                    }

                    // Ensure the default owners group exists
                    if (this._appCatalogWeb.webOwners) {
                        // Set the list permissions
                        listApps.RoleAssignments().addRoleAssignment(this._appCatalogWeb.webOwners.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                            // Log
                            console.log("[Apps List] The default owners permission was added successfully.");
                        });
                        listAssessments.RoleAssignments().addRoleAssignment(this._appCatalogWeb.webOwners.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                            // Log
                            console.log("[Assessments List] The default owners permission was added successfully.");
                        });
                        listRequests.RoleAssignments().addRoleAssignment(this._appCatalogWeb.webOwners.Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                            // Log
                            console.log("[App Catalog Requests List] The default owners permission was added successfully.");
                        });
                    }

                    // Ensure the default members group exists
                    if (this._appCatalogWeb.webVisitors) {
                        // Set the list permissions
                        listApps.RoleAssignments().addRoleAssignment(this._appCatalogWeb.webVisitors.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                            // Log
                            console.log("[Apps List] The default visitors permission was added successfully.");
                        });
                        listAssessments.RoleAssignments().addRoleAssignment(this._appCatalogWeb.webVisitors.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                            // Log
                            console.log("[Assessments List] The default visitors permission was added successfully.");
                        });
                        listRequests.RoleAssignments().addRoleAssignment(this._appCatalogWeb.webVisitors.Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
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
    }

    // Determines if the user is able to deploy to the tenant
    private static initTenantAppCatalogOwner(): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // See if the tenant app catalog is defined
            if (AppConfig.Configuration.tenantAppCatalogUrl) {
                // Get the user's permissions to the app catalog list
                Web(AppConfig.Configuration.tenantAppCatalogUrl).Lists(Strings.Lists.AppCatalog).query({
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

    // Resets the list permissions
    private static resetListPermissions(lists: string[]): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Parse the lists
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

    /** Public Interface */

    static hasErrors() {
        // Check the app web
        if (this.AppWeb.getGroup(Strings.Groups.Approvers) == null ||
            this.AppWeb.getGroup(Strings.Groups.Developers) == null ||
            this.AppWeb.getGroup(Strings.Groups.FinalApprovers) == null ||
            this.AppWeb.getGroup(Strings.Groups.Sponsors) == null) {
            // Groups are missing
            return true;
        }

        // Check the app catalog web
        if (this.AppCatalogWeb.getGroup(Strings.Groups.Approvers) == null ||
            this.AppCatalogWeb.getGroup(Strings.Groups.Developers) == null ||
            this.AppCatalogWeb.getGroup(Strings.Groups.FinalApprovers) == null ||
            this.AppCatalogWeb.getGroup(Strings.Groups.Sponsors) == null) {
            // Groups are missing
            return true;
        }

        // Groups exist
        return false;
    }

    static init() {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Initialize the app web
            this._appWeb = new AppSecurityWeb(Strings.SourceUrl);

            // See if the app catalog is on a different web
            if (AppConfig.Configuration.appCatalogUrl != Strings.SourceUrl) {
                // Initialize the app catalog web
                this._appCatalogWeb = new AppSecurityWeb(AppConfig.Configuration.appCatalogUrl);
            } else {
                // Set the app catalog web
                this._appCatalogWeb = this._appWeb;
            }
            // Execute the requests, and resolve the request when it completes
            Promise.all([
                // Wait for the webs to complete
                this._appWeb.isInitialized(),
                this._appCatalogWeb.isInitialized(),
                this.initTenantAppCatalogOwner(),
            ]).then(resolve, () => {
                // Log
                ErrorDialog.logError("[App Security] Error initializing the App Security component.");

                // Reject the request
                reject();
            });
        });
    }

    // Creates the security groups
    static install(): PromiseLike<any> {
        // Create the security groups
        return Promise.all([
            this.AppCatalogWeb.create(),
            this.createAppSecurityGroups()
        ]);
    }
}