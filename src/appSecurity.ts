import { ContextInfo, Helper, SPTypes, Web } from "gd-sprest-bs";
import { AppConfig } from "./appCfg";
import { AppSecurityWeb } from "./appSecurityWeb";
import { ErrorDialog } from "./errorDialog";
import Strings from "./strings";

/**
 * Application Security
 * The security group/user references for the application.
 */
export class AppSecurity {
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

    // Adds the developer to the security group
    static addDeveloper(): PromiseLike<any> {
        // Return a promise
        return new Promise(resolve => {
            let requests = [];

            // See if the user is not in the developer's group
            if (!AppSecurity.AppWeb.isInGroup(Strings.Groups.Developers)) {
                // Log
                ErrorDialog.logInfo(`App developer not in app web group. Adding the user...`);

                // Add the user to the group
                requests.push(AppSecurity.AppWeb.addUserToGroup(Strings.Groups.Developers, this.AppWeb.CurrentUser.Id));
            }

            // See if the app catalog is on a different site as the app
            if (AppConfig.Configuration.appCatalogUrl.toLowerCase() != Strings.SourceUrl.toLowerCase()) {
                // Log
                ErrorDialog.logInfo(`App developer not in app catalog web group. Adding the user...`);

                // Add the sponsor to the group
                requests.push(AppSecurity.AppWeb.addUserToGroup(Strings.Groups.Developers, this.AppCatalogWeb.CurrentUser.Id));
            }

            // Wait for the requests to complete and resolve the request
            Promise.all(requests).then(resolve);
        });
    }

    // Adds the sponsor to the security group
    static addSponsor(userId: number = -1): PromiseLike<any> {
        // Return a promise
        return new Promise(resolve => {
            let requests = [];

            // Ensure the user id exists
            if (userId < 1) { resolve(null); return; }

            // See if the user is not part of the group
            if (AppSecurity.AppWeb.getUserForGroup(Strings.Groups.Sponsors, userId) == null) {
                // Log
                ErrorDialog.logInfo(`App sponsor not in app web group. Adding the user...`);

                // Add the sponsor to the group
                requests.push(AppSecurity.AppWeb.addUserToGroup(Strings.Groups.Sponsors, userId));
            }

            // See if the app catalog is on the same site as the app
            if (AppConfig.Configuration.appCatalogUrl.toLowerCase() != Strings.SourceUrl.toLowerCase()) {
                // Log
                ErrorDialog.logInfo(`App sponsor not in app catalog web group. Adding the user...`);

                // Add a request
                requests.push(new Promise(resolve => {
                    // Load the user information
                    Web(Strings.SourceUrl).getUserById(userId).execute(user => {
                        // Get the context information of the other web
                        ContextInfo.getWeb(AppConfig.Configuration.appCatalogUrl).execute(context => {
                            // Ensure the user exists in the other web
                            Web(AppConfig.Configuration.appCatalogUrl, { requestDigest: context.GetContextWebInformation.FormDigestValue }).ensureUser(user.LoginName).execute(user => {
                                // Add the sponsor to the group
                                AppSecurity.AppCatalogWeb.addUserToGroup(Strings.Groups.Sponsors, user.Id).then(resolve);
                            }, resolve);
                        }, resolve);
                    }, resolve)
                }));
            }

            // Wait for the requests to complete and resolve the request
            Promise.all(requests).then(resolve);
        });
    }

    // Configures the app site
    static configureWeb(webUrl: string = Strings.SourceUrl): PromiseLike<void> {
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
                    let group = this._appCatalogWeb.getGroup(Strings.Groups.Approvers);
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
                    group = this._appCatalogWeb.getGroup(Strings.Groups.Developers);
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
                    group = this._appCatalogWeb.getGroup(Strings.Groups.FinalApprovers);
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
                    group = this._appCatalogWeb.getGroup(Strings.Groups.Sponsors);
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

    static hasErrors() {
        // Ensure the webs exist
        if (this.AppCatalogWeb == null || this.AppWeb == null) {
            // Sites are missing
            return true;
        }

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
            if (AppConfig.Configuration.appCatalogUrl.toLowerCase() != Strings.SourceUrl.toLowerCase()) {
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
    static createAppSecurityGroups(): PromiseLike<void> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let onComplete = () => {
                // Create the security groups for the app web
                this.AppWeb.create().then(() => {
                    // See if the app catalog is on a different web
                    if (AppConfig.Configuration && AppConfig.Configuration.appCatalogUrl != Strings.SourceUrl) {
                        // Create the security groups for the app catalog web
                        this.AppCatalogWeb.create().then(resolve, reject);
                    } else {
                        // Resolve the request
                        resolve();
                    }
                }, reject);
            }

            // Iniitialize the web
            this.init().then(onComplete, () => {
                // Wait for the app web to be initialized
                this._appWeb.isInitialized().then(onComplete, onComplete);
            });
        });
    }
}