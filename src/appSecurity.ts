import { Helper, SPTypes, Web } from "gd-sprest-bs";
import { AppConfig } from "./appCfg";
import { AppSecurityWeb } from "./appSecurityWeb";
import { ErrorDialog } from "./errorDialog";
import Strings from "./strings";

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

    /** Public Interface */

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