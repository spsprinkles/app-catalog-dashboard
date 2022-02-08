import { ContextInfo, Helper, List, SPTypes, Types, Web } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * SharePoint Assets
 */
export const Configuration = Helper.SPConfig({
    // Lists
    ListCfg: [
        {
            ListInformation: {
                AllowContentTypes: true,
                BaseTemplate: SPTypes.ListTemplateType.DocumentLibrary,
                Description: "Used by developers to submit their solutions to the App Catalog",
                ContentTypesEnabled: true,
                DisableGridEditing: true,
                EnableFolderCreation: false,
                ExcludeFromOfflineClient: true,
                ImageUrl: "/_layouts/15/images/itappcatalog.png?rev=47",
                ListExperienceOptions: SPTypes.ListExperienceOptions.ClassicExperience,
                NoCrawl: true,
                Title: Strings.Lists.Apps
            },
            ContentTypes: [
                {
                    Name: "Document",
                    FieldRefs: [
                        "FileLeafRef",
                        "Title"
                    ]
                },
                {
                    Name: "App",
                    Description: "Displayed in the ribbon new item text.",
                    ParentName: "Document Set",
                    FieldRefs: [
                        "FileLeafRef",
                        { Name: "AppProductID", ReadOnly: true },
                        { Name: "AppVersion", ReadOnly: true },
                        "DevAppStatus",
                        "SharePointAppCategory",
                        "AppPublisher",
                        "AppShortDescription",
                        "AppDescription",
                        "Owners",
                        "Sponsor",
                        "AppSupportURL",
                        "AppThumbnailURL",
                        "AppImageURL1",
                        "AppImageURL2",
                        "AppImageURL3",
                        "AppImageURL4",
                        "AppImageURL5",
                        "AppVideoURL",
                        "IsAppPackageEnabled",
                        { Name: "IsDefaultAppMetadataLocale", ReadOnly: true }
                    ],
                    onCreated: () => {
                        // Get the list document set home page
                        List(Strings.Lists.Apps).RootFolder().Folders("Forms").Folders("App").Files("docsethomepage.aspx").execute(page => {
                            // Add the dashboard webpart
                            Helper.addContentEditorWebPart(page.ServerRelativeUrl, {
                                title: "Dashboard",
                                description: "Displays the custom dashboard.",
                                frameType: "None",
                                index: 0,
                                zone: "WebPartZone_Top",
                                contentLink: Strings.SolutionUrl
                            });
                        });
                    }
                }
            ],
            CustomFields: [
                {
                    name: "AppProductID",
                    title: "Product ID",
                    type: Helper.SPCfgFieldType.Guid,
                    allowDeletion: false,
                    showInNewForm: false
                },
                {
                    name: "AppVersion",
                    title: "App Version",
                    type: Helper.SPCfgFieldType.Text,
                    allowDeletion: false,
                    showInNewForm: false
                },
                {
                    name: "IsDefaultAppMetadataLocale",
                    title: "Default Metadata Language",
                    type: Helper.SPCfgFieldType.Boolean,
                    allowDeletion: false,
                    defaultValue: "1",
                    showInNewForm: false
                },
                {
                    name: "AppShortDescription",
                    title: "Short Description",
                    type: Helper.SPCfgFieldType.Text,
                    allowDeletion: false,
                    required: true
                },
                {
                    name: "AppDescription",
                    title: "Description",
                    type: Helper.SPCfgFieldType.Note,
                    allowDeletion: false,
                    required: true
                },
                {
                    name: "AppThumbnailURL",
                    title: "Icon URL",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false,
                    description: "The URL to the app icon. The icon should have a width and height of 96 pixels.",
                    format: SPTypes.UrlFormatType.Image,
                    required: true
                } as Helper.IFieldInfoUrl,
                {
                    name: "SharePointAppCategory",
                    title: "Category",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    fillInChoice: true
                } as Helper.IFieldInfoChoice,
                {
                    name: "Sponsor",
                    title: "Sponsor",
                    type: Helper.SPCfgFieldType.User,
                    allowDeletion: false,
                    description: "What government poc is responsible for the application?",
                    enforceUniqueValues: false,
                    required: true,
                    selectionMode: SPTypes.FieldUserSelectionType.PeopleAndGroups,
                    showField: "ImnName",
                    sortable: false
                } as Helper.IFieldInfoUser,
                {
                    name: "Owners",
                    title: "Owners",
                    type: Helper.SPCfgFieldType.User,
                    allowDeletion: false,
                    description: "Who is allowed to make updates to this app?",
                    enforceUniqueValues: false,
                    multi: true,
                    required: true,
                    selectionMode: SPTypes.FieldUserSelectionType.PeopleAndGroups,
                    showField: "ImnName",
                    sortable: false
                } as Helper.IFieldInfoUser,
                {
                    name: "AppPublisher",
                    title: "Publisher",
                    type: Helper.SPCfgFieldType.Text,
                    allowDeletion: false,
                    required: true
                },
                {
                    name: "AppSupportURL",
                    title: "Support URL",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false,
                    required: true
                } as Helper.IFieldInfoUrl,
                {
                    name: "AppImageURL1",
                    title: "Image URL 1",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false,
                    description: "An image or screenshot for this app. Images should have a width of 512 pixes and a height of 384 pixels.",
                    format: SPTypes.UrlFormatType.Image,
                    required: true
                } as Helper.IFieldInfoUrl,
                {
                    name: "AppImageURL2",
                    title: "Image URL 2",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false,
                    description: "An image or screenshot for this app. Images should have a width of 512 pixes and a height of 384 pixels.",
                    format: SPTypes.UrlFormatType.Image,
                    required: true
                } as Helper.IFieldInfoUrl,
                {
                    name: "AppImageURL3",
                    title: "Image URL 3",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false,
                    description: "An image or screenshot for this app. Images should have a width of 512 pixes and a height of 384 pixels.",
                    format: SPTypes.UrlFormatType.Image,
                } as Helper.IFieldInfoUrl,
                {
                    name: "AppImageURL4",
                    title: "Image URL 4",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false,
                    description: "An image or screenshot for this app. Images should have a width of 512 pixes and a height of 384 pixels.",
                    format: SPTypes.UrlFormatType.Image
                } as Helper.IFieldInfoUrl,
                {
                    name: "AppImageURL5",
                    title: "Image URL 5",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false,
                    description: "An image or screenshot for this app. Images should have a width of 512 pixes and a height of 384 pixels.",
                    format: SPTypes.UrlFormatType.Image,
                } as Helper.IFieldInfoUrl,
                {
                    name: "AppVideoURL",
                    title: "Video URL",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false,
                    description: "A URL to a video file or web page with an embedded video."
                } as Helper.IFieldInfoUrl,
                {
                    name: "IsAppPackageEnabled",
                    title: "Enabled",
                    type: Helper.SPCfgFieldType.Boolean,
                    allowDeletion: false,
                    defaultValue: "1"
                },
                {
                    name: "AppComments",
                    title: "Comments",
                    type: Helper.SPCfgFieldType.Note,
                    allowDeletion: false,
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    showInEditForm: false,
                    showInNewForm: false,
                    showInViewForms: false
                } as Helper.IFieldInfoNote,
                {
                    name: "DevAppStatus",
                    title: "App Status",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    defaultValue: "Missing Metadata",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    required: true,
                    showInEditForm: false,
                    showInNewForm: false,
                    choices: [
                        "Draft", "Submitted for Review", "In Review", "Requesting Approval", "Requires Attention", "Approved"
                    ]
                } as Helper.IFieldInfoChoice
            ],
            ViewInformation: [{
                ViewName: "All Documents",
                ViewFields: ["DocIcon", "LinkFilename", "Modified", "Editor"],
                ViewQuery: '<OrderBy><FieldRef Name="FileLeafRef" Ascending="TRUE" /></OrderBy>'
            }]
        },
        {
            ListInformation: {
                BaseTemplate: SPTypes.ListTemplateType.GenericList,
                Title: Strings.Lists.Assessments,
                Description: "Used by developers to submit their assessments of other apps",
                ListExperienceOptions: 2,
                AllowContentTypes: true,
                ContentTypesEnabled: false,
                DisableGridEditing: true,
                EnableFolderCreation: false,
                ExcludeFromOfflineClient: true,
                NoCrawl: true,
                EnableAttachments: false,
                ReadSecurity: 1,
                WriteSecurity: 2,
                ImageUrl: "/_layouts/15/images/itcommcat.png?rev=47",
                EnableVersioning: true,
                MajorVersionLimit: 50
            },
            CustomFields: [
                {
                    name: "SupportUrlValid",
                    title: "Support URL",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "Does the &quot;Support URL&quot; have appropriate information to aid users in resolving problems or contacting the app's developers?",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    required: true,
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "IconImages",
                    title: "Icon/Images",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "Are the app's icon and screenshots/images valid and sufficient?",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    required: true,
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "DescriptionText",
                    title: "Description text",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "Are the app's description text fields valid and sufficient to describe the app to users?",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    required: true,
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "AppTested",
                    title: "App tested",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "Have you tested the app and can you confirm it's sufficiently functional?",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    required: true,
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "UnresolvedIssues",
                    title: "Unresolved Issues?",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "Are there any unresolved issues with the app that would cause you to not recommend the app be published to the tenant app catalog for all users to see?",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    required: true,
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "Comments",
                    title: "Comments",
                    type: Helper.SPCfgFieldType.Note,
                    allowDeletion: false,
                    description: "Enter any applicable comments you have for the app developer based on your responses above.",
                    noteType: SPTypes.FieldNoteType.RichText,
                    numberOfLines: 4,
                    required: true,
                    sortable: false
                } as Helper.IFieldInfoNote,
                {
                    name: "Completed",
                    title: "Date Completed",
                    type: Helper.SPCfgFieldType.Date,
                    allowDeletion: false,
                    description: "The assessment completion date.",
                    indexed: true,
                    displayFormat: SPTypes.DateFormat.DateTime,
                    showInEditForm: false,
                    showInNewForm: false,
                    showInViewForms: false
                } as Helper.IFieldInfoDate,
                {
                    name: "RelatedApp",
                    title: "Related App",
                    type: Helper.SPCfgFieldType.Lookup,
                    allowDeletion: false,
                    indexed: true,
                    listName: Strings.Lists.Apps,
                    relationshipBehavior: SPTypes.RelationshipDeleteBehaviorType.Cascade,
                    required: true,
                    showField: "Title",
                    showInEditForm: false,
                    showInNewForm: false,
                    showInViewForms: false
                } as Helper.IFieldInfoLookup,
                {
                    name: "ResultResponse",
                    title: "Result Response",
                    type: Helper.SPCfgFieldType.Calculated,
                    allowDeletion: false,
                    format: SPTypes.DateFormat.DateTime,
                    formula: '=IF([Support URL]&amp;[Icon/Images]&amp;[Description text]&amp;[App tested]&amp;[Unresolved Issues?]="YesYesYesYesNo","Recommend","Oppose")',
                    lcid: 1033,
                    resultType: SPTypes.FieldResultType.Text,
                    fieldRefs: [
                        "UnresolvedIssues", "AppTested", "DescriptionText", "IconImages", "SupportUrlValid"
                    ]
                } as Helper.IFieldInfoCalculated
            ],
            ViewInformation: [{
                ViewName: "All Items",
                ViewFields: ["Title", "SupportUrlValid", "IconImages", "DescriptionText", "AppTested", "UnresolvedIssues", "ResultResponse", "Editor"],
                ViewQuery: '<OrderBy><FieldRef Name="ID" Ascending="FALSE" /></OrderBy>'
            }]
        }
    ]
});

// Adds the solution to a classic page
Configuration["addToPage"] = (pageUrl: string) => {
    // Add a content editor webpart to the page
    Helper.addContentEditorWebPart(pageUrl, {
        contentLink: Strings.SolutionUrl,
        description: Strings.ProjectDescription,
        frameType: "None",
        title: Strings.ProjectName
    }).then(
        // Success
        () => {
            // Load
            console.log("[" + Strings.ProjectName + "] Successfully added the solution to the page.", pageUrl);
        },

        // Error
        ex => {
            // Load
            console.log("[" + Strings.ProjectName + "] Error adding the solution to the page.", ex);
        }
    );
}

// Creates the security groups
export const createSecurityGroups = (): PromiseLike<void> => {
    // Return a promise
    return new Promise((resolve, reject) => {
        let approveGroup: Types.SP.Group = null;
        let devGroup: Types.SP.Group = null;
        let web = Web();

        // Parse the groups to create
        Helper.Executor([Strings.Groups.Approvers, Strings.Groups.Developers], groupName => {
            // Return a promise
            return new Promise((resolve, reject) => {
                // Get the group
                web.SiteGroups().getByName(groupName).execute(
                    // Exists
                    group => {
                        // See if this is the approver group
                        if (group.Title == Strings.Groups.Approvers) {
                            // Set the approver group
                            approveGroup = group;
                        } else {
                            // Set the dev group
                            devGroup = group;
                        }

                        // Resolve the request
                        resolve(null);
                    },

                    // Doesn't exist
                    () => {
                        let isDevGroup = groupName == Strings.Groups.Developers;

                        // Create the group
                        Web().SiteGroups().add({
                            AllowMembersEditMembership: true,
                            AllowRequestToJoinLeave: isDevGroup,
                            AutoAcceptRequestToJoinLeave: isDevGroup,
                            Title: groupName,
                            Description: "Group contains the '" + groupName + "' users.",
                            OnlyAllowMembersViewMembership: false
                        }).execute(
                            // Successful
                            group => {
                                // See if this is the approver group
                                if (isDevGroup) {
                                    // Set the dev group
                                    devGroup = group;
                                } else {
                                    // Set the approver group
                                    approveGroup = group;
                                }

                                // Resolve the request
                                resolve(null);
                            },

                            // Error
                            () => {
                                // The user is probably not an admin
                                console.error("Error creating the security group.");

                                // Reject the request
                                reject();
                            }
                        );
                    }
                );
            });
        }).then(() => {
            // Gets the role definitions for the permission types
            let getPermissionTypes = () => {
                // Return a promise
                return new Promise(resolve => {
                    // Get the definitions
                    Web().RoleDefinitions().execute(roleDefs => {
                        let roles = {};

                        // Parse the role definitions
                        for (let i = 0; i < roleDefs.results.length; i++) {
                            let roleDef = roleDefs.results[i];

                            // Add the role by type
                            roles[roleDef.RoleTypeKind] = roleDef.Id;
                        }

                        // Resolve the request
                        resolve(roles);
                    });
                });
            }

            // Clears the security groups for a list
            let resetListPermissions = () => {
                // Return a promise
                return new Promise(resolve => {
                    Helper.Executor([Strings.Lists.Apps, Strings.Lists.Assessments], listName => {
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

            // Update the group owners
            let updateOwners = () => {
                // Return a promise
                return new Promise((resolve, reject) => {
                    // Execute against the groups
                    Helper.Executor([devGroup], group => {
                        // Set the site owner
                        return Helper.setGroupOwner(group.Title, approveGroup.Title);
                    }).then(resolve, reject);
                });
            }

            // Update the group owners
            updateOwners().then(() => {
                // Reset the list permissions
                resetListPermissions().then(() => {
                    // Get the definitions
                    getPermissionTypes().then(permissions => {
                        // Get the lists to update
                        let listApps = Web().Lists(Strings.Lists.Apps);
                        let listAssessments = Web().Lists(Strings.Lists.Assessments);

                        // Ensure the approver group exists
                        if (approveGroup) {
                            // Set the list permissions
                            listApps.RoleAssignments().addRoleAssignment(approveGroup.Id, permissions[SPTypes.RoleType.WebDesigner]).execute(() => {
                                // Log
                                console.log("[Apps List] The approver permission was added successfully.");
                            });
                            listAssessments.RoleAssignments().addRoleAssignment(approveGroup.Id, permissions[SPTypes.RoleType.WebDesigner]).execute(() => {
                                // Log
                                console.log("[Assessments List] The approver permission was added successfully.");
                            });
                        }

                        // Ensure the dev group exists
                        if (devGroup) {
                            // Set the list permissions
                            listApps.RoleAssignments().addRoleAssignment(devGroup.Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                                // Log
                                console.log("[Apps List] The dev permission was added successfully.");
                            });
                            listAssessments.RoleAssignments().addRoleAssignment(devGroup.Id, permissions[SPTypes.RoleType.Contributor]).execute(() => {
                                // Log
                                console.log("[Assessments List] The dev permission was added successfully.");
                            });
                        }

                        // Wait for the app list updates to complete
                        listApps.done(() => {
                            // Wait for the assessment list updates to complete
                            listAssessments.done(() => {
                                // Resolve the request
                                resolve();
                            });
                        });
                    });
                });
            });
        });
    });
}
