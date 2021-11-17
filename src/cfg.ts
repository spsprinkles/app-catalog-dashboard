import { ContextInfo, Helper, SPTypes } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * SharePoint Assets
 */
export const Configuration = Helper.SPConfig({
    // Lists
    ListCfg: [
        {
            ListInformation: {
                BaseTemplate: SPTypes.ListTemplateType.DocumentLibrary,
                Title: Strings.Lists.Apps,
                Description: "Used by developers to submit their solutions to the App Catalog",
                ListExperienceOptions: SPTypes.ListExperienceOptions.ClassicExperience,
                AllowContentTypes: true,
                ContentTypesEnabled: false,
                DisableGridEditing: true,
                EnableFolderCreation: false,
                ExcludeFromOfflineClient: true,
                NoCrawl: true,
                ImageUrl: "/_layouts/15/images/itappcatalog.png?rev=47",
            },
            //Only runs when asset is actually created, not on subsequent runs
            onCreated: function () {
                //Add scriptEditor
                Helper.addScriptEditorWebPart(ContextInfo.webServerRelativeUrl + "/DeveloperApps/Forms/AllItems.aspx", {
                    title: "Access Notice",
                    description: "Force page to display in classic mode and shows message",
                    chromeType: "None",
                    index: 0,
                    zone: "Main",
                    content: 'This page is for admin use only. Access the <a href="../../SitePages/AppDashboard.aspx">App Dashboard</a> for more features.'
                });
            },
            CustomFields: [
                {
                    name: "AppProductID",
                    title: "Product ID",
                    type: Helper.SPCfgFieldType.Guid,
                    allowDeletion: false,
                    readOnly: true,
                    showInEditForm: false,
                    showInNewForm: false
                },
                {
                    name: "AppVersion",
                    title: "App Version",
                    type: Helper.SPCfgFieldType.Text,
                    allowDeletion: false,
                    showInEditForm: false,
                    showInNewForm: false
                },
                {
                    name: "IsDefaultAppMetadataLocale",
                    title: "Default Metadata Language",
                    type: Helper.SPCfgFieldType.Boolean,
                    allowDeletion: false,
                    defaultValue: "1",
                    showInEditForm: false,
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
                    jslink: Strings.SolutionAppsCSRUrl,
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
                    name: "DevAppStatus",
                    title: "App Status",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    defaultValue: "Missing Metadata",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    required: true,
                    choices: [
                        "Missing Metadata", "Deleting...", "In Testing", "In Review", "Published"
                    ]
                } as Helper.IFieldInfoChoice,
            ],
            ViewInformation: [{
                ViewName: "All Documents",
                ViewFields: ["DocIcon", "LinkFilename", "AppPublisher", "DevAppStatus", "Modified", "Editor"],
                ViewQuery: '<OrderBy>"FileLeafRef" Ascending="TRUE" /></OrderBy>',
                JSLink: Strings.SolutionAppsCSRUrl,
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
                    name: "RelatedApp",
                    title: "Related App",
                    type: Helper.SPCfgFieldType.Lookup,
                    allowDeletion: false,
                    indexed: true,
                    jslink: Strings.SolutionAppAssessmentsCSRUrl,
                    listName: "Developer Apps",
                    relationshipBehavior: SPTypes.RelationshipDeleteBehaviorType.Cascade,
                    required: true,
                    showField: "Title"
                } as Helper.IFieldInfoLookup,
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
                ViewFields: ["Title", "RelatedApp", "SupportUrlValid", "IconImages", "DescriptionText", "AppTested", "UnresolvedIssues", "ResultResponse", "Editor"],
                ViewQuery: '<OrderBy>"ID" Ascending="FALSE" /></OrderBy>'
            }]
        }
        /*
        {
            //List should already be created via the solution package...
            ListInformation: {
                BaseTemplate: SPTypes.ListTemplateType.TasksWithTimelineAndHierarchy,
                Title: "Workflow Tasks",
                //Description: "Used by developers to submit their assessments of other apps",
            },
            //...just need to add this existing content type to the list
            ContentTypes: [{
                Name: "Workflow Task (SharePoint 2013)",
                Description: "Create a SharePoint 2013 Workflow Task",
                ParentName: "Workflow Task (SharePoint 2013)"
            }]
        }
        */
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