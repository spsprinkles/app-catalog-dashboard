import { Helper, List, SPTypes, Types, Web } from "gd-sprest-bs";
import { AppConfig } from "./appCfg";
import { ErrorDialog } from "./errorDialog";
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
                        { Name: "AppStatus", ReadOnly: true },
                        { Name: "AppProductID", ReadOnly: true },
                        { Name: "AppIsClientSideSolution", ReadOnly: true },
                        { Name: "AppIsDomainIsolated", ReadOnly: true },
                        { Name: "AppSharePointMinVersion", ReadOnly: true },
                        { Name: "AppSkipFeatureDeployment", ReadOnly: true },
                        { Name: "AppVersion", ReadOnly: true },
                        { Name: "AppSiteDeployments", ReadOnly: true },
                        "AppPublisher",
                        "AppSponsor",
                        "AppDevelopers",
                        "AppDescription",
                        "AppJustification",
                        "AppThumbnailURL",
                        { Name: "AppAPIPermissions", ReadOnly: true },
                        "AppPermissionsJustification",
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
                    name: "AppDescription",
                    title: "Description",
                    type: Helper.SPCfgFieldType.Note,
                    description: "A meaningful description of the application purpose and what a user can expect from it.",
                    allowDeletion: false,
                    required: true
                },
                {
                    name: "AppDevelopers",
                    title: "Developers",
                    type: Helper.SPCfgFieldType.User,
                    allowDeletion: false,
                    description: "The developers of the application.",
                    enforceUniqueValues: false,
                    multi: true,
                    required: true,
                    selectionMode: SPTypes.FieldUserSelectionType.PeopleAndGroups,
                    showField: "ImnName",
                    sortable: false
                } as Helper.IFieldInfoUser,
                {
                    name: "AppIsRejected",
                    title: "Is Rejected?",
                    type: Helper.SPCfgFieldType.Boolean,
                    allowDeletion: false,
                    defaultValue: "0",
                    showInNewForm: false
                },
                {
                    name: "AppIsTenantDeployed",
                    title: "Is Tenant Deployed?",
                    type: Helper.SPCfgFieldType.Boolean,
                    allowDeletion: false,
                    defaultValue: "0",
                    showInNewForm: false
                },
                {
                    name: "AppJustification",
                    title: "Justification",
                    type: Helper.SPCfgFieldType.Note,
                    description: "A meaningful description of the application purpose and what a user can expect from it.",
                    allowDeletion: false,
                    required: true
                },
                {
                    name: "AppPermissionsJustification",
                    title: "Permissions Justification",
                    type: Helper.SPCfgFieldType.Note,
                    description: "The justification for the Web API permissions being requested.",
                    allowDeletion: false
                },
                {
                    name: "AppPublisher",
                    title: "App Publisher",
                    type: Helper.SPCfgFieldType.Text,
                    description: "The organization/unit responsible for the application.",
                    allowDeletion: false,
                    required: true
                },
                {
                    name: "AppSiteDeployments",
                    title: "App Site Deployments",
                    type: Helper.SPCfgFieldType.Note,
                    allowDeletion: false,
                    description: "The sites where the app has been deployed to.",
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    showInEditForm: false,
                    showInNewForm: false
                } as Helper.IFieldInfoNote,
                {
                    name: "AppSponsor",
                    title: "Sponsor",
                    type: Helper.SPCfgFieldType.User,
                    allowDeletion: false,
                    description: "The sponsor of the application.",
                    required: true,
                    selectionMode: SPTypes.FieldUserSelectionType.PeopleAndGroups,
                    showField: "ImnName"
                } as Helper.IFieldInfoUser,
                {
                    name: "AppStatus",
                    title: "App Status",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    defaultValue: "New",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    required: true,
                    showInEditForm: false,
                    showInNewForm: false,
                    choices: [
                        "New", "Pending Sponsor Approval", "In Testing",
                        "Pending Peer Review", "Pending Approval", "Pending Deployment",
                        "Deployed"
                    ]
                } as Helper.IFieldInfoChoice,
                /** Fields extracted from the SPFx package */
                {
                    name: "AppAPIPermissions",
                    title: "API Permission",
                    type: Helper.SPCfgFieldType.Note,
                    description: "The Web API permission requests for the application.",
                    allowDeletion: false,
                    showInNewForm: false
                },
                {
                    name: "AppIsClientSideSolution",
                    title: "Is ClientSide Solution",
                    type: Helper.SPCfgFieldType.Boolean,
                    allowDeletion: false,
                    showInNewForm: false
                },
                {
                    name: "AppIsDomainIsolated",
                    title: "Is Domain Isolated",
                    type: Helper.SPCfgFieldType.Boolean,
                    allowDeletion: false,
                    showInNewForm: false
                },
                {
                    name: "AppProductID",
                    title: "Product ID",
                    type: Helper.SPCfgFieldType.Guid,
                    allowDeletion: false,
                    showInNewForm: false
                },
                {
                    name: "AppSharePointMinVersion",
                    title: "SharePoint Min Version",
                    type: Helper.SPCfgFieldType.Text,
                    allowDeletion: false,
                    showInNewForm: false
                },
                {
                    name: "AppSkipFeatureDeployment",
                    title: "Skip Feature Deployment",
                    type: Helper.SPCfgFieldType.Boolean,
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
                /** Fields required by the app catalog */
                {
                    name: "AppImageURL1",
                    title: "Screenshot URL 1",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false,
                    description: "An image or screenshot for this app. Images should have a width of 512 pixels and a height of 384 pixels.",
                    format: SPTypes.UrlFormatType.Image
                } as Helper.IFieldInfoUrl,
                {
                    name: "AppImageURL2",
                    title: "Screenshot URL 2",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false,
                    description: "An image or screenshot for this app. Images should have a width of 512 pixels and a height of 384 pixels.",
                    format: SPTypes.UrlFormatType.Image
                } as Helper.IFieldInfoUrl,
                {
                    name: "AppImageURL3",
                    title: "Screenshot URL 3",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false,
                    description: "An image or screenshot for this app. Images should have a width of 512 pixels and a height of 384 pixels.",
                    format: SPTypes.UrlFormatType.Image,
                } as Helper.IFieldInfoUrl,
                {
                    name: "AppImageURL4",
                    title: "Screenshot URL 4",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false,
                    description: "An image or screenshot for this app. Images should have a width of 512 pixels and a height of 384 pixels.",
                    format: SPTypes.UrlFormatType.Image
                } as Helper.IFieldInfoUrl,
                {
                    name: "AppImageURL5",
                    title: "Screenshot URL 5",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false,
                    description: "An image or screenshot for this app. Images should have a width of 512 pixels and a height of 384 pixels.",
                    format: SPTypes.UrlFormatType.Image,
                } as Helper.IFieldInfoUrl,
                {
                    name: "AppSupportURL",
                    title: "Support URL",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false
                } as Helper.IFieldInfoUrl,
                {
                    name: "AppThumbnailURL",
                    title: "Icon URL",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false,
                    description: "The URL to the app icon. The icon should have a width and height of 96 pixels.",
                    format: SPTypes.UrlFormatType.Image
                } as Helper.IFieldInfoUrl,
                {
                    name: "AppVideoURL",
                    title: "Video URL",
                    type: Helper.SPCfgFieldType.Url,
                    allowDeletion: false,
                    description: "A URL to a video file or web page with an embedded video."
                } as Helper.IFieldInfoUrl,
                /** Fields to remove */
                {
                    name: "AppShortDescription",
                    title: "Short Description",
                    type: Helper.SPCfgFieldType.Text,
                    description: "A short description of the application.",
                    allowDeletion: false,
                    required: true
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
                    name: "IsAppPackageEnabled",
                    title: "Enabled",
                    type: Helper.SPCfgFieldType.Boolean,
                    allowDeletion: false,
                    defaultValue: "1"
                }
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
                ListExperienceOptions: SPTypes.ListExperienceOptions.NewExperience,
                AllowContentTypes: true,
                ContentTypesEnabled: true,
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
            ContentTypes: [
                {
                    Name: "Item",
                    FieldRefs: [
                        "TechReview01", "TechReview02", "TechReview03", "TechReview04", "TechReview05",
                        "TechReview06", "TechReview07", "TechReview08", "TechReview09", "TechReview10",
                        /*"TechReview11", "TechReview12", "TechReview13", "TechReview14", "TechReview15",
                        "TechReview16", "TechReview17", "TechReview18", "TechReview19", "TechReview20",*/
                        "TechReviewComments", "Completed", "RelatedApp"
                    ]
                },
                {
                    Name: "TestCases",
                    FieldRefs: [
                        "TestCase01", "TestCase02", "TestCase03", "TestCase04", "TestCase05",
                        "TestCase06", "TestCase07", "TestCase08", "TestCase09", "TestCase10",
                        "TestCase11", /*"TestCase12", "TestCase13", "TestCase14", "TestCase15",
                        "TestCase16", "TestCase17", "TestCase18", "TestCase19", "TestCase20",*/
                        "TestCaseComments", "Completed", "RelatedApp"
                    ]
                }
            ],
            CustomFields: [
                /**
                 * Test Case Fields
                 */

                {
                    name: "TestCase01",
                    title: "Tested in root site?",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No", "N/A"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase02",
                    title: "Tested in non-root site?",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No", "N/A"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase03",
                    title: "Tested in Browsers?",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase04",
                    title: "Tested in a modern page?",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase05",
                    title: "Tested in a classic page?",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No", "N/A"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase06",
                    title: "Tested in Teams?",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No", "N/A"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase07",
                    title: "Valid webpart properties?",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase08",
                    title: "Validation of the links?",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase09",
                    title: "Validation of the response time?",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase10",
                    title: "Validation of the description?",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase11",
                    title: "Validation of the user interface?",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase12",
                    title: "Test Case 12",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase13",
                    title: "Test Case 13",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase14",
                    title: "Test Case 14",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase15",
                    title: "Test Case 15",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase16",
                    title: "Test Case 16",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase17",
                    title: "Test Case 17",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase18",
                    title: "Test Case 18",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase19",
                    title: "Test Case 19",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCase20",
                    title: "Test Case 20",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TestCaseComments",
                    title: "Testing Comments",
                    type: Helper.SPCfgFieldType.Note,
                    allowDeletion: false,
                    description: "The comments from testing the application.",
                    noteType: SPTypes.FieldNoteType.TextOnly
                } as Helper.IFieldInfoNote,

                /**
                 * Technical Review Fields
                 */

                {
                    name: "TechReview01",
                    title: "Support URL",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "Does the &quot;Support URL&quot; have appropriate information to aid users in resolving problems or contacting the app developers?",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview02",
                    title: "Icon/Images",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "Are the app's icon and screenshots/images valid and sufficient?",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview03",
                    title: "App Description",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "Are the app's description text fields valid and sufficient to describe the app to users?",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview04",
                    title: "App Tested",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "Has the app been tested in the client environment?",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview05",
                    title: "Unresolved Issues",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "Are there any open issues with the app that would cause it to not be recommended to be published to the tenant app catalog?",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview06",
                    title: "Graph API Permissions",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "Are there any requests to the Graph API, and is the justification valid?",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No", "N/A"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview07",
                    title: "Web API Permissions",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "Are there any requests to a custom API, and is the justification valid?",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No", "N/A"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview08",
                    title: "External Libraries",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "Have the external libraries of the application been reviewed for vulnerabilities?",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview09",
                    title: "Import Statements",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "Have the import statements been reviewed for vulnerabilities?",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview10",
                    title: "HTTP Requests",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "Have the http requests been reviewed for vulnerabilities?",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview11",
                    title: "Tech Review 11",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview12",
                    title: "Tech Review 12",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview13",
                    title: "Tech Review 13",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview14",
                    title: "Tech Review 14",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview15",
                    title: "Tech Review 15",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview16",
                    title: "Tech Review 16",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview17",
                    title: "Tech Review 17",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview18",
                    title: "Tech Review 18",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview19",
                    title: "Tech Review 19",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReview20",
                    title: "Tech Review 20",
                    type: Helper.SPCfgFieldType.Choice,
                    allowDeletion: false,
                    description: "",
                    format: SPTypes.ChoiceFormatType.Dropdown,
                    defaultValue: "",
                    choices: [
                        "Yes", "No"
                    ]
                } as Helper.IFieldInfoChoice,
                {
                    name: "TechReviewComments",
                    title: "Technical Review Comments",
                    type: Helper.SPCfgFieldType.Note,
                    allowDeletion: false,
                    description: "Enter any applicable comments you have for the app developer based on your responses above.",
                    noteType: SPTypes.FieldNoteType.TextOnly,
                    numberOfLines: 4,
                    sortable: false
                } as Helper.IFieldInfoNote,

                /** Common Fields */

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
                } as Helper.IFieldInfoLookup
            ],
            ViewInformation: [{
                ViewName: "All Items",
                ViewFields: ["RelatedApp", "LinkTitle", "Completed", "ContentType"],
                ViewQuery: '<OrderBy><FieldRef Name="Title" /></OrderBy>'
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
        let sponsorGroup: Types.SP.Group = null;
        let webMembersGroup: Types.SP.Group = null;
        let webOwnersGroup: Types.SP.Group = null;
        let webVisitorsGroup: Types.SP.Group = null;
        let web = Web();

        // Get the default members group
        web.AssociatedMemberGroup().execute(group => {
            // Set the members group
            webMembersGroup = group;
        });

        // Get the default owners group
        web.AssociatedOwnerGroup().execute(group => {
            // Set the owners group
            webOwnersGroup = group;
        });

        // Get the default visitors group
        web.AssociatedVisitorGroup().execute(group => {
            // Set the visitors group
            webVisitorsGroup = group;
        });

        // Parse the groups to create
        Helper.Executor([Strings.Groups.Approvers, Strings.Groups.Developers, Strings.Groups.Sponsors], groupName => {
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
                        }
                        // Else, see if it's the developer group
                        else if (group.Title == Strings.Groups.Developers) {
                            // Set the dev group
                            devGroup = group;
                        } else {
                            // Set the sponsor group
                            sponsorGroup = group;
                        }

                        // Resolve the request
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
                                // See if this is the approver group
                                if (group.Title == Strings.Groups.Approvers) {
                                    // Set the approver group
                                    approveGroup = group;
                                }
                                // Else, see if it's the developer group
                                else if (group.Title == Strings.Groups.Developers) {
                                    // Set the dev group
                                    devGroup = group;
                                } else {
                                    // Set the sponsor group
                                    sponsorGroup = group;
                                }

                                // Resolve the request
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
            let customPermissionLevel = "Contribute and Manage Subwebs";

            // Creates the custom permission level
            let createPermissionLevel = (roles): PromiseLike<void> => {
                // Return a promise
                return new Promise(resolve => {
                    // See if the roles contain the custom permission
                    if (roles[customPermissionLevel]) {
                        // Resolve the request
                        resolve();
                    } else {
                        // Create the custom permission
                        Helper.copyPermissionLevel({
                            BasePermission: "Contribute",
                            Name: customPermissionLevel,
                            Description: "Extends the contribute permission level and adds the ability to create a subweb.",
                            AddPermissions: [SPTypes.BasePermissionTypes.ManageSubwebs],
                            WebUrl: AppConfig.Configuration.appCatalogUrl
                        }).then(
                            role => {
                                // Update the mapper
                                roles[customPermissionLevel] = role.Id;

                                // Resolve the request
                                resolve();
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
                            roles[roleDef.RoleTypeKind > 0 ? roleDef.RoleTypeKind : roleDef.Name] = roleDef.Id;
                        }

                        // Create the custom permission level
                        createPermissionLevel(roles).then(() => {
                            // Resolve the request
                            resolve(roles);
                        });
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
                        return Helper.setGroupOwner(group.Title, approveGroup.Title, Strings.SourceUrl);
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

                        // Ensure the dev group exists
                        if (sponsorGroup) {
                            // Set the list permissions
                            listApps.RoleAssignments().addRoleAssignment(sponsorGroup.Id, permissions[customPermissionLevel]).execute(() => {
                                // Log
                                console.log("[Apps List] The sponsor permission was added successfully.");
                            });
                            listAssessments.RoleAssignments().addRoleAssignment(sponsorGroup.Id, permissions[customPermissionLevel]).execute(() => {
                                // Log
                                console.log("[Assessments List] The sponsor permission was added successfully.");
                            });
                        }

                        // Ensure the default members group exists
                        if (webMembersGroup) {
                            // Set the list permissions
                            listApps.RoleAssignments().addRoleAssignment(webMembersGroup.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                                // Log
                                console.log("[Apps List] The default members permission was added successfully.");
                            });
                            listAssessments.RoleAssignments().addRoleAssignment(webMembersGroup.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                                // Log
                                console.log("[Assessments List] The default members permission was added successfully.");
                            });
                        }

                        // Ensure the default owners group exists
                        if (webOwnersGroup) {
                            // Set the list permissions
                            listApps.RoleAssignments().addRoleAssignment(webOwnersGroup.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                                // Log
                                console.log("[Apps List] The default owners permission was added successfully.");
                            });
                            listAssessments.RoleAssignments().addRoleAssignment(webOwnersGroup.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                                // Log
                                console.log("[Assessments List] The default owners permission was added successfully.");
                            });
                        }

                        // Ensure the default members group exists
                        if (webVisitorsGroup) {
                            // Set the list permissions
                            listApps.RoleAssignments().addRoleAssignment(webVisitorsGroup.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                                // Log
                                console.log("[Apps List] The default visitors permission was added successfully.");
                            });
                            listAssessments.RoleAssignments().addRoleAssignment(webVisitorsGroup.Id, permissions[SPTypes.RoleType.Reader]).execute(() => {
                                // Log
                                console.log("[Assessments List] The default visitors permission was added successfully.");
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