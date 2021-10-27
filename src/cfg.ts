import { ContextInfo, Helper, SPTypes } from "gd-sprest-bs";
import Strings from "./strings";

var devAppsCSR = "~site/code/devAppsCSR.js?v=0";
var devAppAssessmentsCSR = "~site/code/devAppAssessmentsCSR.js?v=0";

/**
 * SharePoint Assets
 */
export const Configuration = Helper.SPConfig({
    // SharePoint Assets Defined Here
    // ContentTypes, Custom Actions, Fields, Lists, WebParts
    // List Configuration
    ListCfg: [{
        ListInformation: {
            BaseTemplate: SPTypes.ListTemplateType.DocumentLibrary,
            Title: Strings.Lists.Apps,
            Description: "Used by developers to submit their solutions to the App Catalog",
            ListExperienceOptions: 2,
            //AllowContentTypes: true,
            //ContentTypesEnabled: false,
            //DisableGridEditing: true,
            EnableFolderCreation: false,
            ExcludeFromOfflineClient: true,
            NoCrawl: true,
            ImageUrl: "/_layouts/15/images/itappcatalog.png?rev=47",
        },
        //Only runs when asset is actually created, not on subsequent runs
        onCreated: function() {
            //Add ScriptEditor
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
                schemaXml: '<Field Type="Guid" DisplayName="Product ID" Name="AppProductID" AllowDeletion="FALSE" ReadOnly="TRUE" ShowInNewForm="FALSE" ShowInEditForm="FALSE" />'
            },
            { 
                name: "AppVersion",
                schemaXml: '<Field Type="Text" DisplayName="App Version" Name="AppVersion" AllowDeletion="FALSE" ShowInNewForm="TRUE" ShowInEditForm="FALSE" />'
            },
            {
                name: "IsDefaultAppMetadataLocale",
                schemaXml: '<Field Type="Boolean" DisplayName="Default Metadata Language" Name="IsDefaultAppMetadataLocale" AllowDeletion="FALSE" ShowInNewForm="FALSE" ShowInEditForm="FALSE"><Default>1</Default></Field>'
            },
            {
                name: "AppShortDescription",
                schemaXml: '<Field Type="Text" DisplayName="Short Description" Name="AppShortDescription" AllowDeletion="FALSE" Required="TRUE" />'
            },
            {
                name: "AppDescription",
                schemaXml: '<Field Type="Note" DisplayName="Description" Name="AppDescription" AllowDeletion="FALSE" Required="TRUE" />'
            },
            {
                name: "AppThumbnailURL",
                schemaXml: '<Field Type="URL" DisplayName="Icon URL" Name="AppThumbnailURL" AllowDeletion="FALSE" Required="TRUE" Format="Image" Description="The URL to the app icon. The icon should have a width and height of 96 pixels." JSLink="' + devAppsCSR + '" />'
            },
            {
                name: "SharePointAppCategory",
                schemaXml: '<Field Type="Choice" DisplayName="Category" Name="SharePointAppCategory" AllowDeletion="FALSE" FillInChoice="TRUE" />'
            },
            {
                name: "Owners",
                schemaXml: '<Field Type="UserMulti" DisplayName="Owners" Name="Owners" AllowDeletion="FALSE" List="UserInfo" Required="TRUE" EnforceUniqueValues="FALSE" ShowField="ImnName" UserSelectionMode="PeopleAndGroups" UserSelectionScope="0" Mult="TRUE" Sortable="FALSE" Description="Who is allowed to make updates to this app?"/>'
            },
            {
                name: "AppPublisher",
                schemaXml: '<Field Type="Text" DisplayName="Publisher Name" Name="AppPublisher" AllowDeletion="FALSE" Required="TRUE" />'
            },
            {
                name: "AppSupportURL",
                schemaXml: '<Field Type="URL" DisplayName="Support URL" Name="AppSupportURL" AllowDeletion="FALSE" Required="TRUE" />'
            },
            {
                name: "AppImageURL1",
                schemaXml: '<Field Type="URL" DisplayName="Image URL 1" Name="AppImageURL1" AllowDeletion="FALSE" Required="TRUE" Format="Image" Description="An image or screenshot for this app. Images should have a width of 512 pixes and a height of 384 pixels."/>'
            },
            {
                name: "AppImageURL2",
                schemaXml: '<Field Type="URL" DisplayName="Image URL 2" Name="AppImageURL2" AllowDeletion="FALSE" Required="TRUE" Format="Image" Description="An image or screenshot for this app. Images should have a width of 512 pixes and a height of 384 pixels."/>'
            },
            {
                name: "AppImageURL3",
                schemaXml: '<Field Type="URL" DisplayName="Image URL 3" Name="AppImageURL3" AllowDeletion="FALSE" Format="Image" Description="An image or screenshot for this app. Images should have a width of 512 pixes and a height of 384 pixels."/>'
            },
            {
                name: "AppImageURL4",
                schemaXml: '<Field Type="URL" DisplayName="Image URL 4" Name="AppImageURL4" AllowDeletion="FALSE" Format="Image" Description="An image or screenshot for this app. Images should have a width of 512 pixes and a height of 384 pixels."/>'
            },
            {
                name: "AppImageURL5",
                schemaXml: '<Field Type="URL" DisplayName="Image URL 5" Name="AppImageURL5" AllowDeletion="FALSE" Format="Image" Description="An image or screenshot for this app. Images should have a width of 512 pixes and a height of 384 pixels."/>'
            },
            {
                name: "AppVideoURL",
                schemaXml: '<Field Type="URL" DisplayName="Video URL" Name="AppVideoURL" AllowDeletion="FALSE" Description="A URL to a video file or web page with an embedded video." />'
            },
            {
                name: "IsAppPackageEnabled",
                schemaXml: '<Field Type="Boolean" DisplayName="Enabled" Name="IsAppPackageEnabled" AllowDeletion="FALSE"><Default>1</Default></Field>'
            },
            {
                name: "DevAppStatus",
                schemaXml: '<Field Type="Choice" DisplayName="App Status" Name="DevAppStatus" AllowDeletion="FALSE" Required="TRUE" Format="Dropdown" FillInChoice="FALSE"><Default>Missing Metadata</Default><CHOICES><CHOICE>Missing Metadata</CHOICE><CHOICE>Deleting...</CHOICE><CHOICE>In Testing</CHOICE><CHOICE>In Review</CHOICE><CHOICE>Published</CHOICE></CHOICES></Field>'
            }
        ],
        ViewInformation: [{
            ViewName: "All Documents",
            ViewFields: ["DocIcon", "LinkFilename", "AppPublisher", "DevAppStatus", "Modified", "Editor"],
            ViewQuery: '<OrderBy><FieldRef Name="FileLeafRef" Ascending="TRUE" /></OrderBy>',
            JSLink: devAppsCSR,
        }]
    },
    {
        ListInformation: {
            BaseTemplate: SPTypes.ListTemplateType.GenericList,
            Title: Strings.Lists.Assessments,
            Description: "Used by developers to submit their assessments of other apps",
            ListExperienceOptions: 2,
            //AllowContentTypes: true,
            //ContentTypesEnabled: false,
            //DisableGridEditing: true,
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
            /*{ 
                name: "RelatedApp",
                schemaXml: '<Field Type="Lookup" DisplayName="Related App" Name="RelatedApp" Indexed="TRUE" AllowDeletion="FALSE" Required="TRUE" EnforceUniqueValues="FALSE" List="DeveloperApps" ShowField="Title" UnlimitedLengthInDocumentLibrary="FALSE" RelationshipDeleteBehavior="Cascade" />'
            },*/
            {
                name: "RelatedApp",
                title: "Related App",
                type: Helper.SPCfgFieldType.Lookup,
                required: true,
                listName: "Developer Apps",
                showField: "Title",
                jslink: devAppAssessmentsCSR,
                indexed: true,
                allowDeletion: false,
                relationshipBehavior: SPTypes.RelationshipDeleteBehaviorType.Cascade
            } as Helper.IFieldInfoLookup,
            { 
                name: "SupportUrlValid",
                schemaXml: '<Field Type="Choice" DisplayName="Support URL" Name="SupportUrlValid" AllowDeletion="FALSE" Required="TRUE" Format="Dropdown" Description="Does the &quot;Support URL&quot; have appropriate information to aid users in resolving problems or contacting the app\'s developers?"><CHOICES><CHOICE>Yes</CHOICE><CHOICE>No</CHOICE></CHOICES></Field>'
            },
            {
                name: "IconImages",
                schemaXml: '<Field Type="Choice" DisplayName="Icon/Images" Name="IconImages" AllowDeletion="FALSE" Required="TRUE" Format="Dropdown" Description="Are the app\'s icon and screenshots/images valid and sufficient?"><CHOICES><CHOICE>Yes</CHOICE><CHOICE>No</CHOICE></CHOICES></Field>'
            },
            {
                name: "DescriptionText",
                schemaXml: '<Field Type="Choice" DisplayName="Description text" Name="DescriptionText" AllowDeletion="FALSE" Required="TRUE" Format="Dropdown" Description="Are the app\'s description text fields valid and sufficient to describe the app to users?"><CHOICES><CHOICE>Yes</CHOICE><CHOICE>No</CHOICE></CHOICES></Field>'
            },
            {
                name: "AppTested",
                schemaXml: '<Field Type="Choice" DisplayName="App tested" Name="AppTested" AllowDeletion="FALSE" Required="TRUE" Format="Dropdown" Description="Have you tested the app and can you confirm it\'s sufficiently functional?"><CHOICES><CHOICE>Yes</CHOICE><CHOICE>No</CHOICE></CHOICES></Field>'
            },
            {
                name: "UnresolvedIssues",
                schemaXml: '<Field Type="Choice" DisplayName="Unresolved Issues?" Name="UnresolvedIssues" AllowDeletion="FALSE" Required="TRUE" Format="Dropdown" FillInChoice="FALSE" Description="Are there any unresolved issues with the app that would cause you to not recommend the app be published to the tenant app catalog for all users to see?"><CHOICES><CHOICE>Yes</CHOICE><CHOICE>No</CHOICE></CHOICES></Field>'
            },
            {
                name: "Comments",
                schemaXml: '<Field Type="Note" DisplayName="Comments" Name="Comments" AllowDeletion="FALSE" Required="FALSE" NumLines="4" RichText="FALSE" Sortable="FALSE" Description="Enter any applicable comments you have for the app developer based on your responses above."/>'
            },
            {
                name: "ResultResponse",
                schemaXml: '<Field Type="Calculated" DisplayName="Result Response" Name="ResultResponse" AllowDeletion="FALSE" Required="FALSE" Format="DateOnly" LCID="1033" ResultType="Text" ReadOnly="TRUE"><Formula>=IF([Support URL]&amp;[Icon/Images]&amp;[Description text]&amp;[App tested]&amp;[Unresolved Issues?]="YesYesYesYesNo","Recommend","Oppose")</Formula><FieldRefs><FieldRef Name="UnresolvedIssues" /><FieldRef Name="AppTested" /><FieldRef Name="DescriptionText" /><FieldRef Name="IconImages" /><FieldRef Name="SupportUrlValid" /></FieldRefs></Field>'
            }
        ],
        ViewInformation: [{
            ViewName: "All Items",
            ViewFields: ["Title", "RelatedApp", "SupportUrlValid", "IconImages", "DescriptionText", "AppTested", "UnresolvedIssues", "ResultResponse", "Editor"],
            ViewQuery: '<OrderBy><FieldRef Name="ID" Ascending="FALSE" /></OrderBy>'
        }]
    },
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
    }]
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