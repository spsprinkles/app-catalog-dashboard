import { ContextInfo, Helper } from "gd-sprest-bs";

/**
 * Global Constants
 */

// Global Path
let AssetsUrl: string = ContextInfo.webServerRelativeUrl + "/SiteAssets/";
let PagesUrl: string = ContextInfo.webServerRelativeUrl + "/SitePages/";
let SPFxContext: { pageContext: any; sdks: { microsoftTeams: any } } = null;

// Updates the strings for SPFx
export const setContext = (context, sourceUrl?: string) => {
    // Set the page context
    SPFxContext = context;
    ContextInfo.setPageContext(SPFxContext.pageContext);

    // Load the default scripts
    Helper.loadSPCore();

    // Set the teams flag
    Strings.IsTeams = SPFxContext.sdks.microsoftTeams ? true : false;

    // Update the global path
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
    AssetsUrl = Strings.SourceUrl + "/SiteAssets/";
    PagesUrl = Strings.SourceUrl + "/SitePages/";

    // Update the values
    Strings.DashboardUrl = PagesUrl + "Dashboard.aspx";
    Strings.SolutionUrl = AssetsUrl + "index.html"; 
}

// Strings
const Strings = {
    AppElementId: "app-catalog-dashboard",
    ConfigUrl: "SiteAssets/config.json",
    DashboardUrl: PagesUrl + "Dashboard.aspx",
    GlobalVariable: "AppDashboard",
    Groups: {
        Approvers: "App Approvers",
        Developers: "App Developers",
        FinalApprovers: "App Final Approvers",
        Sponsors: "App Sponsors"
    },
    IsTeams: false,
    Lists: {
        AppCatalog: "Apps for SharePoint",
        Apps: "Developer Apps",
        AppCatalogRequests: "App Catalog Requests",
        AuditLog: "App Logs",
        Assessments: "Dev App Assessments"
    },
    ProjectName: "App Dashboard",
    ProjectDescription: "App Dashboard for Developers",
    SolutionUrl: AssetsUrl + "index.html",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    Version: "0.0.6.4",
};
export default Strings;