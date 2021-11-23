import { ContextInfo } from "gd-sprest-bs";

/**
 * Global Constants
 */

// Global Path
let SourceUrl: string = ContextInfo.webServerRelativeUrl + "/SiteAssets/";

// Updates the strings for SPFx
export const setContext = (context) => {
    // Set the page context
    ContextInfo.setPageContext(context);

    // Update the global path
    SourceUrl = ContextInfo.webServerRelativeUrl + "/SiteAssets/";

    // Update the values
    Strings.ConfigUrl = SourceUrl + "config.json";
    Strings.SolutionUrl = SourceUrl + "index.html";
    Strings.TermsOfUseUrl = SourceUrl + "TermsOfUse.html";
}

// Strings
const Strings = {
    AppElementId: "app-catalog-dashboard",
    ConfigUrl: SourceUrl + "config.json",
    DashboardUrl: SourceUrl + "dashboard.aspx",
    GlobalVariable: "AppDashboard",
    Group: "App Developers",
    Lists: {
        Apps: "Developer Apps",
        Assessments: "Dev App Assessments"
    },
    ProjectName: "App Dashboard",
    ProjectDescription: "App Dashboard for Developers",
    SolutionUrl: SourceUrl + "index.html",
    TermsOfUseUrl: SourceUrl + "TermsOfUse.html",
    Version: "0.1",
};
export default Strings;