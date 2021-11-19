import { ContextInfo } from "gd-sprest-bs";

/**
 * Global Constants
 */

// Set global path
export const SourceUrl: string = ContextInfo.webServerRelativeUrl + "/SiteAssets/";

// Updates the strings for SPFx
export const setContext = (context) => {
    // Set the page context
    ContextInfo.setPageContext(context);

    // Update the values
    Strings.ConfigUrl = SourceUrl + "config.json";
    Strings.SolutionUrl = SourceUrl + "index.html";
    Strings.TermsOfUseUrl = SourceUrl + "TermsOfUse.html";
}

// Strings
const Strings = {
    AppElementId: "app-catalog-dashboard",
    ConfigUrl: SourceUrl + "config.json",
    GlobalVariable: "AppDashboard",
    Group: "App Developers",
    Lists: {
        Apps: "Developer Apps",
        Assessments: "Dev App Assessments"
    },
    ProjectName: "App Dashboard",
    ProjectDescription: "App Dashboard for Developers",
    SolutionAppsCSRUrl: "~site/SiteAssets/devAppsCSR.js",
    SolutionAppAssessmentsCSRUrl: "~site/SiteAssets/devAppAssessmentsCSR.js",
    SolutionUrl: SourceUrl + "index.html",
    TermsOfUseUrl: SourceUrl + "TermsOfUse.html",
    Version: "0.1",
};
export default Strings;