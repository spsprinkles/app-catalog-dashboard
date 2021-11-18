import { ContextInfo } from "gd-sprest-bs";

/**
 * Global Constants
 */

// Updates the strings for SPFx
export const setContext = (context) => {
    // Set the page context
    ContextInfo.setPageContext(context);

    // Update the values
    Strings.ConfigUrl = ContextInfo.webServerRelativeUrl + "/SiteAssets/config.json";
    Strings.SolutionUrl = ContextInfo.webServerRelativeUrl + "/SiteAssets/index.html";
    Strings.TermsOfUseUrl = ContextInfo.webServerRelativeUrl + "/siteassets/TermsOfUse.html";
}

// Strings
const Strings = {
    AppElementId: "app-catalog-dashboard",
    ConfigUrl: ContextInfo.webServerRelativeUrl + "siteassets/config.json",
    GlobalVariable: "AppDashboard",
    Group: "App Developers",
    Lists: {
        Apps: "Developer Apps",
        Assessments: "Dev App Assessments"
    },
    ProjectName: "App Dashboard",
    ProjectDescription: "App Dashboard for Developers",
    SolutionAppsCSRUrl: "~site/siteassets/devAppsCSR.js",
    SolutionAppAssessmentsCSRUrl: "~site/siteassets/devAppAssessmentsCSR.js",
    SolutionUrl: ContextInfo.webServerRelativeUrl + "/siteassets/index.html",
    TermsOfUseUrl: ContextInfo.webServerRelativeUrl + "/siteassets/TermsOfUse.html",
    Version: "0.1",
};
export default Strings;