import { ContextInfo } from "gd-sprest-bs";

/**
 * Global Constants
 */

// Updates the strings for SPFx
export const setContext = (context) => {
    // Set the page context
    ContextInfo.setPageContext(context);

    // Update the values
    Strings.SolutionUrl = ContextInfo.webServerRelativeUrl + "/SiteAssets/Event-Registration/index.html";
    Strings.TermsOfUseUrl = ContextInfo.webServerRelativeUrl + "/siteassets/TermsOfUse.html";
}

// Strings
const Strings = {
    AppElementId: "app-catalog-dashboard",
    GlobalVariable: "AppDashboard",
    Group: "App Developers",
    Lists: {
        Apps: "Developer Apps",
        Assessments: "Dev App Assessments"
    },
    ProjectName: "App Dashboard",
    ProjectDescription: "App Dashboard for Developers",
    SolutionAppsCSRUrl: "~site/code/devAppsCSR.js",
    SolutionAppAssessmentsCSRUrl: "~site/code/devAppAssessmentsCSR.js",
    SolutionUrl: ContextInfo.webServerRelativeUrl + "/siteassets/index.html",
    TermsOfUseUrl: ContextInfo.webServerRelativeUrl + "/siteassets/TermsOfUse.html",
    Version: "0.1",
};
export default Strings;