import { ContextInfo } from "gd-sprest-bs";

/**
 * Global Constants
 */

// Global Path
let AssetsUrl: string = ContextInfo.webServerRelativeUrl + "/SiteAssets/";
let SPFxContext: { pageContext: any; sdks: { microsoftTeams: any } } = null;

// Updates the strings for SPFx
export const setContext = (context, sourceUrl?: string) => {
    // Set the page context
    SPFxContext = context;
    ContextInfo.setPageContext(SPFxContext.pageContext);

    // Set the teams flag
    Strings.IsTeams = SPFxContext.sdks.microsoftTeams ? true : false;

    // Update the global path
    Strings.SourceUrl = sourceUrl || ContextInfo.webServerRelativeUrl;
    AssetsUrl = Strings.SourceUrl + "/SiteAssets/";

    // Update the values
    Strings.ConfigUrl = AssetsUrl + "config.json";
    Strings.SolutionUrl = AssetsUrl + "index.html";
    Strings.TermsOfUseUrl = AssetsUrl + "TermsOfUse.html";
}

// Strings
const Strings = {
    AppElementId: "app-catalog-dashboard",
    ConfigUrl: AssetsUrl + "config.json",
    DashboardUrl: AssetsUrl + "dashboard.aspx",
    GlobalVariable: "AppDashboard",
    Groups: {
        Approvers: "App Approvers",
        Developers: "App Developers"
    },
    IsTeams: false,
    Lists: {
        Apps: "Developer Apps",
        Assessments: "Dev App Assessments"
    },
    ProjectName: "App Dashboard",
    ProjectDescription: "App Dashboard for Developers",
    SolutionUrl: AssetsUrl + "index.html",
    SourceUrl: ContextInfo.webServerRelativeUrl,
    TermsOfUseUrl: AssetsUrl + "TermsOfUse.html",
    UserAgreementUrl: AssetsUrl + "user_agreement.html",
    Version: "0.1",
};
export default Strings;