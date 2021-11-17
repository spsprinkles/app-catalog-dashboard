import { ContextInfo } from "gd-sprest-bs";

// Constant
export const SourceUrl: string = ContextInfo.webServerRelativeUrl + "/SiteAssets/app-dashboard/";

/**
 * Global Constants
 */
export default {
    AppElementId: "project-app",
    GlobalVariable: "AppDashboard",
    ProjectName: "App Dashboard",
    ProjectDescription: "App Dashboard for Developers",
    Group: "App Developers",
    Lists: {
        Apps: "Developer Apps",
        Assessments: "Dev App Assessments"
    },
    SolutionAppsCSRUrl: "~site/code/devAppsCSR.js",
    SolutionAppAssessmentsCSRUrl: "~site/code/devAppAssessmentsCSR.js",
    SolutionUrl: SourceUrl + "index.html",
    TermsOfUseUrl: SourceUrl + "TermsOfUse.html",
}