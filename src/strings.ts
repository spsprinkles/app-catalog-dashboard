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
    SolutionUrl: SourceUrl + "index.html",
    TermsOfUseUrl: SourceUrl + "TermsOfUse.html",
    Lists: {
        Apps: "Developer Apps",
        Assessments: "Dev App Assessments"
    },
    Group: "App Developers"
}