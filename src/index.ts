import { LoadingDialog, Timeout } from "dattatable";
import { ContextInfo, ThemeManager } from "gd-sprest-bs";
import { AppDashboard } from "./appDashboard";
import { AppInstall } from "./appInstall";
import { AppSecurity } from "./appSecurity";
import { AppView } from "./appView";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import { ErrorDialog } from "./errorDialog";
import Strings, { setContext } from "./strings";

// Styling
import "./styles.scss";

// Properties
interface IProps {
    el: HTMLElement;
    appConfiguration?: string
    context?: any;
    displayMode?: number;
    envType?: number;
    log?: any,
    sourceUrl?: string;
}

// Create the global variable for this solution
const GlobalVariable = {
    App: null,
    Configuration,
    description: Strings.ProjectDescription,
    render: (props: IProps) => {
        // See if the log is set
        if (props.log) {
            // Set the log and scope
            ErrorDialog.Log = props.log;
            ErrorDialog.Scope = props.context.serviceScope;
        }

        // Log
        ErrorDialog.logInfo("App Catalog Manager", "Initializing the solution...");

        // See if the page context exists
        if (props.context) {
            // Set the context
            setContext(props.context, props.envType, props.sourceUrl);

            // Set the web url
            Configuration.setWebUrl(props.sourceUrl || ContextInfo.webServerRelativeUrl);
        }

        // Hide the first column of the webpart zones
        let wpZone: HTMLElement = document.querySelector("#DeltaPlaceHolderMain table > tbody > tr > td");
        wpZone ? wpZone.style.width = "0%" : null;

        // Make the second column of the webpart zones full width
        wpZone = document.querySelector("#DeltaPlaceHolderMain table > tbody > tr > td:last-child");
        wpZone ? wpZone.style.width = "100%" : null;

        // Show a loading dialog
        LoadingDialog.setHeader("Loading Application");
        LoadingDialog.setBody("This may take time based on the number of apps to load...");
        LoadingDialog.show();

        // Initialize the solution
        DataSource.init(props.appConfiguration).then(
            // Success
            () => {
                // Hide the loading dialog
                LoadingDialog.hide();

                // Ensure the security groups exist
                if (AppSecurity.hasErrors()) {
                    // See if an install is required
                    AppInstall.InstallRequired(props.el);
                } else {
                    // Create the app elements
                    props.el.innerHTML = "<div id='apps'></div><div id='app-details' class='d-none'></div>";
                    let elApps = props.el.querySelector("#apps") as HTMLElement;
                    let elAppDetails = props.el.querySelector("#app-details") as HTMLElement;

                    // See if this is a document set and we are not in teams
                    if (!Strings.IsTeams && DataSource.AppItem) {
                        // View the application dashboard
                        new AppDashboard(elAppDetails, elApps, DataSource.AppItem.Id);

                        // Show the details
                        elAppDetails.classList.remove("d-none");
                    } else {
                        // View all of the applications
                        GlobalVariable.App = new AppView(elApps, elAppDetails);
                    }

                    // Start the timeout
                    Timeout.setTimer(300000);
                    Timeout.start();
                }
            },
            // Error
            () => {
                // See if an install is required
                AppInstall.InstallRequired(props.el);

                // Hide the loading dialog
                LoadingDialog.hide();
            }
        );
    },
    updateTheme: (themeInfo) => {
        // Set the theme
        ThemeManager.setCurrentTheme(themeInfo);
    },
    version: Strings.Version
};

// Update the DOM
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Remove the extra border spacing on the webpart
    let contentBox = document.querySelector("#contentBox table.ms-core-tableNoSpace");
    contentBox ? contentBox.classList.remove("ms-webpartPage-root") : null;

    // Hide the page title if it exists
    let pageTitle = document.querySelector("#pageTitle");
    pageTitle ? pageTitle.setAttribute("style", "display:none;") : null;

    // Render the application
    GlobalVariable.render({ el: elApp });
}