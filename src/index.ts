import { ContextInfo } from "gd-sprest-bs";
import { AppDashboard } from "./appDashboard";
import { AppSecurity } from "./appSecurity";
import { AppView } from "./appView";
import { Configuration } from "./cfg";
import { DataSource } from "./ds";
import Strings, { setContext } from "./strings";

// Styling
import "./styles.scss";

// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    render: (el: HTMLElement, context?, cfgUrl?: string, sourceUrl?: string) => {
        // See if the page context exists
        if (context) {
            // Set the context
            setContext(context, sourceUrl);

            // Set the web url
            Configuration.setWebUrl(sourceUrl || ContextInfo.webServerRelativeUrl);
        }

        // Hide the first column of the webpart zones
        let wpZone: HTMLElement = document.querySelector("#DeltaPlaceHolderMain table > tbody > tr > td");
        wpZone ? wpZone.style.width = "0%" : null;

        // Make the second column of the webpart zones full width
        wpZone = document.querySelector("#DeltaPlaceHolderMain table > tbody > tr > td:last-child");
        wpZone ? wpZone.style.width = "100%" : null;

        // Initialize the solution
        DataSource.init(cfgUrl).then(
            // Success
            () => {
                // Ensure the security groups exist
                if (AppSecurity.ApproverGroup == null || AppSecurity.DevGroup == null) {
                    // See if an install is required
                    DataSource.InstallRequired(el);
                } else {
                    // Create the app elements
                    el.innerHTML = "<div id='apps'></div><div id='app-details' class='d-none'></div>";
                    let elApps = el.querySelector("#apps") as HTMLElement;
                    let elAppDetails = el.querySelector("#app-details") as HTMLElement;

                    // See if this is a document set and we are not in teams
                    if (!Strings.IsTeams && DataSource.DocSetItem) {
                        // View the application dashboard
                        new AppDashboard(elAppDetails, elApps);

                        // Show the details
                        elAppDetails.classList.remove("d-none");
                    } else {
                        // View all of the applications
                        new AppView(elApps, elAppDetails);
                    }
                }
            },
            // Error
            () => {
                // See if an install is required
                DataSource.InstallRequired(el);
            }
        );
    }
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
    GlobalVariable.render(elApp);
}