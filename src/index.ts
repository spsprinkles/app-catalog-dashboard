import { App } from "./app";
import { Configuration } from "./cfg";
import { AppDashboard } from "./dashboard";
import { DataSource } from "./ds";
import Strings, { setContext } from "./strings";
import { UserAgreement } from "./userAgreement";

// Styling
import "./styles.scss";

// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    render: (el: HTMLElement, context?, sourceUrl?: string) => {
        // See if the page context exists
        if (context) {
            // Set the context
            setContext(context, sourceUrl);
        }

        // Hide the first column of the webpart zones
        let wpZone: HTMLElement = document.querySelector("#DeltaPlaceHolderMain table > tbody > tr > td");
        wpZone ? wpZone.style.width = "0%" : null;

        // Make the second column of the webpart zones full width
        wpZone = document.querySelector("#DeltaPlaceHolderMain table > tbody > tr > td:last-child");
        wpZone ? wpZone.style.width = "100%" : null;

        // Initialize the solution
        DataSource.init().then(
            // Success
            () => {
                // Create the app elements
                el.innerHTML = "<div id='apps'></div><div id='app-details' style='display: none;'></div>";
                let elApps = el.querySelector("#apps") as HTMLElement;
                let elAppDetails = el.querySelector("#app-details") as HTMLElement;

                // Ensure the security groups exist
                if (DataSource.ApproverGroup == null || DataSource.DevGroup == null) {
                    // See if an install is required
                    DataSource.InstallRequired();
                }
                // Ensure the user is not an approver or developer
                if (!DataSource.IsApprover && !DataSource.IsDeveloper) {
                    // Show the user agreement
                    new UserAgreement();
                }
                // Else, see if this is a document set and we are not in teams
                else if (!Strings.IsTeams && DataSource.DocSetItem) {
                    // Create the application
                    new App(elAppDetails);
                } else {
                    // Render the dashboard
                    new AppDashboard(elApps, elAppDetails);
                }
            },
            // Error
            () => {
                // Ensure the user is not an approver or developer
                if (!DataSource.IsApprover && !DataSource.IsDeveloper) {
                    // Show the user agreement
                    new UserAgreement();
                } else {
                    // See if an install is required
                    DataSource.InstallRequired();
                }
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