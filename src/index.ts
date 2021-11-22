import { InstallationRequired } from "dattatable";
import { App } from "./app";
import { Configuration } from "./cfg";
import { AppDashboard } from "./dashboard";
import { DataSource } from "./ds";
import Strings, { setContext } from "./strings";

// Styling
import "./styles.scss";

// Create the global variable for this solution
const GlobalVariable = {
    Configuration,
    render: (el: HTMLElement, context?) => {
        // See if the page context exists
        if (context) {
            // Set the context
            setContext(context);
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
                // See if this is a document set
                if (DataSource.IsDocSet) {
                    // Create the application
                    new App(el);
                } else {
                    // Render the dashboard
                    new AppDashboard(el);
                }
            },
            // Error
            () => {
                // See if an install is required
                InstallationRequired.requiresInstall(Configuration).then(installFl => {
                    // See if an installation is required
                    if (installFl) {
                        // Show the installation dialog
                        InstallationRequired.showDialog();
                    } else {
                        // Log
                        console.error("[" + Strings.ProjectName + "] Error initializing the solution.");
                    }
                });
            }
        );
    }
};

// Update the DOM
window[Strings.GlobalVariable] = GlobalVariable;

// Get the element and render the app if it is found
let elApp = document.querySelector("#" + Strings.AppElementId) as HTMLElement;
if (elApp) {
    // Render the application
    GlobalVariable.render(elApp);
}