import { Components, ContextInfo, Utility } from "gd-sprest-bs";
import { CanvasForm, LoadingDialog, Modal } from "dattatable";
import { AppConfig } from "./appCfg";
import Strings from "./strings";

/**
 * Error Dialog
 */
export class ErrorDialog {
    // Sends an email, based on the configuration
    private static sendEmail(title: string, message: string, ex?: any) {
        // Ensure a configuration exists
        if (AppConfig.Configuration.errorEmails && AppConfig.Configuration.errorEmails.length > 0) {
            // Set the exception
            let exception = "";
            if (ex) {
                // Try to stringify it
                try { exception = JSON.stringify(ex); }
                catch {
                    let exObj = {};

                    // Parse the properties
                    for (let key in ex) {
                        let value = ex[key];

                        // Only copy specific types
                        if (typeof (value) === "boolean" || typeof (value) === "number" || typeof (value) == "string") {
                            exObj[key] = value;
                        }
                    }

                    // Try to parse it again
                    try { exception = JSON.stringify(exObj); }
                    catch {
                        // Set the exception
                        exception = "Unable to stringify the exception details.";
                    }
                }
            }

            // Send a notification
            Utility().sendEmail({
                To: AppConfig.Configuration.errorEmails,
                Subject: Strings.ProjectName + " Error",
                Body: `
                    <h6>App Catalog Admins,</h6>
                    <p>An error occurred in the app. Details are listed below.</p>
                    <ul>
                        <li><b>User: </b>${ContextInfo.userDisplayName}</li>
                        <li><b>User Email: </b>${ContextInfo.userEmail}</li>
                        <li><b>Title: </b>${title}</li>
                        <li><b>Details: </b>${message}</li>
                        <li><b>Exception Response: </b>${ex && ex.response ? ex.response : ""}</li>
                        <li><b>Exception Details: </b>${exception}</li>
                    </ul>
                    <p>r/,</p>
                    <h6>App Catalog Manager Tool</h6>
                `.trim()
            }).execute();
        }
    }

    // Shows an error
    static show(title: string, message: string, ex?: any) {
        // Log the exception
        console.error(`[${Strings.ProjectName}][${title}] ${message}`, ex);

        // Close the dialogs/forms
        CanvasForm.hide();
        Modal.hide();
        LoadingDialog.hide();

        // Clear the modal
        Modal.clear();

        // Set the modal
        Modal.setHeader("Error: " + title);
        Modal.setBody(message);
        Modal.setType(Components.ModalTypes.Large);

        // Show the modal
        setTimeout(() => { Modal.show(); }, 250);

        // Send the email
        this.sendEmail(title, message, ex);
    }
}