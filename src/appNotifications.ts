import { LoadingDialog } from "dattatable";
import { ContextInfo, Utility } from "gd-sprest-bs";
import { AppConfig } from "./appCfg";
import { AppSecurity } from "./appSecurity";
import { IAppItem } from "./ds";

/**
 * Application Notifications
 */
export class AppNotifications {
    // Sends a rejection notification
    static rejectEmail(status: string, item: IAppItem, comments: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Parse the developers
            let CC = [];
            for (let i = 0; i < AppSecurity.DevGroup.Users.results.length; i++) {
                // Append the email
                CC.push(AppSecurity.DevGroup.Users.results[i].Email);
            }

            // Get the app developers
            let To = [];
            let owners = item.AppDevelopers && item.AppDevelopers.results ? item.AppDevelopers.results : [];
            for (let i = 0; i < owners.length; i++) {
                // Append the email
                To.push(owners[i].EMail);
            }

            // Ensure emails exist
            if (To.length > 0 || CC.length > 0) {
                // Display a loading dialog
                LoadingDialog.setHeader("Sending Notification");
                LoadingDialog.setBody("This dialog will close after the notification is sent.");
                LoadingDialog.show();

                // Send an email
                Utility().sendEmail({
                    To, CC,
                    Subject: "App '" + item.Title + "' Sent Back",
                    Body: "App Developers,<br /><br />The '" + item.Title +
                        "' app has been sent back based on the comments below.<br /><br />" + comments
                }).execute(() => {
                    // Close the loading dialog
                    LoadingDialog.hide();

                    // Resolve the request
                    resolve();
                });
            } else {
                // Resolve the request
                resolve();
            }
        });
    }

    // Sends an email notification based on the status
    static sendEmail(status: string, item: IAppItem): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Parse the email configurations for this status
            let notificationCfgs = (AppConfig.Status[status] ? AppConfig.Status[status].notification : null) || [];
            for (let i = 0; i < notificationCfgs.length; i++) {
                let notificationCfg = notificationCfgs[i];
                let CC: string[] = [];
                let To: string[] = [];

                // Parse the to configuration
                let toEmails = notificationCfg.to || [];
                for (let i = 0; i < toEmails.length; i++) {
                    // Append the email
                    To.push(toEmails[i]);
                }

                // Parse the cc configuration
                let ccEmails = notificationCfg.cc || [];
                for (let i = 0; i < ccEmails.length; i++) {
                    // Append the email
                    CC.push(ccEmails[i]);
                }

                // Parse the email content and find [[Key]] instances w/in it.
                let Body = notificationCfg.content || "";
                let startIdx = Body.indexOf("[[");
                let endIdx = Body.indexOf("]]");
                while (startIdx > 0 && endIdx > startIdx) {
                    // Get the key value
                    let key = Body.substring(startIdx + 2, endIdx);
                    let value = "";

                    // See if this is the status
                    if (key == "Status") {
                        // Set the value
                        value = status;
                    }
                    // Else, see if this is a field value
                    else if (item[key]) {
                        // Set the value
                        value = item[key];
                    }
                    // Else, see if this is the user name
                    else if (key == "User.Name") {
                        // Set the value
                        value = ContextInfo.userDisplayName;
                    }

                    // Replace the value
                    let oldContent = Body;
                    Body = Body.substring(0, startIdx) + value + Body.substring(endIdx + 2);

                    // Find the next instance of it
                    startIdx = oldContent.indexOf("[[", endIdx);
                    endIdx = oldContent.indexOf("]]", endIdx + 2);
                }

                // Ensure emails exist
                if (To.length > 0 || CC.length > 0) {
                    // Display a loading dialog
                    LoadingDialog.setHeader("Sending Notification");
                    LoadingDialog.setBody("This dialog will close after the notification is sent.");
                    LoadingDialog.show();

                    // Send an email
                    Utility().sendEmail({
                        To, CC,
                        Body,
                        Subject: notificationCfg.subject,
                    }).execute(() => {
                        // Close the loading dialog
                        LoadingDialog.hide();

                        // Resolve the request
                        resolve();
                    });

                } else {
                    // Resolve the request
                    resolve();
                }
            }

        });
    }
}