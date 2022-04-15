import { LoadingDialog } from "dattatable";
import { ContextInfo, Utility } from "gd-sprest-bs";
import { AppConfig } from "./appCfg";
import { AppSecurity } from "./appSecurity";
import { IAppItem } from "./ds";
import Strings from "./strings";

/**
 * Application Notifications
 */
export class AppNotifications {
    // Gets the emails for a specific key
    private static getEmails(key: string, item: IAppItem) {
        // See which email(s) we need to return
        switch (key) {
            // Approver's Group
            case "ApproversGroup":
                return AppSecurity.ApproverEmails;

            // Developer's Group
            case "DevelopersGroup":
                return AppSecurity.DeveloperEmails;

            // Developers
            case "Developers":
                let emails = [];

                // Parse the developers
                let developers = item.AppDevelopers.results;
                for (let i = 0; i < developers.length; i++) {
                    // Append the email
                    emails.push(developers[i].EMail);
                }

                // Return the emails
                return emails;

            // Sponsor
            case "Sponsor":
                // Return the email
                return item.AppSponsor ? [item.AppSponsor.EMail] : [];
        }
    }

    // Gets the value
    private static getValue(key: string, item: IAppItem, status: string) {
        // Default the value
        let value = item[key];

        // Return the value
        switch (key) {
            // See if this is the status field
            case "AppStatus":
                // Set the value
                value = status;
                break;

            // Url to the dashboard page
            case "PageUrl":
                // Set the value
                let pageUrl = window.location.origin + AppConfig.Configuration.dashboardUrl + "?app-id=" + item.Id;
                value = "<a href='" + pageUrl + "'>" + pageUrl + "</a>";
                break;

            // Url to the dashboard page
            case "PageUrlText":
                // Set the value
                value = window.location.origin + AppConfig.Configuration.dashboardUrl + "?app-id=" + item.Id;
                break;

            // Current user's name
            case "User.Name":
                // Set the value
                value = ContextInfo.userDisplayName;
                break;

            // Item Metadata
            default:
                let values = [];

                // Default the value to be a collection if it's not one
                let results = value && value.results ? value.results : [value];
                value = "";

                // Parse the collection
                for (let i = 0; i < results.length; i++) {
                    let result = results[i];
                    if (result) {
                        // Parse the properties of complex fields
                        let keys = key.split('.');
                        for (let j = 0; j < keys.length; j++) {
                            // Set the value
                            value = result[key] ? value[keys[j]] : null;
                        }

                        // Append the value if it exists
                        value ? values.push(value) : null;
                    }
                }

                // Set the value
                value = values.join(', ');
                break;
        }

        // Return the value
        return value;
    }

    // Sends a rejection notification
    static rejectEmail(status: string, item: IAppItem, comments: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Get the app developers
            let To = [];
            let owners = item.AppDevelopers && item.AppDevelopers.results ? item.AppDevelopers.results : [];
            for (let i = 0; i < owners.length; i++) {
                // Append the email
                To.push(owners[i].EMail);
            }

            // Ensure emails exist
            if (To.length > 0 || AppSecurity.DeveloperEmails.length > 0) {
                // Display a loading dialog
                LoadingDialog.setHeader("Sending Notification");
                LoadingDialog.setBody("This dialog will close after the notification is sent.");
                LoadingDialog.show();

                // Send an email
                Utility(Strings.SourceUrl).sendEmail({
                    To, CC: AppSecurity.DeveloperEmails,
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
            // Get the notification configuration and ensure they exist
            let notificationCfgs = (AppConfig.Status[status] ? AppConfig.Status[status].notification : null) || [];
            if (notificationCfgs.length == 0) {
                // Resolve the request
                resolve();
                return;
            }

            // Parse the email configurations for this status
            for (let i = 0; i < notificationCfgs.length; i++) {
                let notificationCfg = notificationCfgs[i];
                let CC: string[] = [];
                let To: string[] = [];

                // Parse the to configuration
                let toEmails = notificationCfg.to || [];
                for (let i = 0; i < toEmails.length; i++) {
                    // Append the emails
                    let emails = this.getEmails(toEmails[i], item);
                    To = To.concat(emails);
                }

                // Parse the cc configuration
                let ccEmails = notificationCfg.cc || [];
                for (let i = 0; i < ccEmails.length; i++) {
                    // Append the emails
                    let emails = this.getEmails(ccEmails[i], item);
                    CC = CC.concat(emails);
                }

                // Parse the email content and find [[Key]] instances w/in it.
                let Body = notificationCfg.content || "";
                let startIdx = Body.indexOf("[[");
                let endIdx = Body.indexOf("]]");
                while (startIdx > 0 && endIdx > startIdx) {
                    // Get the key value
                    let key = Body.substring(startIdx + 2, endIdx);
                    let value = this.getValue(key, item, status);

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
                    Utility(Strings.SourceUrl).sendEmail({
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


    // Sends an email notification
    static sendNotification(item: IAppItem, userTypes: string[], subject: string, body: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            let To: string[] = [];

            // Parse the user types
            for (let i = 0; i < userTypes.length; i++) {
                // Append the email addresses
                let emails = this.getEmails(userTypes[i], item);
                To = To.concat(emails);
            }

            // Parse the email content and find [[Key]] instances w/in it.
            let Body = body || "";
            let startIdx = Body.indexOf("[[");
            let endIdx = Body.indexOf("]]");
            while (startIdx > 0 && endIdx > startIdx) {
                // Get the key value
                let key = Body.substring(startIdx + 2, endIdx);
                let value = this.getValue(key, item, item.AppStatus);

                // Replace the value
                let oldContent = Body;
                Body = Body.substring(0, startIdx) + value + Body.substring(endIdx + 2);

                // Find the next instance of it
                startIdx = oldContent.indexOf("[[", endIdx);
                endIdx = oldContent.indexOf("]]", endIdx + 2);
            }

            // Ensure emails exist
            if (To.length > 0) {
                // Display a loading dialog
                LoadingDialog.setHeader("Sending Notification");
                LoadingDialog.setBody("This dialog will close after the notification is sent.");
                LoadingDialog.show();

                // Send an email
                Utility(Strings.SourceUrl).sendEmail({
                    To,
                    Body,
                    Subject: subject,
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
}