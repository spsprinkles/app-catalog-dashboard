import { LoadingDialog } from "dattatable";
import { ContextInfo, Types, Utility } from "gd-sprest-bs";
import { AppConfig, IEmail } from "./appCfg";
import { AppSecurity } from "./appSecurity";
import * as Common from "./common";
import { IAppItem } from "./ds";
import { ErrorDialog } from "./errorDialog";
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
                return AppSecurity.AppWeb.getEmailsForGroup(Strings.Groups.Approvers);

            // Current User
            case "CurrentUser":
                return AppSecurity.AppWeb.CurrentUser.Email;

            // Developers
            case "Developers":
                let emails = [];

                // Parse the developers
                let developers = item.AppDevelopers.results;
                for (let i = 0; i < developers.length; i++) {
                    // Get the email from the login name
                    let email = Common.getEmail(developers[i].Name);

                    // Append the email
                    email ? emails.push(email) : null;
                }

                // Return the emails
                return emails;

            // Developer's Group
            case "DevelopersGroup":
                return AppSecurity.AppWeb.getEmailsForGroup(Strings.Groups.Developers);

            // Final Approver's Group
            case "FinalApproversGroup":
                return AppSecurity.AppWeb.getEmailsForGroup(Strings.Groups.FinalApprovers);

            // Helpdesk
            case "Helpdesk":
                // Return the email
                return AppConfig.Configuration.helpdesk ? [AppConfig.Configuration.helpdesk] : [];

            // Sponsor
            case "Sponsor":
                // Get the email from the login name
                let email = Common.getEmail(item.AppSponsor ? item.AppSponsor.Name : "");

                // Return the email
                return email ? [email] : [];

            // Sponsor's Group
            case "SponsorsGroup":
                return AppSecurity.AppWeb.getEmailsForGroup(Strings.Groups.Sponsors);
        }
    }

    // Gets the value
    private static getValue(key: string, item: IAppItem, status: string, deploySiteUrl: string) {
        // Default the value
        let value = item[key];

        // See if the value is complex
        if (key.indexOf('.') > 0) {
            // See if it's the user name
            if (key == "User.Name") {
                // Set the value
                value = ContextInfo.userDisplayName;
            }
            // Else, see if it's the current user
            else if (key.indexOf("CurrentUser") == 0) {
                // Get the sub-keys and update the value
                let keys = key.split('.');
                value = AppSecurity.AppWeb.CurrentUser;

                // Parse the properties of complex fields
                for (let j = 1; j < keys.length; j++) {
                    // Set the value
                    value = value && value[keys[j]] ? value[keys[j]] : null;
                }
            }
            else {
                // Get the sub-keys and update the value
                let keys = key.split('.');
                value = item[keys[0]]

                // Parse the properties of complex fields
                for (let j = 1; j < keys.length; j++) {
                    // Set the value
                    value = value && value[keys[j]] ? value[keys[j]] : null;
                }
            }

            // Return the value
            return value;
        }

        // Return the value
        switch (key) {
            // See if this is the status field
            case "AppStatus":
                // Set the value
                value = status;
                break;

            // See if it's the current user
            case "CurrentUser":
                // Default to the user name
                value = AppSecurity.AppWeb.CurrentUser.Title;
                break;

            // Url to the site the app was deployed to
            case "DeploySiteUrl":
                // Set the value
                value = "<a href='" + deploySiteUrl + "'>" + deploySiteUrl + "</a>";
                break;

            // Url to the site the app was deployed to
            case "DeploySiteUrlText":
                // Set the value
                value = deploySiteUrl;
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

            // Url to the test site
            case "TestSiteUrl":
                // Set the url
                let testSiteUrl = (AppConfig.Configuration.appCatalogUrl || ContextInfo.webAbsoluteUrl) + "/" + item.AppProductID;
                if (testSiteUrl.toLowerCase().indexOf("http") != 0) {
                    // Prepend the origin
                    testSiteUrl = window.location.origin + testSiteUrl;
                }

                // Set the value
                value = "<a href='" + testSiteUrl + "'>" + testSiteUrl + "</a>";
                break;

            // Url to the test site
            case "TestSiteUrlText":
                // Set the value
                value = (AppConfig.Configuration.appCatalogUrl || ContextInfo.webAbsoluteUrl) + "/" + item.AppProductID;
                if (value.toLowerCase().indexOf("http") != 0) {
                    // Prepend the origin
                    value = window.location.origin + value;
                }
                break;

            // Item Metadata
            default:
                let values = [];

                // See if we are dealing with a collection
                // Otherwise, revert to the item
                let results = value && value.results ? value.results : [item];
                value = "";

                // Parse the collection
                for (let i = 0; i < results.length; i++) {
                    let result = results[i];
                    if (result) {
                        // Parse the properties of complex fields
                        let keys = key.split('.');
                        value = result[keys[0]];
                        for (let j = 1; j < keys.length; j++) {
                            // Set the value
                            value = value[keys[j]] ? value[keys[j]] : null;
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
    static rejectEmail(item: IAppItem, comments: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Get the app developers
            let To = [];
            let owners = item.AppDevelopers && item.AppDevelopers.results ? item.AppDevelopers.results : [];
            for (let i = 0; i < owners.length; i++) {
                // Append the email
                let email = Common.getEmail(owners[i].Name);
                email ? To.push(email) : null;
            }

            // Ensure emails exist
            if (To.length > 0) {
                // Display a loading dialog
                LoadingDialog.setHeader("Sending Notification");
                LoadingDialog.setBody("This dialog will close after the notification is sent.");
                LoadingDialog.show();

                // Set the email properties
                let emailProps = {
                    To,
                    CC: [AppSecurity.AppWeb.CurrentUser.Email],
                    Subject: "App '" + item.Title + "' Sent Back",
                    Body: "App Developers,<br /><br />The '" + item.Title +
                        "' app has been sent back based on the comments below.<br /><br />" + comments
                } as Types.IEmail;

                // Send an email
                Utility(Strings.SourceUrl).sendEmail(emailProps).execute(
                    () => {
                        // Close the loading dialog
                        LoadingDialog.hide();

                        // Resolve the request
                        resolve();
                    },
                    ex => {
                        // Log
                        ErrorDialog.logError(`The rejectEmail method failed with properties: ${JSON.stringify(emailProps)}`);

                        // Log the error
                        ErrorDialog.show("Sending Email", "There was an error sending the reject notification email.", ex);
                    }
                );
            } else {
                // Resolve the request
                resolve();
            }
        });
    }

    // Sends an email when an app is deployed
    static sendAppDeployedEmail(item: IAppItem, siteUrl: string): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // See if a configuration exists
            if (AppConfig.Configuration.appNotifications && AppConfig.Configuration.appNotifications.deploy) {
                // Send the email
                this.sendEmail(AppConfig.Configuration.appNotifications.deploy, item, false, true, siteUrl).then(resolve);
            } else {
                // Resolve the requst
                resolve();
            }
        });
    }

    // Sends an email when an app is upgraded in the test site
    static sendAppTestSiteUpgradedEmail(item: IAppItem): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // See if a configuration exists
            if (AppConfig.Configuration.appNotifications && AppConfig.Configuration.appNotifications.test) {
                // Send the email
                this.sendEmail(AppConfig.Configuration.appNotifications.test, item, false, true).then(resolve);
            } else {
                // Resolve the requst
                resolve();
            }
        });
    }

    // Sends an email when an app is upgraded
    static sendAppUpgradedEmail(item: IAppItem): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // See if a configuration exists
            if (AppConfig.Configuration.appNotifications && AppConfig.Configuration.appNotifications.upgrade) {
                // Send the email
                this.sendEmail(AppConfig.Configuration.appNotifications.upgrade, item, false, true).then(resolve);
            } else {
                // Resolve the requst
                resolve();
            }
        });
    }

    // Sends an email notification based on the status
    // Need a new property "approval/submission" to determine what type?
    // Maybe send in the status configuration and we can figure it out on our own?
    // New property for the app notification configuration required.
    static sendEmail(notificationCfgs: IEmail[] = [], item: IAppItem, isApproval: boolean = true, isAppAction: boolean = false, deploySiteUrl: string = ""): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            let emailsToSend = 0;

            // Ensure notifications exist
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

                // See if we are doing an approval
                if (isApproval) {
                    // Ensure we are doing an approval
                    if (notificationCfg.approval != true) { continue; }
                }
                // Else, see if this isn't an app action
                else if (!isAppAction) {
                    // Ensure we are doing a submission
                    if (notificationCfg.submission != true) { continue; }
                }

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

                // Parse the email content and find [Key] instances w/in it.
                let Body = notificationCfg.content || "";
                let startIdx = Body.indexOf("[");
                let endIdx = Body.indexOf("]");
                while (startIdx >= 0 && endIdx > startIdx) {
                    // Get the key value
                    let key = Body.substring(startIdx + 1, endIdx);
                    let value = this.getValue(key, item, status, deploySiteUrl);

                    // Replace the value
                    Body = Body.substring(0, startIdx) + value + Body.substring(endIdx + 1);

                    // Find the next instance of it
                    startIdx = Body.indexOf("[");
                    endIdx = Body.indexOf("]");
                }

                // Parse the email content and find [Key] instances w/in it.
                let Subject = notificationCfg.subject || "";
                startIdx = Subject.indexOf("[");
                endIdx = Subject.indexOf("]");
                while (startIdx >= 0 && endIdx > startIdx) {
                    // Get the key value
                    let key = Subject.substring(startIdx + 1, endIdx);
                    let value = this.getValue(key, item, status, deploySiteUrl);

                    // Replace the value
                    Subject = Subject.substring(0, startIdx) + value + Subject.substring(endIdx + 1);

                    // Find the next instance of it
                    startIdx = Subject.indexOf("[");
                    endIdx = Subject.indexOf("]");
                }

                // Ensure emails exist
                if (To.length > 0 || CC.length > 0) {
                    // Increment the send counter
                    emailsToSend++;

                    // Display a loading dialog
                    LoadingDialog.setHeader("Sending Notification");
                    LoadingDialog.setBody("This dialog will close after the notification is sent.");
                    LoadingDialog.show();

                    // Add the current user to the CC
                    CC.push(AppSecurity.AppWeb.CurrentUser.Email);

                    // Set the email properties
                    let emailProps = {
                        To, CC,
                        Body,
                        Subject
                    };

                    // Send an email
                    Utility(Strings.SourceUrl).sendEmail(emailProps).execute(
                        () => {
                            // Close the loading dialog
                            LoadingDialog.hide();

                            // Resolve the request
                            resolve();
                        },
                        ex => {
                            // Log
                            ErrorDialog.logError(`The sendEmail method failed with properties: ${JSON.stringify(emailProps)}`);

                            // Log the error
                            ErrorDialog.show("Sending Email", "There was an error sending the notification email.", ex);
                        }
                    );
                } else {
                    // Resolve the request
                    resolve();
                }
            }

            // See if no emails were sent
            if (emailsToSend == 0) {
                // Resolve the request
                resolve();
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

            // Parse the email content and find [Key] instances w/in it.
            let Body = body || "";
            let startIdx = Body.indexOf("[");
            let endIdx = Body.indexOf("]");
            while (startIdx >= 0 && endIdx > startIdx) {
                // Get the key value
                let key = Body.substring(startIdx + 1, endIdx);
                let value = this.getValue(key, item, item.AppStatus, "");

                // Replace the value
                Body = Body.substring(0, startIdx) + value + Body.substring(endIdx + 1);

                // Find the next instance of it
                startIdx = Body.indexOf("[");
                endIdx = Body.indexOf("]");
            }

            // Parse the email content and find [Key] instances w/in it.
            let Subject = subject || "";
            startIdx = Subject.indexOf("[");
            endIdx = Subject.indexOf("]");
            while (startIdx >= 0 && endIdx > startIdx) {
                // Get the key value
                let key = Subject.substring(startIdx + 1, endIdx);
                let value = this.getValue(key, item, item.AppStatus, "");

                // Replace the value
                Subject = Subject.substring(0, startIdx) + value + Subject.substring(endIdx + 1);

                // Find the next instance of it
                startIdx = Subject.indexOf("[");
                endIdx = Subject.indexOf("]");
            }

            // Ensure emails exist
            if (To.length > 0) {
                // Display a loading dialog
                LoadingDialog.setHeader("Sending Notification");
                LoadingDialog.setBody("This dialog will close after the notification is sent.");
                LoadingDialog.show();

                // Set the email properties
                let emailProps = {
                    To,
                    CC: [AppSecurity.AppWeb.CurrentUser.Email],
                    Body,
                    Subject,
                } as Types.IEmail;

                // Send an email
                Utility(Strings.SourceUrl).sendEmail(emailProps).execute(
                    () => {
                        // Close the loading dialog
                        LoadingDialog.hide();

                        // Resolve the request
                        resolve();
                    },
                    ex => {
                        // Log
                        ErrorDialog.logError(`The sendNotification method failed with properties: ${JSON.stringify(emailProps)}`);

                        // Log the error
                        ErrorDialog.show("Sending Email", "There was an error sending the user's notification to the system.", ex);
                    }
                );
            } else {
                // Resolve the request
                resolve();
            }
        });
    }
}