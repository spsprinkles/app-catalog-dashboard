import { LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo } from "gd-sprest-bs";
import { AppConfig } from "./appCfg";
import { AppSecurity } from "./appSecurity";
import Strings from "./strings";

/**
 * User Agreement
 */
export class UserAgreement {
    // Constructor
    constructor(onAgree: () => void) {
        // Ensure the user agreement exists
        if (AppConfig.Configuration.userAgreement) {
            // Render the modal
            this.render(onAgree);
        } else {
            // Clear the modal
            Modal.clear();

            // Set the header and body
            Modal.setHeader("User Agreement");
            Modal.setBody("The user agreement hasn't been setup. Please contact your administrator to update the configuration.");
            Modal.show();
        }
    }

    // Renders the user agreement modal
    private render(onAgree: () => void) {
        // Clear the modal
        Modal.clear();

        // Set the options
        Modal.setAutoClose(false);
        Modal.setScrollable(true);
        Modal.setType(Components.ModalTypes.Large)

        // Set the header and body
        Modal.setHeader("User Agreement");
        Modal.setBody(AppConfig.Configuration.userAgreement);

        // Set the footer
        Modal.setFooter(Components.Tooltip({
            content: "Clicking 'Agree' will allow you to submit applications.",
            btnProps: {
                text: "Agree",
                type: Components.ButtonTypes.OutlinePrimary,
                onClick: () => {
                    // Close the dialog
                    Modal.hide();

                    // See if the user is already in the developer's group
                    if (AppSecurity.AppWeb.isInGroup(Strings.Groups.Developers)) {
                        // Call the event
                        onAgree();
                    } else {
                        // See if the user is in the developer's group
                        // Show a loading dialog
                        LoadingDialog.setHeader("Adding User");
                        LoadingDialog.setBody("Adding the user to the developer's security group. This will close after the user is added.");
                        LoadingDialog.show();

                        // Add the user to the group
                        AppSecurity.AppWeb.addUserToGroup(Strings.Groups.Developers, ContextInfo.userId).then(() => {
                            // Close the dialog
                            LoadingDialog.hide();

                            // Call the event
                            onAgree();
                        });
                    }
                }
            },
            placement: Components.TooltipPlacements.Left
        }).el);

        // Show the modal
        Modal.show();
    }
}