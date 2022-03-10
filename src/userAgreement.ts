import { LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo } from "gd-sprest-bs";
import { AppConfig } from "./appCfg";
import { AppSecurity } from "./appSecurity";

/**
 * User Agreement
 */
export class UserAgreement {
    // Constructor
    constructor() {
        // Ensure the user agreement exists
        if (AppConfig.Configuration.userAgreement) {
            // Render the modal
            this.render();
        } else {
            // Clear the modal
            Modal.clear();

            // Set the header and body
            Modal.setHeader("User Agreement");
            Modal.setBody("The user agreement hasn't been setup. Please contact your administrator to update the configuration.");
            Modal.show();
        }
    }

    // Add the user to the developer group
    private addUserToGroup() {
        // Show a loading dialog
        LoadingDialog.setHeader("Adding User");
        LoadingDialog.setBody("Adding the user to the developer's security group. This will close after the user is added.");
        LoadingDialog.show();

        // Add the user
        AppSecurity.DevGroup.Users.addUserById(ContextInfo.userId).execute(() => {
            // Close the dialog
            LoadingDialog.hide();

            // Refresh the page
            window.location.reload();
        });
    }

    // Renders the user agreement modal
    private render() {
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
                    // Add the user to the group
                    this.addUserToGroup();
                }
            },
            placement: Components.TooltipPlacements.Left
        }).el);

        // Show the modal
        Modal.show();
    }
}