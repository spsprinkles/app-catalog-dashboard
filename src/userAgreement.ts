import { LoadingDialog, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
import { AppConfig } from "./appCfg";
import { AppSecurity } from "./appSecurity";

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

                    // Show a loading dialog
                    LoadingDialog.setHeader("Adding User");
                    LoadingDialog.setBody("Adding the user to the developer's security group. This will close after the user is added.");
                    LoadingDialog.show();

                    // Ensure the user is added to the developer's group
                    AppSecurity.addDeveloper().then(() => {
                        // Close the dialog
                        LoadingDialog.hide();

                        // Resolve the request
                        onAgree();
                    });
                }
            },
            placement: Components.TooltipPlacements.Left
        }).el);

        // Show the modal
        Modal.show();
    }
}