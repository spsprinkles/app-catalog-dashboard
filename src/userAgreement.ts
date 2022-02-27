import { LoadingDialog, Modal } from "dattatable";
import { Components, ContextInfo } from "gd-sprest-bs";
import { AppSecurity } from "./appSecurity";
import { DataSource } from "./ds";

/**
 * User Agreement
 */
export class UserAgreement {
    // Constructor
    constructor() {
        // Load the user agreement
        DataSource.loadUserAgreement().then(() => {
            // Render the modal
            this.render();
        });
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
        Modal.setCloseButtonVisibility(false);
        Modal.setScrollable(true);
        Modal.setType(Components.ModalTypes.Large)

        // Set the header and body
        Modal.setHeader("Developer Agreement");
        Modal.setBody(DataSource.UserAgreement || "Developer agreement has not been identified by the client");

        // Set the footer
        Modal.setFooter(Components.Tooltip({
            content: "Clicking 'Agree' will allow you to continue",
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