import { Components } from "gd-sprest-bs";
import { CanvasForm, LoadingDialog, Modal } from "dattatable";
import Strings from "./strings";

/**
 * Error Dialog
 */
export class ErrorDialog {
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
        Modal.setHeader(title);
        Modal.setBody(message);
        Modal.setType(Components.ModalTypes.Large);

        // Show the modal
        Modal.show();
    }
}