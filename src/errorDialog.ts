import { Components } from "gd-sprest-bs";
import { CanvasForm, LoadingDialog, Modal } from "dattatable";
import Strings from "./strings";

/**
 * Error Dialog
 */
export class ErrorDialog {
    private static _log = null;
    static get Log() { return this._log; }
    static set Log(value) { this._log = value; }

    private static _scope = null;
    static get Scope() { return this._scope; }
    static set Scope(value) { this._scope = value; }

    // Logs the message to the console
    static logError(message: string) {
        // See if the log exists
        if (this.Log) {
            // Log to the dev dashboard
            this.Log.error(Strings.ProjectName, message, this.Scope);
        } else {
            // Log to the console
            console.error(message);
        }
    }

    // Logs the message to the console
    static logInfo(message: string, ...args) {
        // See if the log exists
        if (this.Log) {
            // Log to the dev dashboard
            this.Log.info(Strings.ProjectName, message, this.Scope);
        } else {
            // Log to the console
            console.info(message);
        }
    }

    // Shows an error
    static show(title: string, message: string, ex?: any) {
        // Log the exception
        this.logError(`[${Strings.ProjectName}][${title}] ${message} ${ex}`)

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
        Modal.show();
    }
}