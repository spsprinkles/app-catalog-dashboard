import { ItemForm } from "dattatable";
import { Helper, List } from "gd-sprest-bs";
import Strings from "./strings";

/**
 * App Item Forms
 */
export class AppForms {
    // Constructor
    constructor() {
        // Set the list name
        ItemForm.ListName = Strings.Lists.Apps;
    }

    // Edit form
    edit(itemId: number, onUpdate: () => void) {
        // Show the item form
        ItemForm.edit({
            itemId,
            onSetHeader: el => {
                // Update the header
                el.innerHTML = "New App"
            },
            onCreateEditForm: props => {
                // Exclude fields
                props.excludeFields = ["DevAppStatus"];

                // Update the field
                props.onControlRendering = (ctrl, field) => {
                    // See if this is a read-only field
                    if (field.InternalName == "FileLeafRef") {
                        // Make it read-only
                        ctrl.isReadonly = true;
                    }
                }

                return props;
            },
            onUpdate: () => {
                // Call the update event
                onUpdate();
            }
        });
    }

    // Submit form
    submit(itemId: number, onUpdate: () => void) {
        // TODO
    }

    // Upload form
    upload(onUpdate: () => void) {
        // Show the upload file dialog
        Helper.ListForm.showFileDialog().then(file => {
            // Ensure this is an spfx package
            if (file.name.toLowerCase().endsWith(".sppkg")) {
                // Upload the file
                List(Strings.Lists.Apps).RootFolder().Files().add(file.name, true, file.data).execute(
                    // Success
                    file => {
                        // Get the item id
                        file.ListItemAllFields().execute(item => {
                            // Display the edit form
                            this.edit(item.Id, onUpdate);
                        });
                    },

                    // Error
                    () => {
                        // TODO
                    }
                )
            }
        });
    }
}