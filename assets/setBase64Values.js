/**
 * Run this in the console to set the base64 values for the images.
 * This is needed for the new base64 fields added.
 */

const { $REST } = require("gd-sprest-bs");
const { DattaTable } = require("dattatable");

/**
 * Copy everything under this line and paste in the dev console for the dashboard.
 */

// Show a loading dialog
DattaTable.LoadingDialog.setHeader("Setting Base64 Values")
DattaTable.LoadingDialog.setBody("Reading the files....");
DattaTable.LoadingDialog.show();

// Log
console.log("Starting the script...");

// Get the items
$REST.List("Developer Apps").Items().query({
    Filter: "ContentType eq 'App'"
}).execute(
    // Success
    (items) => {
        let counter = 0;

        // Parse the items
        $REST.Helper.Executor(items.results, item => {
            // Return a promise
            return new Promise(resolve => {
                let hasValues = false;
                let values = {};

                // Update the dialog
                DattaTable.LoadingDialog.setBody("Updating item " + (++counter) + " of " + items.results.length);

                // Read the image files
                $REST.Helper.Executor([
                    "AppImageURL1", "AppImageURL2", "AppImageURL3",
                    "AppImageURL4", "AppImageURL5", "AppThumbnailURL"
                ], fieldName => {
                    // Return a promise
                    return new Promise(resolve => {
                        let urlValue = item[fieldName];

                        // See if an image exists
                        if (urlValue && urlValue.Url) {
                            // Read the item
                            let request = new XMLHttpRequest();
                            request.open("GET", urlValue.Url, true);
                            request.responseType = "blob";
                            request.onload = () => {
                                // Convert the image to base64
                                let reader = new FileReader();
                                reader.readAsDataURL(request.response);
                                reader.onload = e => {
                                    // Update the values
                                    values[fieldName + "Base64"] = e.target.result;

                                    // Set the flag
                                    hasValues = true;

                                    // Check the next item
                                    resolve();
                                }
                            };
                            request.onerror = () => {
                                // Log
                                console.error("Error getting the image from: " + urlValue.Url);

                                // Check the next item
                                resolve();
                            }

                            // Execute the request
                            request.send();
                        } else {
                            // Check the next item
                            resolve();
                        }
                    }, resolve);
                }).then(() => {
                    // See if values were set
                    if (hasValues) {
                        // Update the item
                        item.update(values).execute(() => {
                            // Log
                            console.log("App [" + item.Title + "] was updated successfully.", values);

                            // Check the next item
                            resolve();
                        }, () => {
                            // Log
                            console.error("App [" + item.Title + "] failed to be updated.", values);

                            // Check the next item
                            resolve();
                        });
                    } else {
                        // Check the next item
                        resolve();
                    }
                });
            });
        }).then(() => {
            // Close the dialog
            DattaTable.LoadingDialog.hide();

            // Log
            console.log("Completed the script...");
        });
    },

    // Error
    () => {
        // Hide the dialog
        DattaTable.LoadingDialog.hide();

        // Show the error
        DattaTable.Modal.clear();
        DattaTable.Modal.setHeader("Error Reading Items");
        DattaTable.Modal.setBody("There was an error reading the items...");
        DattaTable.Modal.show();
    }
);
