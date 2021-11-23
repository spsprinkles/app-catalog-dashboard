import { Components, ContextInfo, Helper, List, SPTypes, Types, Utility, Web } from "gd-sprest-bs";
import { appIndicator } from "gd-sprest-bs/build/icons/svgs/appIndicator";
import { chatSquareDots } from "gd-sprest-bs/build/icons/svgs/chatSquareDots";
import { pencilSquare } from "gd-sprest-bs/build/icons/svgs/pencilSquare";
import { trash } from "gd-sprest-bs/build/icons/svgs/trash";
import * as jQuery from "jquery";
import Strings from "../strings";
import Toast from "./toast";

/**
 * Table
 */
export class Table {
    private _el: HTMLElement = null;
    private _datatable = null;
    private _formUrl: string = null;
    private _search: string = null;
    private _filter: string = null;
    private _table: Components.ITable = null;

    // Constructor
    constructor(el: HTMLElement) {
        // Save the parameters
        this._el = el;

        // Load the data
        this.load().then(() => {
            // Render the table
            this.render();
        });
    }

    // Clears the component
    private clear() {
        // Clear the elements
        while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }
    }

    // Load the form url
    private load() {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the form for the list
            List(Strings.Lists.Apps).Forms().query({
                Filter: "FormType eq " + SPTypes.PageType.DisplayForm
            }).execute(forms => {
                // Save the form url
                this._formUrl = forms.results[0].ServerRelativeUrl;

                // Resolve the promise
                resolve(null);
            }, reject);
        });
    }

    // Load the items
    private loadItems(): PromiseLike<Array<any>> {
        // Return a promise
        return new Promise((resolve, reject) => {
            let filter = null;
            if (this._filter) {
                switch (this._filter)  {
                    case "Show apps from others":
                        filter = null;
                        break;
                    case "Show all apps for review":
                        filter = "DevAppStatus eq 'In Review'";
                        break;
                }
            }
            else //Default
                filter = "Owners eq '" + ContextInfo.userId + "'";

            // Load the list items
            List(Strings.Lists.Apps).Items().query({
                Select: ["Id", "FileLeafRef", "CheckoutUser/Title", "AppThumbnailURL", "AuthorId", "OwnersId", "DevAppStatus", 'Title', "AppVersion", "AppPublisher", "Owners", "IsAppPackageEnabled", "PublishApp", "Owners/Id", "Owners/EMail"],
                Expand: ["CheckoutUser", "Owners"],
                //Filter: this._filter == "Show apps from others" ? null : "Owners eq '" + ContextInfo.userId + "'"
                Filter: filter
            }).execute(items => {
                // Resolve the promise
                resolve(items.results);
            });
        });
    }

    // Render the table
    private render() {
        // Clear the component
        this.clear();
        var that = this; //Save reference

        // Render a table
        this._table = Components.Table({
            el: this._el,
            className: "table-striped",
            columns: [
                {
                    name: "Options",
                    //title: "",
                    /*onRenderHeader: (el) => {
                        // Remove the header name
                        //if (el.innerText == "Info") { el.innerText = ""; }
                        el.innerText = "";
                    },*/
                    onRenderCell: (el, col, item: Types.SP.ListItem) => {
                        var canEdit = (item["AuthorId"] == ContextInfo.userId || (item["OwnersId"] ? item["OwnersId"].results.indexOf(ContextInfo.userId) != -1 : false));
                        
                        // Create a div to wrap the button icons in, giving it a white background
                        let elDiv = document.createElement("div");
                        elDiv.classList.add("bg-white");
                        elDiv.classList.add("d-inline-flex");
                        elDiv.classList.add("rounded");
                        elDiv.style.marginRight = "8px";
                        el.appendChild(elDiv);

                        // Render a button
                        Components.Button({
                            el: elDiv,
                            //text: item.Title,
                            iconType: pencilSquare,
                            iconSize: 20, //28
                            title: "Edit Properties",
                            isDisabled: !canEdit,
                            type: Components.ButtonTypes.OutlineSecondary,
                            onClick: () => {
                                // Show the display form
                                Helper.SP.ModalDialog.showModalDialog({
                                    //title: "View Item",
                                    url: this._formUrl + "?ID=" + item.Id,
                                    dialogReturnValueCallback: result => {
                                        // See if the item was updated
                                        if (result == SPTypes.ModalDialogResult.OK) {
                                            //Ensure file is checked in: May not be if user closed/uncompleted form after initial app upload
                                            if (item["CheckoutUser"].Title) {
                                                List(Strings.Lists.Apps).Items().query({
                                                    Select: ["FileLeafRef"],
                                                    Filter: "Id eq " + item.Id
                                                }).execute((obj) => {
                                                    if (obj.results.length == 0) {
                                                        that.render();
                                                        return;
                                                    }
                                                    Web().getFolderByServerRelativeUrl(ContextInfo.webServerRelativeUrl + "/DeveloperApps")
                                                    .Files().getByUrl(ContextInfo.webServerRelativeUrl + "/DeveloperApps/" + obj.results[0]["FileLeafRef"])
                                                    .checkIn("New file", 1).execute(function() { //Minor:0, Major:1, OverwriteCheckIn:2
                                                        that.render();
                                                    //Error
                                                    }, function (obj) {
                                                        var response = JSON.parse(obj.response);
                                                        //if (response.error.code == "-2146232832, Microsoft.SharePoint.SPException") {
                                                        //"Cannot check in due to missing required field values");
                                                        that.render();
                                                    });
                                                //Error finding item
                                                }, (obj) => {
                                                    that.render();
                                                });
                                            }
                                            else // Refresh the page
                                                that.render();
                                        }
                                    }
                                });
                            }
                        });

                        let elDiv2 = document.createElement("div");
                        elDiv2.classList.add("bg-white");
                        elDiv2.classList.add("d-inline-flex");
                        elDiv2.classList.add("rounded");
                        elDiv2.style.marginRight = "8px";
                        el.appendChild(elDiv2);

                        // Render a button
                        Components.Button({
                            el: elDiv2,
                            //text: item.Title,
                            iconType: appIndicator, //Award ChatSquareDots
                            iconSize: 20,
                            isDisabled: (canEdit && item["DevAppStatus"] == "In Testing" ? false : true),
                            title: "Submit for review",
                            type: Components.ButtonTypes.OutlineSecondary,
                            onClick: () => {
                                var div = document.createElement("div");
                                div.innerHTML = 'Are you sure you want to submit your app for review?' +
                                    '<br /><br/><div style="text-align:right;"><button class="ms-emphasis" onclick="$REST.Helper.SP.ModalDialog.commonModalDialogClose($REST.SPTypes.ModalDialogResult.OK)">Confirm</button></div>';
                                Helper.SP.ModalDialog.showModalDialog({
                                    title: "Submit for Review",
                                    html: div,
                                    width: 420,
                                    dialogReturnValueCallback: (results, args) => {
                                        if (results == SPTypes.ModalDialogResult.OK)
                                            List(Strings.Lists.Apps).Items(item.Id).update({
                                                "DevAppStatus": "In Review"
                                            }).execute((obj) =>{
                                                that.refresh();
                                                Toast.showMsg("/_layouts/images/mysitetitle.png", "App Submitted for Review", "Other developers can now review your app");
                                                //Email
                                                Utility().sendEmail({
                                                    To: ["App Developers"],
                                                    Subject: "App '" + item.Title + "' submitted for review",
                                                    Body: "App Developers,<br /><br />The '" + item.Title + "' app has been submitted for review by " + ContextInfo.userDisplayName + ". Please take some time to test this app and submit an assessment/review using the App Dashboard."
                                                }).execute();
                                            });
                                    }
                                });
                            }
                        });

                        let elDiv3 = document.createElement("div");
                        elDiv3.classList.add("bg-white");
                        elDiv3.classList.add("d-inline-flex");
                        elDiv3.classList.add("rounded");
                        elDiv3.style.marginRight = "8px";
                        el.appendChild(elDiv3);

                        // Render a button
                        Components.Button({
                            el: elDiv3,
                            //text: item.Title,
                            iconType: chatSquareDots,
                            iconSize: 20,
                            isDisabled: (item["AuthorId"] == ContextInfo.userId ? true : (item["DevAppStatus"] != "In Review" ? true : false)),
                            title: "Review this app",
                            type: Components.ButtonTypes.OutlineSecondary,
                            onClick: () => {
                                List(Strings.Lists.Assessments).Items().query({
                                    Select: ["Id"],
                                    Filter: "RelatedApp eq " + item.Id.toString() + " and Author eq " + ContextInfo.userId
                                }).execute((obj) => {
                                    var assessment = obj.results[0];
                                    var endUrl = "NewForm.aspx?relatedId=" + item.Id.toString();
                                    if (assessment)
                                        endUrl = "EditForm.aspx?ID=" + assessment.Id.toString();

                                    Helper.SP.ModalDialog.showModalDialog({
                                        url: ContextInfo.webAbsoluteUrl + "/Lists/DevAppAssessments/" + endUrl,
                                        dialogReturnValueCallback: (results, args) => {
                                            if (results == SPTypes.ModalDialogResult.OK) {
                                                Toast.showMsg("/_layouts/images/NoteBoard_32x32.png", "App Review Submitted", "Thank you for your assessment of this app");
                                            }
                                        }
                                    });
                                });
                            }
                        });

                        let elDiv4 = document.createElement("div");
                        elDiv4.classList.add("bg-white");
                        elDiv4.classList.add("d-inline-flex");
                        elDiv4.classList.add("rounded");
                        el.appendChild(elDiv4);

                        // Render a button
                        Components.Button({
                            el: elDiv4,
                            //text: item.Title,
                            iconType: trash,
                            iconSize: 20,
                            isDisabled: (canEdit ? false : true),
                            title: "Delete app/solution",
                            type: Components.ButtonTypes.OutlineSecondary,
                            onClick: () => {
                                //Helper.SP.ModalDialog.showWaitScreenWithNoClose("Please wait", "Verifying details...");
                                List("Apps for SharePoint").Items().query({
                                    Select: ["IsAppPackageEnabled", "HasUniqueRoleAssignments"],
                                    Filter: "FileLeafRef eq '" + item["FileLeafRef"] + "'"
                                }).execute((obj) => {
                                    let deleteItem = true; //default
                                    let dotIndex = item["FileLeafRef"].lastIndexOf(".");
                                    let fileExt = item["FileLeafRef"].substring(dotIndex); //includes the .
                                    if (fileExt.toLowerCase() == ".sppkg" && obj.results[0] && obj.results[0].HasUniqueRoleAssignments == false)
                                        deleteItem = false;

                                    let div = document.createElement("div");
                                    if (deleteItem)
                                        div.innerHTML = 'Are you sure you want to delete this app? It will also be deleted from the tenant App Catalog.' +
                                            '<br /><br/><div style="text-align:right;"><button class="ms-emphasis" onclick="$REST.Helper.SP.ModalDialog.commonModalDialogClose($REST.SPTypes.ModalDialogResult.OK, true)">Confirm</button></div>';
                                    else
                                    div.innerHTML = 'This SP framework solution has already been published and cannot be deleted because it would break existing deployed instances. Would you like to disable the app instead?' +
                                        '<br /><br/><div style="text-align:right;"><button class="ms-emphasis" onclick="$REST.Helper.SP.ModalDialog.commonModalDialogClose($REST.SPTypes.ModalDialogResult.OK, false)">Disable App</button></div>';
                                    //Helper.SP.ModalDialog.commonModalDialogClose();
                                    Helper.SP.ModalDialog.showModalDialog({
                                        title: "Delete App",
                                        html: div,
                                        width: 500,
                                        dialogReturnValueCallback: (result, args) => {
                                            // See if the item was updated
                                            if (result == SPTypes.ModalDialogResult.OK) {
                                                let data = {} as any;
                                                if (args == true)
                                                    data.DevAppStatus = "Deleting...";
                                                else
                                                    data.IsAppPackageEnabled = false;

                                                List(Strings.Lists.Apps).Items(item.Id).update(data).execute(function() {
                                                    if (args == true)
                                                        Toast.showMsg("/_layouts/images/ManageWorkflow32.png", "Deleting App", "Your app will be deleted shortly");
                                                    that.render();
                                                },
                                                function (obj) {
                                                    var response = JSON.parse(obj.response);
                                                    alert("Error: " + response.error.message.value);
                                                });
                                            }
                                        }
                                    });
                                });
                            }
                        });
                    }
                },
                {
                    name: "AppThumbnailURL",
                    title: "Icon",
                    onRenderCell: (el, col, item: Types.SP.ListItem) => {
                        var appIcon = item[col.name] || {};
                        //var blankImageUrl = "/_layouts/15/images/blank.gif";
                        //el.innerHTML = '<img class="bg-sharepoint" src="' + (appIcon.Url || blankImageUrl) + '" style="width:42px; height:42px;">';
                        if (appIcon.Url)
                            el.innerHTML = '<img class="bg-sharepoint" src="' + appIcon.Url + '" style="width:42px; height:42px;">';
                        else
                            el.innerHTML = "";
                    }
                },
                {
                    name: "Title", //LinkFilename, FileLeafRef
                    title: "App Title",
                    onRenderCell: (el, col, item: Types.SP.ListItem) => {
                        el.innerHTML = "";
                        if (item["CheckoutUser"].Title)
                            el.innerHTML = '<img src="/_layouts/15/images/checkoutoverlay.gif" title="Checked Out To: ' + item["CheckoutUser"].Title + '" /> ';
                        el.innerHTML += item[col.name];
                    }
                },
                {
                    name: "AppVersion",
                    title: "Version"
                },
                {
                    name: "DevAppStatus",
                    title: "App Status"
                },
                {
                    name: "AppPublisher",
                    title: "Publisher"
                },
                {
                    name: "Owners",
                    onRenderCell: (el, col, item: Types.SP.ListItem) => {
                        var owners = item[col.name] && item[col.name].results || [];
                        var strOwners = "";
                        owners.forEach(owner => {
                            strOwners += owner.EMail.split("@")[0] + "; ";
                        });
                        el.innerText = strOwners;
                    }
                },
                {
                    name: "IsAppPackageEnabled",
                    title: "Enabled",
                    onRenderCell: (el, col, item: Types.SP.ListItem) => {
                        el.innerText = item[col.name] ? "Yes" : "No";
                    }
                }/*,
                {
                    name: "PublishApp",
                    title: "Workflow",
                    onRenderCell: (el, col, item: Types.SP.ListItem) => {
                        var wfUrl = item[col.name] || {};
                        if (wfUrl.Url)
                            el.innerHTML = '<a href="' + wfUrl.Url + '" target="_blank">' + wfUrl.Description + '</a>';
                    }
                }*/
            ]
        });

        // Load the items
        this.loadItems().then(items => {
            // Add the table rows
            this._table.addRows(items);

            // Apply the datatable
            this._datatable = jQuery(this._table.el).DataTable();

            // See if there is a search value set
            if (this._search) {
                // Search the table
                this._datatable.search(this._search).draw();
            }
        });
    }

    /**
     * Public Interface
     */

    // Applies a filter to the table
    applyFilter(filter: string) {
        // Render the table
        this._filter = filter;
        this.render();
    }

    // Refresh the table
    refresh() {
        this.render();
    }

    // Search the table
    search(search: string) {
        // Set the search value
        this._search = search;

        // Search the table
        this._datatable.search(search).draw();
    }
}