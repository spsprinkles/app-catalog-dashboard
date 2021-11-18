import { Dashboard, LoadingDialog } from "dattatable";
import { Components, ContextInfo } from "gd-sprest-bs";
import { appIndicator } from "gd-sprest-bs/build/icons/svgs/appIndicator";
import { chatSquareDots } from "gd-sprest-bs/build/icons/svgs/chatSquareDots";
import { columnsGap } from "gd-sprest-bs/build/icons/svgs/columnsGap";
import { pencilSquare } from "gd-sprest-bs/build/icons/svgs/pencilSquare";
import { trash } from "gd-sprest-bs/build/icons/svgs/trash";
import { AppForms } from "./itemForms";
import * as jQuery from "jquery";
import { DataSource, IAppItem } from "./ds";
import Strings from "./strings";


/**
 * Main Application
 */
export class App {
    private _dashboard: Dashboard = null;
    private _el: HTMLElement = null;
    private _forms: AppForms = null;

    // Constructor
    constructor(el: HTMLElement) {
        // Set the global variables
        this._el = el;
        this._forms = new AppForms();

        // Render the dashboard
        this.render();
    }

    // Refreshes the dashboard
    private refresh() {
        // Show a loading dialog
        LoadingDialog.setHeader("Refreshing the Data");
        LoadingDialog.setBody("This will close after the data is loaded.");

        // Load the events
        DataSource.init().then(() => {
            // Clear the element
            while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

            // Render the dashboard
            this.render();

            // Hide the dialog
            LoadingDialog.hide();
        });
    }

    // Renders the dashboard
    private render() {
        // Create the dashboard
        this._dashboard = new Dashboard({
            el: this._el,
            hideHeader: true,
            useModal: true,
            filters: {
                items: [
                    {
                        header: "Event Status",
                        items: DataSource.StatusFilters,
                        onFilter: (value: string) => {
                            // Filter the dashboard
                            this._dashboard.filter(4, value);
                        },
                    },
                ],
            },
            navigation: {
                title: columnsGap(32).outerHTML + '<span class="ms-3">' + Strings.ProjectName + "</span>",
                items: [
                    {
                        className: "btn-outline-light",
                        isButton: true,
                        text: "Add/Update App",
                        onClick: () => {
                            // Upload the package file
                            this._forms.upload(() => {
                                // Refresh the dashboard
                                this.refresh();
                            });
                        }
                    },
                    {
                        className: "ms-2 btn-outline-light",
                        isButton: true,
                        text: "Refresh"
                    }
                ],
            },
            footer: {
                itemsEnd: [
                    {
                        text: "v" + Strings.Version,
                    },
                ],
            },
            table: {
                rows: DataSource.Items,
                dtProps: {
                    dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                    columnDefs: [
                        {
                            targets: [0],
                            orderable: false,
                            searchable: false
                        }
                    ],
                    // Add some classes to the dataTable elements
                    drawCallback: function () {
                        jQuery(".table", this._table).removeClass("no-footer");
                        jQuery(".table", this._table).addClass("tbl-footer");
                        jQuery(".table", this._table).addClass("table-striped");
                        jQuery(".table thead th", this._table).addClass("align-middle");
                        jQuery(".table tbody td", this._table).addClass("align-middle");
                        jQuery(".dataTables_info", this._table).addClass("text-center");
                        jQuery(".dataTables_length", this._table).addClass("pt-2");
                        jQuery(".dataTables_paginate", this._table).addClass("pt-03");
                    },
                    // Sort descending by Start Date
                    order: [[2, "asc"]],
                    language: {
                        emptyTable: "No apps were found",
                    },
                },
                columns: [
                    {
                        name: "",
                        title: "Options",
                        onRenderCell: (el, column, item: IAppItem) => {
                            let canEdit = (item["AuthorId"] == ContextInfo.userId || (item["OwnersId"] ? item["OwnersId"].results.indexOf(ContextInfo.userId) != -1 : false));
                            Components.ButtonGroup({
                                el,
                                buttons: [
                                    {
                                        text: "Edit Properties",
                                        iconSize: 20,
                                        iconType: pencilSquare,
                                        isSmall: true,
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Display the edit form
                                            this._forms.edit(item.Id, () => {
                                                // Refresh the table
                                                this.refresh();
                                            });
                                        }
                                    },
                                    {
                                        text: "Submit for Review",
                                        iconSize: 20,
                                        iconType: appIndicator,
                                        isDisabled: canEdit && item.DevAppStatus != "In Review" ? false : true,
                                        isSmall: true,
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Display the submit form
                                            this._forms.submit(item, () => {
                                                // Refresh the table
                                                this.refresh();
                                            });
                                        }
                                    },
                                    {
                                        text: "Review this App",
                                        iconSize: 20,
                                        iconType: chatSquareDots,
                                        isDisabled: item.DevAppStatus != "In Review" ? true : !canEdit,
                                        isSmall: true,
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Display the review form
                                            this._forms.review(item, () => {
                                                // Refresh the table
                                                this.refresh();
                                            });
                                        }
                                    },
                                    {
                                        text: "Delete App/Solution",
                                        iconSize: 20,
                                        iconType: trash,
                                        isDisabled: !canEdit,
                                        isSmall: true,
                                        type: Components.ButtonTypes.OutlineDanger,
                                        onClick: () => {
                                            // Display the delete form
                                            this._forms.delete(item, () => {
                                                // Refresh the table
                                                this.refresh();
                                            });
                                        }
                                    },
                                ]
                            });
                        }
                    },
                    {
                        name: "",
                        title: "Icon",
                        onRenderCell: (el, column, item: IAppItem) => {
                            // Ensure a url exists
                            if (item.AppThumbnailURL && item.AppThumbnailURL.Url) {
                                // Render the link
                                el.innerHTML = '<img class="bg-sharepoint" src="' + item.AppThumbnailURL.Url + '" style="width:42px; height:42px;">';
                            }
                        }
                    },
                    {
                        name: "Title",
                        title: "App Title",
                        onRenderCell: (el, column, item: IAppItem) => {
                            // See if the item is checked out
                            if (item.CheckoutUser && item.CheckoutUser.Title) {
                                // Prepend the image
                                el.innerHTML = '<img src="/_layouts/15/images/checkoutoverlay.gif" title="Checked Out To: ' + item.CheckoutUser.Title + '" /> ' + el.innerHTML;
                            }
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
                        name: "",
                        title: "Owners",
                        onRenderCell: (el, column, item: IAppItem) => {
                            var owners = item.Owners && item.Owners.results || [];

                            // Parse the owners
                            var strOwners = [];
                            owners.forEach(owner => {
                                // Append the email
                                strOwners.push(owner.EMail.split("@")[0]);
                            });

                            // Display the owners
                            el.innerText = strOwners.join("; ");
                        }
                    },
                    {
                        name: "",
                        title: "Enabled",
                        onRenderCell: (el, column, item: IAppItem) => {
                            // Set the text
                            el.innerText = item.IsAppPackageEnabled ? "Yes" : "No";
                        }
                    }
                ]
            }
        });
    }
}