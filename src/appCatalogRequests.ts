import { DataTable, LoadingDialog, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
import * as jQuery from "jquery";
import { DataSource, IAppCatalogRequestItem } from "./ds";

/**
 * App Catalog Requests
 */
export class AppCatalogRequests {
    // Displays the user's app catalog requests
    static viewAppCatalogRequests() {
        // Set the header
        Modal.setHeader("App Catalog Requests");
        Modal.setType(Components.ModalTypes.Large);

        // Display a loading dialog
        LoadingDialog.setHeader("Loading Data");
        LoadingDialog.setBody("This will close after the requests have been loaded...");
        LoadingDialog.show();

        // Load the data
        DataSource.loadAppCatalogRequests().then(items => {
            // Display the dashboard
            new DataTable({
                el: Modal.BodyElement,
                rows: items,
                dtProps: {
                    dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                    columnDefs: [
                        {
                            targets: [3],
                            orderable: false,
                            searchable: false
                        }
                    ],
                    createdRow: function (row, data, index) {
                        jQuery('td', row).addClass('align-middle');
                    },
                    // Add some classes to the dataTable elements
                    drawCallback: function (settings) {
                        let api = new jQuery.fn.dataTable.Api(settings) as any;
                        let div = api.table().container() as HTMLDivElement;
                        let table = api.table().node() as HTMLTableElement;
                        div.querySelector(".dataTables_info").classList.add("text-center");
                        div.querySelector(".dataTables_length").classList.add("pt-2");
                        div.querySelector(".dataTables_paginate").classList.add("pt-03");
                        table.classList.remove("no-footer");
                        table.classList.add("tbl-footer");
                        table.classList.add("table-striped");
                    },
                    headerCallback: function (thead, data, start, end, display) {
                        jQuery('th', thead).addClass('align-middle');
                    },
                    // Sort descending by Start Date
                    order: [[0, "asc"]],
                    language: {
                        emptyTable: "No requests were found",
                    },
                },
                columns: [
                    {
                        name: "",
                        title: "Site",
                        onRenderCell: (el, col, item: IAppCatalogRequestItem) => {
                            // Display the url
                            if (item.SiteCollectionUrl) {
                                // Display the url
                                el.innerHTML = item.SiteCollectionUrl.Url;
                            }
                        }
                    },
                    {
                        name: "RequestStatus",
                        title: "Status",
                    },
                    {
                        name: "RequestNotes",
                        title: "Notes",
                    },
                    {
                        name: "",
                        title: "",
                        onRenderCell: (el, col, item: IAppCatalogRequestItem) => {
                            // Render a delete button
                            let btn = Components.Button({
                                el,
                                text: "Delete",
                                type: Components.ButtonTypes.OutlineDanger,
                                isSmall: true,
                                onClick: () => {
                                    // Disable the button
                                    btn.disable();

                                    // Delete the item
                                    item.delete().execute(() => {
                                        // Update the button text
                                        btn.setText("Deleted");
                                    });
                                }
                            });
                        }
                    }
                ]
            });

            // Hide the dialog
            LoadingDialog.hide();
        });

        // Render the footer
        Modal.setFooter(Components.Button({
            text: "Close",
            type: Components.ButtonTypes.OutlineSuccess,
            onClick: () => {
                // Close the modal
                Modal.hide();
            }
        }).el);

        // Show the modal
        Modal.show();
    }
}