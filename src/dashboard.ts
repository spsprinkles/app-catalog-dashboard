import { Dashboard, LoadingDialog } from "dattatable";
import { Components } from "gd-sprest-bs";
import { appIndicator } from "gd-sprest-bs/build/icons/svgs/appIndicator";
import { arrowClockwise } from "gd-sprest-bs/build/icons/svgs/arrowClockwise";
import { chatSquareDots } from "gd-sprest-bs/build/icons/svgs/chatSquareDots";
import { columnsGap } from "gd-sprest-bs/build/icons/svgs/columnsGap";
import { fileEarmarkArrowUp } from "gd-sprest-bs/build/icons/svgs/fileEarmarkArrowUp";
import { pencilSquare } from "gd-sprest-bs/build/icons/svgs/pencilSquare";
import * as Common from "./common";
import { AppForms } from "./itemForms";
import * as jQuery from "jquery";
import { DataSource, IAppItem } from "./ds";
import Strings from "./strings";


/**
 * Dashboard
 */
export class AppDashboard {
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
        // See if this is an approver
        let navLinks: Components.INavbarItem[] = null;
        if (DataSource.IsApprover) {
            // Set the admin buttons
            navLinks = [{
                text: "Admin",
                items: [
                    {
                        text: "Manage Dev Group",
                        onClick: () => {
                            // Show the group in a new tab
                            window.open(DataSource.ApproverUrl, "_blank");
                        }
                    },
                    {
                        text: "Manage Approver Group",
                        onClick: () => {
                            // Show the group in a new tab
                            window.open(DataSource.DevUrl, "_blank");
                        }
                    },
                    {
                        text: "Manage App",
                        onClick: () => {
                            // Show the install modal
                            DataSource.InstallRequired(true);
                        }
                    }
                ]
            }];
        }

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
                itemsEnd: navLinks,
                onRendering: (props) => {
                    let brandName = document.createElement("div");
                    brandName.className = "ms-75";
                    brandName.append(Strings.ProjectName);
                    let brand = document.createElement("div");
                    brand.className = "d-flex";
                    brand.appendChild(columnsGap());
                    brand.appendChild(brandName);
                    props.brand = brand;
                },
                onRendered: (el) => {
                    el.querySelector("nav div.container-fluid").classList.add("ps-3");
                    el.querySelector("nav div.container-fluid a.navbar-brand").classList.add("pe-none");
                },
                onSearchRendered: (el) => {
                    el.setAttribute("placeholder", "Search this dashboard");
                },
                showFilter: false
            },
            subNavigation: {
                itemsEnd: [
                    {
                        text: "Add/Update App",
                        onRender: (el, item) => {
                            // Clear the existing button
                            el.innerHTML = "";
                            // Create a span to wrap the icon in
                            let span = document.createElement("span");
                            span.className = "bg-white d-inline-flex ms-2 rounded";
                            el.appendChild(span);

                            // Render a tooltip
                            Components.Tooltip({
                                el: span,
                                content: item.text,
                                btnProps: {
                                    // Render the icon button
                                    className: "p-1",
                                    iconType: fileEarmarkArrowUp,
                                    iconSize: 24,
                                    type: Components.ButtonTypes.OutlineSecondary,
                                    onClick: () => {
                                        // Upload the package file
                                        this._forms.upload(() => {
                                            // Refresh the dashboard
                                            this.refresh();
                                        });
                                    }
                                },
                            });
                        }
                    },
                    {
                        text: "Refresh",
                        onRender: (el, item) => {
                            // Clear the existing button
                            el.innerHTML = "";
                            // Create a span to wrap the icon in
                            let span = document.createElement("span");
                            span.className = "bg-white d-inline-flex ms-2 rounded";
                            el.appendChild(span);

                            // Render a tooltip
                            Components.Tooltip({
                                el: span,
                                content: item.text,
                                btnProps: {
                                    // Render the icon button
                                    className: "p-1",
                                    iconType: arrowClockwise,
                                    iconSize: 24,
                                    type: Components.ButtonTypes.OutlineSecondary,
                                    onClick: () => {
                                        // Refresh the dashboard
                                        this.refresh();
                                    }
                                },
                            });
                        }
                    }
                ],
                onRendered: (el) => {
                    let filter = el.querySelector(".filter-icon");
                    if (filter) {
                        let filterItem = document.createElement("li");
                        filterItem.className = "nav-item";
                        filterItem.appendChild(filter);
                        let ul = el.querySelector("#navbar_content ul:last-child");
                        let li = ul.querySelector("li:last-child");
                        ul.insertBefore(filterItem, li);
                    }
                },
                showFilter: true,
            },
            footer: {
                itemsEnd: [
                    {
                        className: "pe-none",
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
                    createdRow: function (row, data, index) {
                        jQuery('td', row).addClass('align-middle');
                    },
                    // Add some classes to the dataTable elements
                    drawCallback: function (settings) {
                        let api = new jQuery.fn.dataTable.Api(settings) as any;
                        jQuery(api.context[0].nTable).removeClass('no-footer');
                        jQuery(api.context[0].nTable).addClass('tbl-footer');
                        jQuery(api.context[0].nTable).addClass('table-striped');
                        jQuery(api.context[0].nTableWrapper).find('.dataTables_info').addClass('text-center');
                        jQuery(api.context[0].nTableWrapper).find('.dataTables_length').addClass('pt-2');
                        jQuery(api.context[0].nTableWrapper).find('.dataTables_paginate').addClass('pt-03');
                    },
                    headerCallback: function (thead, data, start, end, display) {
                        jQuery('th', thead).addClass('align-middle');
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
                            let tooltips: Components.ITooltipProps[] = [];

                            // Determine if the user can edit
                            let canEdit = Common.canEdit(item);

                            // Render the view button to redirect the user to the document set dashboard
                            tooltips.push({
                                content: "Go to the app dashboard",
                                btnProps: {
                                    text: "View",
                                    iconSize: 20,
                                    iconType: appIndicator,
                                    isSmall: true,
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // Redirect to the docset item
                                        window.open(Common.generateDocSetUrl(item), "_self");
                                    }
                                }
                            });

                            // See if the user can edit
                            if (canEdit) {
                                // Render the edit properties button
                                tooltips.push({
                                    content: "Edit the app properties",
                                    btnProps: {
                                        text: "Edit",
                                        iconSize: 20,
                                        iconType: pencilSquare,
                                        isSmall: true,
                                        type: Components.ButtonTypes.OutlineSecondary,
                                        onClick: () => {
                                            // Display the edit form
                                            this._forms.edit(item.Id, () => {
                                                // Refresh the table
                                                this.refresh();
                                            });
                                        }
                                    }
                                });
                            }

                            // See if this is an owner
                            if ((item.DevAppStatus == "Draft" || item.DevAppStatus == "Requires Attention") && canEdit) {
                                // Submit button
                                tooltips.push({
                                    content: "Submit the app for review",
                                    btnProps: {
                                        text: "Submit",
                                        iconSize: 20,
                                        iconType: appIndicator,
                                        isDisabled: !canEdit,
                                        isSmall: true,
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Display the submit form
                                            this._forms.submit(item, () => {
                                                // Refresh the table
                                                this.refresh();
                                            });
                                        }
                                    }
                                });
                            }

                            // Ensure this was not submitted by the user
                            if (item.DevAppStatus.indexOf("Review") > 0 &&
                                (!Common.isSubmitter(item) || DataSource.IsApprover)) {
                                // Review button
                                tooltips.push({
                                    content: "Start/Continue an assessment of the app.",
                                    btnProps: {
                                        text: "Review",
                                        iconSize: 20,
                                        iconType: chatSquareDots,
                                        isDisabled: !canEdit,
                                        isSmall: true,
                                        type: Components.ButtonTypes.OutlinePrimary,
                                        onClick: () => {
                                            // Display the review form
                                            this._forms.review(item, () => {
                                                // Refresh the table
                                                this.refresh();
                                            });
                                        }
                                    }
                                });

                                // Reject button
                                tooltips.push({
                                    content: "Sends the request back to the developer(s).",
                                    btnProps: {
                                        text: "Reject",
                                        iconSize: 20,
                                        iconType: chatSquareDots,
                                        isSmall: true,
                                        type: Components.ButtonTypes.OutlineDanger,
                                        onClick: () => {
                                            // Display the reject form
                                            this._forms.reject(item, () => {
                                                // Refresh the page
                                                window.location.reload();
                                            });
                                        }
                                    }
                                });
                            }

                            // Render the tooltips
                            Components.TooltipGroup({
                                el,
                                tooltips
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
                        title: "App Status",
                        onRenderCell: (el, column, item: IAppItem) => {
                            let lineBreak = document.createElement("br");
                            el.appendChild(lineBreak);

                            // See if this app is deployed in the catalog
                            let app = DataSource.getAppById(item.AppProductID);
                            if (app) {
                                // See if it's deployed
                                if (app.Deployed) {
                                    // Render a badge
                                    Components.Badge({
                                        el,
                                        content: "Deployed",
                                        isPill: true,
                                        type: Components.BadgeTypes.Primary
                                    });
                                } else {
                                    // Render a badge
                                    Components.Badge({
                                        el,
                                        content: "Not Deployed",
                                        isPill: true,
                                        type: Components.BadgeTypes.Secondary
                                    });
                                }
                            } else {
                                // Render a badge
                                Components.Badge({
                                    el,
                                    content: "Not in App Catalog",
                                    isPill: true,
                                    type: Components.BadgeTypes.Info
                                });
                            }
                        }
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