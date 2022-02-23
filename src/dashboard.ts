import { Dashboard, LoadingDialog, Modal } from "dattatable";
import { Components } from "gd-sprest-bs";
import { appIndicator } from "gd-sprest-bs/build/icons/svgs/appIndicator";
import { arrowClockwise } from "gd-sprest-bs/build/icons/svgs/arrowClockwise";
import { chatSquareDots } from "gd-sprest-bs/build/icons/svgs/chatSquareDots";
import { columnsGap } from "gd-sprest-bs/build/icons/svgs/columnsGap";
import { fileEarmarkArrowUp } from "gd-sprest-bs/build/icons/svgs/fileEarmarkArrowUp";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import { layoutTextWindow } from "gd-sprest-bs/build/icons/svgs/layoutTextWindow";
import { pencilSquare } from "gd-sprest-bs/build/icons/svgs/pencilSquare";
import { questionLg } from "gd-sprest-bs/build/icons/svgs/questionLg";
import { App } from "./app";
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
    private _elAppDetails: HTMLElement = null;
    private _forms: AppForms = null;

    // Constructor
    constructor(el: HTMLElement, elAppDetails: HTMLElement) {
        // Set the global variables
        this._el = el;
        this._elAppDetails = elAppDetails;
        this._forms = new AppForms();

        // Render the dashboard
        this.render();
    }

    // Refreshes the dashboard
    private refresh() {
        // Show a loading dialog
        LoadingDialog.setHeader("Refreshing the Data");
        LoadingDialog.setBody("This will close after the data is loaded.");

        // Load the apps
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
        // See if this is an owner
        let navLinks: Components.INavbarItem[] = [];
        if (DataSource.IsOwner) {
            // Set the admin buttons
            navLinks.push({
                className: "btn-outline-light ms-2 pt-1",
                text: "Settings",
                iconClassName: "me-1",
                iconSize: 24,
                iconType: gearWideConnected,
                isButton: true,
                items: [
                    {
                        text: "Manage App Configuration",
                        onClick: () => {
                            // Show the install modal
                            DataSource.InstallRequired(true);
                        }
                    },
                    {
                        text: "Manage Approver Group",
                        onClick: () => {
                            // Show the group in a new tab
                            window.open(DataSource.ApproverUrl, "_blank");
                        }
                    },
                    {
                        text: "Manage Developer Group",
                        onClick: () => {
                            // Show the group in a new tab
                            window.open(DataSource.DevUrl, "_blank");
                        }
                    }
                ]
            });
        }

        // See if the help url exists
        if (DataSource.Configuration.helpPageUrl) {
            // Add the item
            navLinks.push({
                className: "btn-outline-light mx-2 ps-1 pt-1",
                iconSize: 24,
                iconType: questionLg,
                isButton: true,
                text: "Help",
                onClick: () => {
                    // Display in a new tab
                    window.open(DataSource.Configuration.helpPageUrl, "_blank");
                }
            });
        }

        // Create the dashboard
        this._dashboard = new Dashboard({
            el: this._el,
            hideHeader: true,
            useModal: true,
            filters: {
                items: [
                    {
                        header: "App Status",
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
                // Add the branding icon & text
                onRendering: (props) => {
                    // Set the class names
                    props.className = "bg-sharepoint navbar-expand rounded-top";

                    // Set the brand
                    let brand = document.createElement("div");
                    brand.className = "d-flex";
                    brand.appendChild(columnsGap());
                    brand.append(Strings.ProjectName);
                    brand.querySelector("svg").classList.add("me-75");
                    props.brand = brand;
                },
                // Adjust the brand alignment
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
                // Move the filter icon in front of the last icon
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
                            targets: [0, 6],
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
                    order: [[1, "asc"]],
                    language: {
                        emptyTable: "No apps were found",
                    },
                },
                columns: [
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
                        title: "App Name",
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
                        name: "AppStatus",
                        title: "Status",
                        onRenderCell: (el, column, item: IAppItem) => {
                            // See if the site app catalog exists
                            if (DataSource.SiteCollectionAppCatalogExists) {
                                // Append a new line
                                let lineBreak = document.createElement("br");
                                el.appendChild(lineBreak);

                                // See if it's deployed
                                let app = DataSource.getSiteCollectionAppById(item.AppProductID);
                                if (app) {
                                    // Render a badge
                                    Components.Badge({
                                        el,
                                        content: "Deployed to Site Catalog",
                                        isPill: true,
                                        type: Components.BadgeTypes.Primary
                                    });
                                } else {
                                    // Render a badge
                                    Components.Badge({
                                        el,
                                        content: "Not in Site Catalog",
                                        isPill: true,
                                        type: Components.BadgeTypes.Secondary
                                    });
                                }
                            }

                            // Append a new line
                            let lineBreak = document.createElement("br");
                            el.appendChild(lineBreak);

                            // See if this app is deployed in the tenant app catalog
                            let app = DataSource.getTenantAppById(item.AppProductID);
                            if (app) {
                                // See if it's deployed
                                if (app.Deployed) {
                                    // Render a badge
                                    Components.Badge({
                                        el,
                                        content: "Deployed to Tenant Catalog",
                                        isPill: true,
                                        type: Components.BadgeTypes.Primary
                                    });
                                } else {
                                    // Render a badge
                                    Components.Badge({
                                        el,
                                        content: "Not Deployed to Tenant Catalog",
                                        isPill: true,
                                        type: Components.BadgeTypes.Secondary
                                    });
                                }
                            } else {
                                // Render a badge
                                Components.Badge({
                                    el,
                                    className: "text-dark",
                                    content: "Not in Tenant App Catalog",
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
                        title: "AppDevelopers",
                        onRenderCell: (el, column, item: IAppItem) => {
                            var owners = item.AppDevelopers && item.AppDevelopers.results || [];

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
                        title: "Actions",
                        className: "text-end",
                        onRenderCell: (el, column, item: IAppItem) => {
                            let tooltips: Components.ITooltipProps[] = [];

                            // Determine if the user can edit
                            let canEdit = Common.canEdit(item);

                            // Render the view button to redirect the user to the document set dashboard
                            tooltips.push({
                                content: "View app details",
                                btnProps: {
                                    data: item,
                                    text: "Details",
                                    className: "p-1",
                                    iconClassName: "me-1",
                                    iconSize: 20,
                                    iconType: layoutTextWindow,
                                    isSmall: true,
                                    type: Components.ButtonTypes.OutlinePrimary,
                                    onClick: () => {
                                        // Redirect to the docset item
                                        //window.open(Common.generateDocSetUrl(item), "_self");

                                        // Show a loading dialog
                                        LoadingDialog.setHeader("Loading Application Information");
                                        LoadingDialog.setBody("This will close after the data is loaded...");
                                        LoadingDialog.show();

                                        // Load the doc set item
                                        DataSource.loadDocSet(item.Id).then(() => {
                                            // Clear the details
                                            while (this._elAppDetails.firstChild) { this._elAppDetails.removeChild(this._elAppDetails.firstChild); }

                                            // Set the body
                                            new App(this._elAppDetails, this._el, item.Id);

                                            // Hide the apps
                                            this._el.style.display = "none";

                                            // Show the details
                                            this._elAppDetails.style.display = "";

                                            // Hide the loading dialog
                                            LoadingDialog.hide();
                                        });
                                    }
                                }
                            });

                            // See if the user can edit
                            if (canEdit) {
                                // Render the edit properties button
                                tooltips.push({
                                    content: "Edit app properties",
                                    btnProps: {
                                        text: "Edit",
                                        className: "p-1",
                                        iconClassName: "me-1",
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
                            if ((item.AppStatus == "Draft" || item.AppStatus == "Requires Attention") && canEdit) {
                                // Submit button
                                tooltips.push({
                                    content: "Submit app for review",
                                    btnProps: {
                                        text: "Submit",
                                        className: "p-1",
                                        iconClassName: "me-1",
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
                            if (item.AppStatus.indexOf("Review") > 0 &&
                                (!Common.isSubmitter(item) || DataSource.IsApprover)) {
                                // Review button
                                tooltips.push({
                                    content: "Start/Continue an app assessment",
                                    btnProps: {
                                        text: "Review",
                                        className: "p-1",
                                        iconClassName: "me-1",
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
                                    content: "Send the app back to the developer(s)",
                                    btnProps: {
                                        text: "Reject",
                                        className: "p-1",
                                        iconClassName: "me-1",
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
                    }
                ]
            }
        });
    }
}