import { Dashboard, LoadingDialog } from "dattatable";
import { Components } from "gd-sprest-bs";
import { arrowClockwise } from "gd-sprest-bs/build/icons/svgs/arrowClockwise";
import { chatSquareDots } from "gd-sprest-bs/build/icons/svgs/chatSquareDots";
import { fileEarmarkArrowUp } from "gd-sprest-bs/build/icons/svgs/fileEarmarkArrowUp";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import { layoutTextWindow } from "gd-sprest-bs/build/icons/svgs/layoutTextWindow";
import { questionLg } from "gd-sprest-bs/build/icons/svgs/questionLg";
import { AppActions } from "./appActions";
import { AppConfig } from "./appCfg";
import { AppDashboard } from "./appDashboard";
import { AppForms } from "./appForms";
import { AppSecurity } from "./appSecurity";
import { UserAgreement } from "./userAgreement";
import * as Common from "./common";
import * as jQuery from "jquery";
import { DataSource, IAppItem } from "./ds";
import Strings from "./strings";

/**
 * Application View
 * Displays a high-level overview of all the applications.
 */
export class AppView {
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

        // Refresh the data source
        DataSource.refresh().then(() => {
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
        if (AppSecurity.IsOwner) {
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
                            DataSource.InstallRequired(null, true);
                        }
                    },
                    {
                        text: "Manage Approver Group",
                        onClick: () => {
                            // Show the group in a new tab
                            window.open(AppSecurity.ApproverUrl, "_blank");
                        }
                    },
                    {
                        text: "Manage Developer Group",
                        onClick: () => {
                            // Show the group in a new tab
                            window.open(AppSecurity.DevUrl, "_blank");
                        }
                    },
                    {
                        text: "Manage Sponsor Group",
                        onClick: () => {
                            // Show the group in a new tab
                            window.open(AppSecurity.SponsorUrl, "_blank");
                        }
                    }
                ]
            });
        }

        // See if the help url exists
        if (AppConfig.Configuration.helpPageUrl) {
            // Add the item
            navLinks.push({
                className: "btn-outline-light mx-2 ps-1 pt-1",
                iconSize: 24,
                iconType: questionLg,
                isButton: true,
                text: "Help",
                onClick: () => {
                    // Display in a new tab
                    window.open(AppConfig.Configuration.helpPageUrl, "_blank");
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
                            this._dashboard.filter(3, value);
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
                    brand.appendChild(Common.getAppIcon());
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
                                        // Ensure the user is not an approver or developer
                                        if (!AppSecurity.IsApprover && !AppSecurity.IsDeveloper) {
                                            // Show the user agreement
                                            new UserAgreement();
                                        } else {
                                            // Upload the package file
                                            AppActions.upload(item => {
                                                // Refresh the dashboard
                                                this.refresh();

                                                // Display the edit form
                                                this._forms.edit(item.Id, () => {
                                                    // Refresh the dashboard
                                                    this.refresh();
                                                });
                                            });
                                        }
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
                    pageLength: AppConfig.Configuration.paging,
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
                                el.innerHTML = '<img src="' + item.AppThumbnailURL.Url + '" height="32px" width="32px" title="' + item.Title + '">';
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
                        name: "",
                        title: "Status",
                        onRenderCell: (el, column, item: IAppItem) => {
                            // Set the status
                            el.innerHTML = Common.appStatus(item);

                            // Ensure it doesn't contain the url already
                            if (item.AppSiteDeployments && item.AppSiteDeployments.length > 0) {
                                // Append a badge indicating that the app is deployed
                                let badge = Components.Badge({
                                    el,
                                    className: "ms-1",
                                    isPill: true,
                                    content: "Deployed",
                                    type: Components.BadgeTypes.Success
                                });

                                // Create the popover element
                                let elSiteInfo = document.createElement("div");
                                elSiteInfo.classList.add("m-2");
                                elSiteInfo.innerHTML = item.AppSiteDeployments;

                                // Set the popover
                                Components.Popover({
                                    target: badge.el,
                                    options: {
                                        content: elSiteInfo
                                    }
                                });
                            }
                        }
                    },
                    {
                        name: "AppSponsor",
                        title: "App Sponsor",
                        onRenderCell: (el, column, item: IAppItem) => {
                            // Display the sponsor
                            el.innerText = (item.AppSponsor ? item.AppSponsor.Title : null) || "";
                        }
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
                                strOwners.push(owner.Title);
                            });

                            // Display the owners
                            el.innerHTML = strOwners.join("<br/>");
                        }
                    },
                    {
                        name: "",
                        title: "Actions",
                        className: "text-end",
                        onRenderCell: (el, column, item: IAppItem) => {
                            let tooltips: Components.ITooltipProps[] = [];

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
                                            new AppDashboard(this._elAppDetails, this._el, item.Id);

                                            // Hide the apps
                                            this._el.classList.add("d-none");

                                            // Show the details
                                            this._elAppDetails.classList.remove("d-none");

                                            // Hide the loading dialog
                                            LoadingDialog.hide();
                                        });
                                    }
                                }
                            });

                            // Render the edit properties button
                            tooltips.push({
                                content: "View app properties",
                                btnProps: {
                                    text: "View",
                                    className: "p-1",
                                    iconClassName: "me-1",
                                    iconSize: 20,
                                    iconType: chatSquareDots,
                                    isSmall: true,
                                    type: Components.ButtonTypes.OutlineSecondary,
                                    onClick: () => {
                                        // Display the view form
                                        this._forms.display(item.Id);
                                    }
                                }
                            });

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