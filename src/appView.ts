import { Dashboard, LoadingDialog } from "dattatable";
import { Components, ContextInfo, ThemeManager } from "gd-sprest-bs";
import { chatSquareDots } from "gd-sprest-bs/build/icons/svgs/chatSquareDots";
import { fileEarmarkArrowUp } from "gd-sprest-bs/build/icons/svgs/fileEarmarkArrowUp";
import { filterSquare } from "gd-sprest-bs/build/icons/svgs/filterSquare";
import { gearWideConnected } from "gd-sprest-bs/build/icons/svgs/gearWideConnected";
import { layoutTextWindow } from "gd-sprest-bs/build/icons/svgs/layoutTextWindow";
import { personBoundingBox } from "gd-sprest-bs/build/icons/svgs/personBoundingBox";
import { questionLg } from "gd-sprest-bs/build/icons/svgs/questionLg";
import { ticketDetailed } from "gd-sprest-bs/build/icons/svgs/ticketDetailed";
import * as jQuery from "jquery";
import { AppActions } from "./appActions";
import { AppCatalogRequests } from "./appCatalogRequests";
import { AppConfig } from "./appCfg";
import { AppDashboard } from "./appDashboard";
import { AppForms } from "./appForms";
import { AppInstall } from "./appInstall";
import { AppSecurity } from "./appSecurity";
import { UserAgreement } from "./userAgreement";
import * as Common from "./common";
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
    private _elFilterButtons: HTMLButtonElement[] = [];
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

    // Determines if an app belongs to the user
    private isMyApp(item: IAppItem) {
        // See if the user is a sponsor
        if (item.AppSponsorId == ContextInfo.userId) { return true; }

        // Parse the developers
        let userIds = item.AppDevelopersId ? item.AppDevelopersId.results : [];
        for (let i = 0; i < userIds.length; i++) {
            // See if the user is a developer
            if (userIds[i] == ContextInfo.userId) { return true; }
        }

        // Not the user's app
        return false;
    }

    // Refreshes the dashboard
    private refresh(itemId: number): PromiseLike<void> {
        // Return a promise
        return new Promise(resolve => {
            // Refresh the data source
            DataSource.refresh(itemId).then(() => {
                // Clear the element
                while (this._el.firstChild) { this._el.removeChild(this._el.firstChild); }

                // Resolve the request
                resolve();
            });
        });
    }

    // Renders the dashboard
    private render() {
        let navLinks: Components.INavbarItem[] = [];

        // See if the app catalog is on the same site as the app
        let isSameWeb = AppConfig.Configuration.appCatalogUrl.toLowerCase() == Strings.SourceUrl.toLowerCase();

        // See if the help url exists
        if (AppConfig.Configuration.helpPageUrl) {
            // Add the item
            navLinks.push({
                text: "Help",
                onRender: (el, item) => {
                    // Clear the existing button
                    el.innerHTML = "";

                    // Render a tooltip
                    Components.Tooltip({
                        el,
                        content: "Get " + item.text,
                        type: Components.TooltipTypes.LightBorder,
                        btnProps: {
                            // Render the icon button
                            className: "btn-icon btn-outline-light me-2 p-1 pe-2 py-1",
                            iconSize: 22,
                            iconType: questionLg,
                            text: item.text,
                            type: Components.ButtonTypes.OutlineLight,
                            onClick: () => {
                                // Display in a new tab
                                window.open(AppConfig.Configuration.helpPageUrl, "_blank");
                            }
                        }
                    });
                }
            });
        }

        // See if this is a developer/sponsor/owner
        if (DataSource.HasAppCatalogRequests) {
            // Add the requests button
            navLinks.push({
                text: "My Requests",
                onRender: (el, item) => {
                    // Clear the existing button
                    el.innerHTML = "";

                    // Render a tooltip
                    Components.Tooltip({
                        el,
                        content: "View " + item.text,
                        type: Components.TooltipTypes.LightBorder,
                        btnProps: {
                            // Render the icon button
                            className: "btn-icon btn-outline-light me-2 p-2 py-1",
                            iconClassName: "me-2",
                            iconSize: 22,
                            iconType: ticketDetailed,
                            text: item.text,
                            type: Components.ButtonTypes.OutlineLight,
                            onClick: () => {
                                // Display in app catalog requests
                                AppCatalogRequests.viewAppCatalogRequests();
                            }
                        }
                    });
                }
            });
        }

        // See if this is an owner
        if (AppSecurity.AppWeb.IsAdmin || AppSecurity.AppWeb.IsOwner) {
            // Set the admin buttons
            navLinks.push({
                className: "btn-icon btn-outline-light me-2 p-2 py-1",
                text: "Settings",
                iconSize: 22,
                iconType: gearWideConnected,
                isButton: true,
                items: [
                    {
                        text: "App Configuration",
                        onClick: () => {
                            // Show the install modal
                            AppInstall.InstallRequired(null, true);
                        }
                    },
                    {
                        text: "Local Approver Group",
                        onClick: () => {
                            // Show the group in a new tab
                            window.open(AppSecurity.AppWeb.getUrlForGroup(Strings.Groups.Approvers), "_blank");
                        }
                    },
                    {
                        text: "Local Developer Group",
                        onClick: () => {
                            // Show the group in a new tab
                            window.open(AppSecurity.AppWeb.getUrlForGroup(Strings.Groups.Developers), "_blank");
                        }
                    },
                    {
                        text: "Local Final Approver Group",
                        onClick: () => {
                            // Show the group in a new tab
                            window.open(AppSecurity.AppWeb.getUrlForGroup(Strings.Groups.FinalApprovers), "_blank");
                        }
                    },
                    {
                        text: "Local Sponsor Group",
                        onClick: () => {
                            // Show the group in a new tab
                            window.open(AppSecurity.AppWeb.getUrlForGroup(Strings.Groups.Sponsors), "_blank");
                        }
                    },
                    {
                        text: "Manage Audit Log",
                        onClick: () => {
                            // Show the audit log list
                            window.open(DataSource.AuditLog.List.ListUrl), "_blank";
                        }
                    }
                ]
            });

            if (!isSameWeb) {
                // Add the links to manage remote groups
                navLinks[navLinks.length - 1].items.push(
                    {
                        text: "Remote Approver Group",
                        onClick: () => {
                            // Show the group in a new tab
                            window.open(AppSecurity.AppCatalogWeb.getUrlForGroup(Strings.Groups.Approvers), "_blank");
                        }
                    },
                    {
                        text: "Remote Developer Group",
                        onClick: () => {
                            // Show the group in a new tab
                            window.open(AppSecurity.AppCatalogWeb.getUrlForGroup(Strings.Groups.Developers), "_blank");
                        }
                    },
                    {
                        text: "Remote Final Approver Group",
                        onClick: () => {
                            // Show the group in a new tab
                            window.open(AppSecurity.AppCatalogWeb.getUrlForGroup(Strings.Groups.FinalApprovers), "_blank");
                        }
                    },
                    {
                        text: "Remote Sponsor Group",
                        onClick: () => {
                            // Show the group in a new tab
                            window.open(AppSecurity.AppCatalogWeb.getUrlForGroup(Strings.Groups.Sponsors), "_blank");
                        }
                    }
                );
            }
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
                    {
                        header: "App Deployment",
                        items: [
                            {
                                label: "Site Collection Deployed",
                                type: Components.CheckboxGroupTypes.Switch
                            },
                            {
                                label: "Tenant Deployed",
                                type: Components.CheckboxGroupTypes.Switch
                            }
                        ],
                        onFilter: (value: string) => {
                            // Filter the dashboard
                            this._dashboard.filter(4, value);
                        },
                    }
                ],
                onRendered: (el) => {
                    // Add the dark class if theme is inverted
                    if (ThemeManager.IsInverted) {
                        el.querySelectorAll("div.form-check.form-switch input[type=checkbox].form-check-input").forEach((el: HTMLElement) => {
                            el.classList.add("dark");
                        });
                    }
                }
            },
            navigation: {
                itemsEnd: navLinks,
                // Add the branding icon & text
                onRendering: (props) => {
                    // Set the class names
                    props.className = "navbar-expand rounded-top";
                    props.type = Components.NavbarTypes.Primary

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
                items: [
                    {
                        text: "My Apps",
                        onRender: (el, item) => {
                            // Clear the existing button
                            el.innerHTML = "";
                            // Create a span to wrap the icon in
                            let span = document.createElement("span");
                            span.className = "bg-white d-inline-flex rounded";
                            el.appendChild(span);

                            // Render a tooltip
                            Components.Tooltip({
                                el: span,
                                content: "Show " + item.text,
                                btnProps: {
                                    // Render the icon button
                                    className: "p-1 pe-2",
                                    iconClassName: "me-1",
                                    iconType: personBoundingBox,
                                    iconSize: 24,
                                    isSmall: true,
                                    text: item.text,
                                    type: Components.ButtonTypes.OutlineSecondary,
                                    onClick: (item, ev) => {
                                        // See if we are setting the filter
                                        let elButton = ev.currentTarget as HTMLButtonElement;
                                        if (elButton.classList.contains("active")) {
                                            // Filter the data for apps belonging to the current user
                                            this._dashboard.filter(0, "");

                                            // Update the styling
                                            elButton.classList.remove("active");

                                            // Update the tooltip
                                            (elButton as any)._tippy.setContent("Show " + item.text);
                                        } else {
                                            // Filter the data for apps belonging to the current user
                                            this._dashboard.filter(0, "MyApp");

                                            // Update the styling
                                            elButton.classList.add("active");

                                            // Update the tooltip
                                            (elButton as any)._tippy.setContent("Hide " + item.text);
                                        }
                                    }
                                },
                            });
                            // Add the element
                            this._elFilterButtons.push(el.querySelector(".btn"));
                        }
                    }
                ],
                itemsEnd: [
                    {
                        text: "Add or Update an App",
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
                                    className: "p-1 pe-2",
                                    iconClassName: "me-1",
                                    iconType: fileEarmarkArrowUp,
                                    iconSize: 24,
                                    isSmall: true,
                                    text: "Upload App",
                                    type: Components.ButtonTypes.OutlineSecondary,
                                    onClick: () => {
                                        // Show the user agreement
                                        new UserAgreement(() => {
                                            // Upload the package file
                                            AppActions.upload(item => {
                                                // See if there is a flow
                                                if (AppConfig.Configuration.appFlows && AppConfig.Configuration.appFlows.newApp) {
                                                    // Run the flow for this app
                                                    AppActions.runFlow(AppConfig.Configuration.appFlows.newApp);
                                                }

                                                // Log
                                                DataSource.logItem({
                                                    LogUserId: ContextInfo.userId,
                                                    ParentId: item.AppProductID,
                                                    ParentListName: Strings.Lists.Apps,
                                                    Title: DataSource.AuditLogStates.AppAdded,
                                                    LogComment: `A new app ${item.Title} was added.`
                                                }, item);

                                                // View the app details
                                                this.viewAppDetails(item.Id, true);
                                            });
                                        });
                                    }
                                },
                            });
                        }
                    },
                    {
                        text: "Filters",
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
                                content: "Show " + item.text,
                                btnProps: {
                                    // Render the icon button
                                    className: "p-1 pe-2",
                                    iconClassName: "me-1",
                                    iconType: filterSquare,
                                    iconSize: 24,
                                    isSmall: true,
                                    text: item.text,
                                    type: Components.ButtonTypes.OutlineSecondary,
                                    onClick: () => {
                                        // Show the filter panel
                                        this._dashboard.showFilter();
                                    }
                                },
                            });
                        }
                    }
                ]
            },
            footer: {
                itemsEnd: [
                    {
                        className: "p-0 pe-none text-body",
                        text: "v" + Strings.Version,
                        onRender: (el) => {
                            // Hide version footer in a modern page
                            Strings.IsClassic ? null : el.classList.add("d-none");
                        }
                    }
                ]
            },
            table: {
                rows: DataSource.DocSetList.Items,
                dtProps: {
                    dom: 'rt<"row"<"col-sm-4"l><"col-sm-4"i><"col-sm-4"p>>',
                    pageLength: AppConfig.Configuration.paging,
                    columnDefs: [
                        {
                            targets: [0],
                            orderable: false
                        },
                        {
                            targets: [7],
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
                            // Set the filter value
                            el.setAttribute("data-filter", this.isMyApp(item) ? "MyApp" : "");

                            // See if the base64 value of the icon exists
                            if (item.AppThumbnailURLBase64) {
                                let elImg = document.createElement("img");
                                elImg.classList.add("icon");
                                ThemeManager.IsInverted ? elImg.classList.add("invert") : null;
                                elImg.src = item.AppThumbnailURLBase64;
                                el.appendChild(elImg);
                            }
                            // Ensure a url exists
                            else if (item.AppThumbnailURL && item.AppThumbnailURL.Url) {
                                // Render the link
                                el.innerHTML = `<img ${ThemeManager.IsInverted ? "class=\"invert\"" : ""} src="${item.AppThumbnailURL.Url}" height="32px" width="32px" title="${item.Title}">`;
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
                            // Set the filter
                            el.setAttribute("data-filter", item.AppStatus);
                            el.setAttribute("data-sort", item.AppStatus);

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
                        name: "",
                        title: "Deployed",
                        onRenderCell: (el, column, item: IAppItem) => {
                            // Set the status
                            let status = Common.appStatus(item);
                            status = status == item.AppStatus ? "Not Deployed" : status;
                            el.innerHTML = status;

                            // Set the filter
                            el.setAttribute("data-filter", status);
                            el.setAttribute("data-sort", status);
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
                        title: "App Developers",
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
                        title: "",
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
                                        // View the app details
                                        this.viewAppDetails(item.Id);
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

    // Method to view the app details
    private viewAppDetails(itemId: number, showEditForm: boolean = false) {
        // Show a loading dialog
        LoadingDialog.setHeader("Refreshing the Data");
        LoadingDialog.setBody("This will close after the data is loaded.");
        LoadingDialog.show();

        // Refresh the data
        this.refresh(itemId).then(() => {
            // Show a loading dialog
            LoadingDialog.setHeader("Loading Application Information");
            LoadingDialog.setBody("This will close after the data is loaded...");

            // Load the app dashboard
            DataSource.loadAppDashboard(itemId).then(() => {
                // Clear the details
                while (this._elAppDetails.firstChild) { this._elAppDetails.removeChild(this._elAppDetails.firstChild); }

                // Set the body
                new AppDashboard(this._elAppDetails, this._el, itemId);

                // Hide the apps
                this._el.classList.add("d-none");

                // Show the details
                this._elAppDetails.classList.remove("d-none");

                // Hide the loading dialog
                LoadingDialog.hide();

                // See if we are showing the edit form
                if (showEditForm) {
                    // Display the edit form
                    this._forms.edit(itemId, () => {
                        // Refresh the app details
                        this.viewAppDetails(itemId);
                    });
                }
            });
        });
    }
}