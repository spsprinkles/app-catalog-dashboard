import { Helper, List, SPTypes } from "gd-sprest-bs";
import * as jQuery from "jquery";
import { Filter } from "./filter";
import { Navigation } from "./navigation";
import { Table } from "./table";
import { TermsOfUse } from "./tou";
import Strings from "../strings";
import "./styles.css";
declare var SP;

/**
 * Dashboard
 */
export class Dashboard {
    /**
     * Renders the project.
     * @param el - The element to render the dashboard to.
     */
    constructor(el: HTMLElement) {
        // Render the dashboard
        this.render(el);
    }

    /**
     * Main render method
     * @param el - The element to render the dashboard to.
     */
    private render(el: HTMLElement) {
        // Create the element
        let divMain = document.createElement("div");
        let divRow = document.createElement("div");
        let divCol = document.createElement("div");
        divRow.className = "row";
        divCol.className = "col";

        let divNavRow = divRow.cloneNode(false) as HTMLDivElement;
        let divNavCol = divCol.cloneNode(false) as HTMLDivElement;
        let divTouRow = divRow.cloneNode(false) as HTMLDivElement;
        let divTouCol = divCol.cloneNode(false) as HTMLDivElement;
        let divFilterRow = divRow.cloneNode(false) as HTMLDivElement;
        let divFilterCol = divCol.cloneNode(false) as HTMLDivElement;
        let divTableRow = divRow.cloneNode(false) as HTMLDivElement;
        let divTableCol = divCol.cloneNode(false) as HTMLDivElement;

        divFilterCol.classList.add("mb-4");
        divNavCol.id = "navigation";
        divTouRow.id = "touRow";
        divTouCol.id = "tou";
        divFilterCol.id = "filter";
        divTableCol.id = "table";

        divNavRow.appendChild(divNavCol);
        divTouRow.appendChild(divTouCol);
        divFilterRow.appendChild(divFilterCol);
        divTableRow.appendChild(divTableCol);
        divMain.appendChild(divNavRow);
        divMain.appendChild(divTouRow);
        divMain.appendChild(divFilterRow);
        divMain.appendChild(divTableRow);

        // Append it to the parent element
        el.appendChild(divMain);

        //Initially hide components before TOU check
        jQuery("#touRow").hide();
        jQuery("#filter").hide();
        jQuery("#table").hide();

        // Render the navigation
        new Navigation({
            el: el.querySelector("#navigation"),
            onSearch: value => {
                // Search the table
                table.search(value);
            }
        });

        new TermsOfUse({
            el: el.querySelector("#tou"),
            /*onSearch: value => {
                // Search the table
                table.search(value);
            }*/
        });

        SP.SOD.executeOrDelayUntilScriptLoaded(() => {
            //Need to disable caching
            List(Strings.Lists.Apps, {disableCache: true}).query({
                Select: ["EffectiveBasePermissions"]
            }).execute(obj => {
                if (Helper.hasPermissions(obj.EffectiveBasePermissions, SPTypes.BasePermissionTypes.EditListItems)) {
                    jQuery("#filter").slideDown();
                    jQuery("#table").slideDown();
                }
                else {
                    jQuery("#touRow").show();
                }
            });
        }, "sp.js");

        jQuery.ajax({
            type: "GET",
            url: Strings.TermsOfUseUrl,
            success: data => {
                document.getElementById("touInject").innerHTML = data;
            }
        });

        // Render the filter
        new Filter({
            el: el.querySelector("#filter"),
            onChange: value => {
                // Filter the table data
                table.applyFilter(value);
            },
            onRefresh: () => {
                table.refresh();
            }
        });

        // Render the table
        let table = new Table(el.querySelector("#table"));
    }
}