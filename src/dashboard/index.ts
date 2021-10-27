import { ContextInfo, Helper, List, SPTypes } from "gd-sprest-bs";
import * as jQuery from "jquery";
import { Filter } from "./filter";
import { Navigation } from "./navigation";
import { Table } from "./table";
import { TermsOfUse } from "./tou";
import * as HTML from "./index.html";
import Strings from "../strings";
import "./styles.css";
declare var SP;

/**
 * Dashboard
 */
export class Dashboard {
    private _el: HTMLElement = null;

    /**
     * Renders the project.
     * @param el - The element to render the dashboard to.
     */
    constructor(elParent: HTMLElement) {
        // Create the element
        let el = document.createElement("div");
        el.innerHTML = HTML as any;
        this._el = el.firstChild as HTMLElement;

        // Append it to the parent element
        elParent.appendChild(this._el);

        // Render the dashboard
        this.render();
    }

    /**
     * Main render method
     * @param el - The element to render the dashboard to.
     */
    private render() {
        //Initially hide components before TOU check
        jQuery("#touRow").hide();
        jQuery("#filter").hide();
        jQuery("#table").hide();

        // Render the navigation
        new Navigation({
            el: this._el.querySelector("#navigation"),
            onSearch: value => {
                // Search the table
                table.search(value);
            }
        });

        new TermsOfUse({
            el: this._el.querySelector("#tou"),
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
            url: ContextInfo.webAbsoluteUrl + "/code/TermsOfUse.html",
            success: data => {
                document.getElementById("touInject").innerHTML = data;
            }
        });

        // Render the filter
        new Filter({
            el: this._el.querySelector("#filter"),
            onChange: value => {
                // Filter the table data
                table.applyFilter(value);
            },
            onRefresh: () => {
                table.refresh();
            }
        });

        // Render the table
        let table = new Table(this._el.querySelector("#table"));
    }
}