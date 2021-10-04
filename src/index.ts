import { Configuration } from "./cfg";
import { Dashboard } from "./dashboard";
import Strings from "./strings";

// Ensure the data tables plugin is configured
import * as DataTable from "datatables.net";
import { jQuery } from "gd-sprest-bs";
//import "datatables.net-dt/css/jquery.dataTables.min.css";
import "datatables.net-bs4/js/dataTables.bootstrap4.min.js";
import "datatables.net-colreorder/js/dataTables.colReorder.min.js";

// See if jQuery is defined in the DataTable lib
if (DataTable.prototype.constructor.$ == undefined) {
    // Set the reference
    DataTable.prototype.constructor.$ = jQuery;
} else {
    // Update this jQuery reference for this library
    window["$REST"].jQuery = DataTable.prototype.constructor.$;
}

// Create the global variable for this solution
window[Strings.GlobalVariable] = {
    Configuration
}

// Get the main element to render the solution to
let el = document.getElementById(Strings.AppElementId);
if (el) {
    // Initialize the solution
    new Dashboard(el);
} else {
    // Log
    console.log("[" + Strings.ProjectName + "] Error finding the element with id '" + Strings.AppElementId + "'");
}

// Remove the extra border spacing on the webpart
jQuery('table.ms-core-tableNoSpace', '#contentBox').removeClass('ms-webpartPage-root');