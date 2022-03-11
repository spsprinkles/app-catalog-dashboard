import { ContextInfo } from "gd-sprest-bs";
import { IAppItem } from "./ds";

// Determines the current application status
export function appStatus(item: IAppItem) {
    // See if item is rejected
    return item.AppIsRejected ? "Rejected" : (item.AppIsTenantDeployed ? "Deployed" : item.AppStatus);
}

// Determines if the user can edit the item
export function canEdit(item: IAppItem) {
    // The user is the author or owner
    return isOwner(item) || isSubmitter(item);
}

// Generates an embedded SVG image to embed in a style tag
export function generateEmbeddedSVG(svg: SVGElement) {
    return "url(\"data:image/svg+xml," + (svg.outerHTML as any).replaceAll("\"", "'").replaceAll("<", "%3C").replaceAll(">", "%3E").replaceAll("\n", "").replaceAll("  ", " ") + "\");";
}

// Returns the AppIcon file as an SVG element
export function getAppIcon(height?, width?, className?) {
    if (height === void 0) { height = 32; }
    if (width === void 0) { width = 32; }
    // Get the icon element
    var elDiv = document.createElement("div");
    elDiv.innerHTML = "<svg xmlns='http://www.w3.org/2000/svg' width='16' height='16' fill='currentColor' class='dr-appicon' viewBox='0 0 1000 1000'><path d='M500,10C229.4,10,10,229.4,10,500c0,270.6,219.4,490,490,490c270.6,0,490-219.4,490-490C990,229.4,770.6,10,500,10z M456.9,731.9c0,10.2-3.6,18.9-10.9,26c-7.3,7.1-16.1,10.6-26.3,10.6H271.9c-10.2,0-18.9-3.5-26-10.6c-7.1-7.1-10.6-15.8-10.6-26V583.7c0-10.2,3.5-18.9,10.6-26c7.1-7.1,15.8-10.6,26-10.6h147.7c10.2,0,19,3.5,26.3,10.6c7.3,7.1,10.9,15.8,10.9,26L456.9,731.9L456.9,731.9z M456.9,436c0,10.2-3.6,19-10.9,26.3c-7.3,7.3-16.1,10.9-26.3,10.9H271.9c-10.2,0-18.9-3.6-26-10.9c-7.1-7.3-10.6-16.1-10.6-26.3V288.3c0-10.2,3.5-18.9,10.6-26c7.1-7.1,15.8-10.6,26-10.6h147.7c10.2,0,19,3.5,26.3,10.6c7.3,7.1,10.9,15.8,10.9,26L456.9,436L456.9,436z M752.3,731.9c0,10.2-3.5,18.9-10.6,26c-7.1,7.1-15.8,10.6-26,10.6H567.9c-10.2,0-19-3.5-26.3-10.6c-7.3-7.1-10.9-15.8-10.9-26V583.7c0-10.2,3.6-18.9,10.9-26c7.3-7.1,16.1-10.6,26.3-10.6h147.7c10.2,0,18.9,3.5,26,10.6c7.1,7.1,10.6,15.8,10.6,26V731.9z M797.8,371.6L687.9,481.4c-7.9,7.9-17,11.8-27.5,11.8c-10.4,0-19.6-3.9-27.5-11.8L523,371.6c-7.5-7.5-11.2-16.5-11.2-27.2c0-10.6,3.7-19.9,11.2-27.8l109.9-109.9c7.9-7.5,17-11.2,27.5-11.2c10.4,0,19.6,3.7,27.5,11.2l109.9,109.9c7.9,7.9,11.8,17.1,11.8,27.8C809.6,355,805.6,364.1,797.8,371.6z'/></svg>";
    var icon = elDiv.firstChild as SVGImageElement;
    if (icon) {
        // See if a class name exists
        if (className) {
            // Parse the class names
            var classNames = className.split(' ');
            for (var i = 0; i < classNames.length; i++) {
                // Add the class name
                icon.classList.add(classNames[i]);
            }
        }
        // Set the height/width
        icon.setAttribute("height", (height ? height : 32).toString());
        icon.setAttribute("width", (width ? width : 32).toString());
        // Update the styling
        icon.style.pointerEvents = "none";
        // Support for IE
        icon.setAttribute("focusable", "false");
    }
    // Return the icon
    return icon;
}

// Determines if the user is an owner
export function isOwner(item: IAppItem) {
    // Determine if the user is an owner
    return item.AppDevelopersId ? item.AppDevelopersId.results.indexOf(ContextInfo.userId) != -1 : false;
}

// Determines if the user is a submitter
export function isSubmitter(item: IAppItem) {
    // Determine if the user submitted the app
    return item.AuthorId == ContextInfo.userId;
}

// Updates the url
export function updateUrl(url: string) {
    return url.replace("~site/", ContextInfo.webServerRelativeUrl + "/")
        .replace("~sitecollection/", ContextInfo.siteServerRelativeUrl + "/");
}