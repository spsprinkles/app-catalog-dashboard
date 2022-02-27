import { ContextInfo } from "gd-sprest-bs";
import { IAppItem } from "./ds";

// Determines the current application status
export function appStatus(item: IAppItem) {
    // See if item is rejected
    return item.AppIsRejected ? "Rejected" : item.AppStatus;
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