import { ContextInfo } from "gd-sprest-bs";
import { IAppItem } from "./ds";
import Strings from "./strings";

// Determine if the user can edit the item
export function canEdit(item: IAppItem) {
    // Determine if the user is the author or owner
    return isOwner(item) || isSubmitter(item);
}

// Generates the document set dashboard url
export function generateDocSetUrl(item: IAppItem) {
    // Set the list url
    let listName = Strings.Lists.Apps.replace(/ /g, '');
    let listUrl = ContextInfo.webAbsoluteUrl + "/" + listName;
    let listUrlFolder = ContextInfo.webServerRelativeUrl + "/" + listName;

    // Return the url
    return listUrl + "/Forms/App/docsethomepage.aspx?ID=" +
        item.Id + "&FolderCTID=" + item.ContentTypeId + "&RootFolder=" + listUrlFolder + "/" + item.Id;
}

// Determine if the user is an owner
export function isOwner(item: IAppItem) {
    // Determine if the user is an owner
    return item.OwnersId ? item.OwnersId.results.indexOf(ContextInfo.userId) != -1 : false;
}

// Determine if the user is a submitter
export function isSubmitter(item: IAppItem) {
    // Determine if the user submitted the app
    return item.AuthorId == ContextInfo.userId;
}