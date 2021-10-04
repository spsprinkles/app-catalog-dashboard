import { Components, Helper, Icons, IconTypes, List, SPTypes, jQuery } from "gd-sprest-bs";
import Strings from "../strings";

// Navigation Props
export class INavigationProps {
    el: HTMLElement;
    onSearch: (value: string) => void;
}

/**
 * Navigation
 */
export class Navigation {
    private _props: INavigationProps = null;
    private _formUrl: string = null;

    // Constructor
    constructor(props: INavigationProps) {
        // Save the properties
        this._props = props;

        // Load the form url
        this.load().then(() => {
            // Render the navigation
            this.render();
        });
    }

    // Load the form url
    private load() {
        // Return a promise
        return new Promise((resolve, reject) => {
            // Load the form for the list
            List(Strings.Lists.Apps).Forms().query({
                Filter: "FormType eq " + SPTypes.PageType.NewForm
            }).execute(forms => {
                // Save the form url
                this._formUrl = forms.results[0].ServerRelativeUrl;

                // Resolve the promise
                resolve(null);
            }, reject);
        });
    }

    // Renders the navigation
    private render() {
        // Render the navigation
        let navBar = Components.Navbar({
            el: this._props.el,
            brand: Strings.ProjectName,
            className: "bg-sharepoint rounded-top",
            type: Components.NavbarTypes.Primary,
            searchBox: {
                hideButton: true,
                onChange: value => {
                    // Execute the search event
                    this._props.onSearch(value);
                },
                onSearch: value => {
                    // Execute the search event
                    this._props.onSearch(value);
                }
            }/*,
            items: [
                {
                    text: "New Item",
                    onClick: () => {
                        // Display a new item form
                        Helper.SP.ModalDialog.showModalDialog({
                            title: "New Item",
                            url: this._formUrl,
                            dialogReturnValueCallback: result => {
                                // See if an item was created
                                if (result == SPTypes.ModalDialogResult.OK) {
                                    // Refresh the page
                                    document.location.reload();
                                }
                            }
                        });
                    }
                },
                {
                    text: "Reports",
                    onClick: () => { }
                },
                {
                    text: "Administration",
                    onClick: () => { }
                },
                {
                    text: "Help",
                    items: [
                        {
                            text: "Common Questions",
                            href: "#"
                        },
                        {
                            text: "How To",
                            href: "#"
                        },
                        {
                            text: "Contact",
                            href: "#"
                        }
                    ],
                    onClick: (item, ev) => {
                        // Prevent postback
                        ev.preventDefault();

                        // Display the page in a modal
                        Helper.SP.ModalDialog.showModalDialog({
                            showMaximized: true,
                            title: item.text,
                            url: item.href
                        });
                    }
                }
            ]*/
        });

        // Create a Shield Icon
        let shieldIcon = Icons(IconTypes.ColumnsGap, 32, 32); //AppIndicator Bag BoundingBox Box ColumnsGap AspectRatio Award
        // Add margin spacing to the shield icon
        shieldIcon.classList.add("ml-n1");
        shieldIcon.classList.add("mr-2");
        // Color it white
        shieldIcon.classList.add("text-white");
        // Insert it to the beginning of the NavBar
        jQuery(navBar.el).prepend(shieldIcon);
        // Create a Search Icon
        let searchIcon = Icons(IconTypes.Search, 24, 24);
        // Add a right margin to the search icon
        searchIcon.classList.add("mr-2");
        // Color it white
        searchIcon.classList.add("text-white");
        // Insert it to the beginning of the search form
        jQuery(navBar.el.querySelector(".form-inline")).prepend(searchIcon);
        // Remove the right margin on the search box
        navBar.el.querySelector(".form-control").classList.remove("mr-sm-2");
        // Update the placeholder attribute
        navBar.el.querySelector(".form-control").setAttribute("placeholder", "Search app list")
        // Set the title attribute
        navBar.el.querySelector(".form-control").setAttribute("title", "Search app list")
    }
}