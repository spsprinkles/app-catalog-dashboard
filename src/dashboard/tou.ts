import { Components, Web } from "gd-sprest-bs";
import * as jQuery from "jquery";
import Strings from "../strings"
import * as HTML from "./tou.html";

// TermsOfUse Properties
export interface ITermsOfUseProps {
    el: HTMLElement;
}

/**
 * TermsOfUse
 */
export class TermsOfUse {
    private _props: ITermsOfUseProps = null;

    // Constructor
    constructor(props: ITermsOfUseProps) {
        // Save the parameters
        this._props = props;

        // Render
        this.render();
    }

    // Render
    private render() {
        var userLogin = "";
        Web().CurrentUser().query({
            Select: ["LoginName"]
        }).execute(obj => {
            userLogin = obj.LoginName;
            //userLogin = obj["d"].LoginName; //"i:0#.f|membership|john.doe@vasiloff.onmicrosoft.com"
        })

        // Render a card
        Components.Card({
            el: this._props.el,
            body: [
                {
                    onRender: (el) => {
                        el.innerHTML = HTML as any;
                        
                        var btn = Components.Button({
                            el: el,
                            text: " I Agree ",
                            type: Components.ButtonTypes.OutlineSecondary,
                            onClick: () => {
                                btn.disable();
                                Web().SiteGroups(Strings.Group).Users().add({
                                    LoginName: userLogin //Format: i:0#.f|membership|{email}
                                }).execute(obj => {
                                    jQuery("#touRow").slideUp();
                                    jQuery("#filter").slideDown();
                                    jQuery("#table").slideDown();
                                }, obj => {
                                    var response = JSON.parse(obj.response);
				                    alert(response.error.message.value);
                                });
                            }
                        });

                        //jQuery(el.querySelector(".col-12")).prepend(elDivNew);
                    }
                }
            ]
        });
    }
}