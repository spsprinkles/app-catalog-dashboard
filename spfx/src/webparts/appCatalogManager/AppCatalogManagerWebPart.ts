import { Log, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import styles from './AppCatalogManagerWebPart.module.scss';
import * as strings from 'AppCatalogManagerWebPartStrings';

export interface IAppCatalogManagerWebPartProps {
  configuration: string;
}

// Reference the solution
import "../../../../dist/app-catalog-dashboard.js";
declare var AppDashboard;

export default class AppCatalogManagerWebPart extends BaseClientSideWebPart<IAppCatalogManagerWebPartProps> {

  public render(): void {
      // Render the application
      AppDashboard.render(this.domElement, this.context, Log, this.properties.configuration);
  }

  protected get dataVersion(): Version { return Version.parse('1.0'); }

  protected get disableReactivePropertyChanges(): boolean { return true; }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupFields: [
                PropertyPaneTextField('configuration', {
                  label: strings.ConfigLabel,
                  multiline: true,
                  rows: 40
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
