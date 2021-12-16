import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './AppCatalogManagerWebPart.module.scss';
import * as strings from 'AppCatalogManagerWebPartStrings';

export interface IAppCatalogManagerWebPartProps {
  description: string;
}

// Reference the solution
import "../../../../dist/app-catalog-dashboard.js";
declare var AppDashboard;

export default class AppCatalogManagerWebPart extends BaseClientSideWebPart<IAppCatalogManagerWebPartProps> {

  public render(): void {
    // Render the application
    AppDashboard.render(this.domElement, this.context, "/sites/demo/appcatalog");
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
