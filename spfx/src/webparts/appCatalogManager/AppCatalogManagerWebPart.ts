import { DisplayMode, Environment, Log, Version } from '@microsoft/sp-core-library';
import { IPropertyPaneConfiguration, PropertyPaneLabel, PropertyPaneTextField } from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as strings from 'AppCatalogManagerWebPartStrings';

export interface IAppCatalogManagerWebPartProps {
  configuration: string;
}

// Reference the solution
import "../../../../dist/app-catalog-dashboard.min.js";
declare const AppDashboard: {
  description: string;
  render: (props: {
    el: HTMLElement;
    appConfiguration?: string;
    context?: WebPartContext;
    displayMode?: DisplayMode;
    envType?: number;
    log?: Log,
    sourceUrl?: string;
  }) => void;
  updateTheme: (currentTheme: Partial<IReadonlyTheme>) => void;
  version: string;
};

export default class AppCatalogManagerWebPart extends BaseClientSideWebPart<IAppCatalogManagerWebPartProps> {
  public render(): void {
      // Render the application
      AppDashboard.render({
        el: this.domElement,
        appConfiguration: this.properties.configuration,
        context: this.context,
        displayMode: this.displayMode,
        envType: Environment.type,
        log: Log
      });
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    // Update the theme
    AppDashboard.updateTheme(currentTheme);
  }

  protected get dataVersion(): Version { return Version.parse(AppDashboard.version); }

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
                  rows: 30
                }),
                PropertyPaneLabel('version', {
                  text: "v" + AppDashboard.version
                })
              ]
            }
          ],
          header: {
            description: AppDashboard.description
          }
        }
      ]
    };
  }
}
