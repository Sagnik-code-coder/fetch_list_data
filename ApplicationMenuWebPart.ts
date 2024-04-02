import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPComponentLoader } from '@microsoft/sp-loader';
import * as strings from 'ApplicationMenuWebPartStrings';
import ApplicationMenu from './components/ApplicationMenu';
import { IApplicationMenuProps } from './components/IApplicationMenuProps';

export interface IApplicationMenuWebPartProps {
  description: string;
}
declare global {
    interface Window {
      loadDropdown: () => void;
    }
  }

export default class ApplicationMenuWebPart extends BaseClientSideWebPart<IApplicationMenuWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IApplicationMenuProps> = React.createElement(
      ApplicationMenu,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);

    this.domElement.innerHTML = 
    `
    <h3>Application:</h3>
    
    <div id="linksContainer">
    
    </div>
    `
  
  }

  protected onInit(): Promise<void> {
    //SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCISPOTest/SiteAssets/Branding/css/OneMazda_PubCollab_Home.css');
    //SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCISPOTest/SiteAssets/Branding/css/AppLinks.css');
    //SPComponentLoader.loadCss('https://mazdausa.sharepoint.com/sites/MCISPOTest/SiteAssets/Branding/css/styles/SuiteNav.css');
    SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/NameTag/js/jquery-1.10.2.min.js',{globalExportsName:'jquery'}).then(()=>{    
      SPComponentLoader.loadScript('https://y3mbk.sharepoint.com/sites/SharePointCRUD/SiteAssets/NameTag/js/ApplicationMenuClick.js');
      //SPComponentLoader.loadScript('https://mazdausa.sharepoint.com/sites/MCISPOTest/SiteAssets/Branding/js/ApplicationMenuClick.js');
    });
    
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
