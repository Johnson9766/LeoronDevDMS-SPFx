import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';

// import styles from './WpHomePageWebPart.module.scss';
import * as strings from 'WpHomePageWebPartStrings';

export interface IWpHomePageWebPartProps {
  description: string;
}

export default class WpHomePageWebPart extends BaseClientSideWebPart<IWpHomePageWebPartProps> {



  public render(): void {
    this.domElement.innerHTML = `
    <div class="main-wrapper w-100 float-start">
    <div class="w-100 float-start leoron-banner-wrapper">
      <div class="w-100 float-start leoron-banner-swiper swiper">
        <div class="swiper-wrapper">
          <div class="leoron-banner-swiper-item swiper-slide">
            <img src="./resources/images/banner.png"/>
          </div>
          <div class="leoron-banner-swiper-item swiper-slide">
            <img src="./resources/images/banner.png"/>
          </div>
        </div>
         <div class="swiper-pagination"></div>
      </div>
      <div class="leoron-banner-title float-start w-100">
        <div class="container container-leoron mx-auto clearfix px-3 py-4 px-lg-4">
          <p class="text-white text-size-48 font-bold">Digital <br/>Archive</p>
        </div>
      </div>
    </div>
    <div class="w-100 float-start home-page">
      <div class="container-leoron mx-auto clearfix px-3 py-4 px-lg-4">
        <div class="w-100 float-start d-flex flex-column gap-4 home-page-tab-wrapper">
          <div class="home-page-tab-title-wrapper d-flex flex-column flex-sm-row w-100 float-start gap-3">
            <div data-tab-obj="tabGridData" class="home-page-tab-title flex-fill text-center">Quick Overview</div>
            <div data-tab-obj="tabGridData2" class="home-page-tab-title home-page-tab-title-active flex-fill text-center">Company Directory</div>
            <div data-tab-obj="tabGridData" class="home-page-tab-title flex-fill text-center">Centralized Functions</div>
          </div>
          <div class="w-100 float-start home-page-tab-content-wrapper">
           <div class="w-100 float-start home-page-tab-grid-view py-2 gap-3" id="tabGridRoot"></div>
          </div>
        </div>
      </div>
    </div>
  </div>`;
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      // this._environmentMessage = message;
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

    // this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

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
