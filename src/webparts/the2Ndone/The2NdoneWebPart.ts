import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './The2NdoneWebPart.module.scss';
import * as strings from 'The2NdoneWebPartStrings';

export interface IThe2NdoneWebPartProps {
  description: string;
}

export default class The2NdoneWebPart extends BaseClientSideWebPart<IThe2NdoneWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.the2Ndone} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <!-- Premium Header -->
      <div class="${styles.premiumHeader}">
        <div class="${styles.premiumBadge}">⭐ PREMIUM</div>
        <h1 class="${styles.title}">🏆 Featured Projects</h1>
        <p class="${styles.subtitle}">Showcasing Excellence in Every Delivery</p>
        <p class="${styles.greeting}">Hello, <strong>${escape(this.context.pageContext.user.displayName)}</strong>!</p>
      </div>

      <!-- Featured Projects -->
      <div class="${styles.projectsSection}">
        <div class="${styles.projectGrid}">
          <div class="${styles.projectCard}">
            <div class="${styles.projectHeader}">
              <span class="${styles.projectIcon}">🏢</span>
              <span class="${styles.successBadge}">Success</span>
            </div>
            <h3>Enterprise Portal Redesign</h3>
            <p class="${styles.client}">Global Financial Corp</p>
            <p class="${styles.description}">Complete overhaul of internal portal serving 10,000+ employees with modern React architecture</p>
            <div class="${styles.metrics}">
              <span class="${styles.metric}"><strong>40%</strong> ↑ Performance</span>
              <span class="${styles.metric}"><strong>95%</strong> User Satisfaction</span>
            </div>
          </div>

          <div class="${styles.projectCard}">
            <div class="${styles.projectHeader}">
              <span class="${styles.projectIcon}">🚀</span>
              <span class="${styles.successBadge}">Success</span>
            </div>
            <h3>Mobile Banking App</h3>
            <p class="${styles.client}">National Bank</p>
            <p class="${styles.description}">Cross-platform mobile app with biometric authentication and real-time transactions</p>
            <div class="${styles.metrics}">
              <span class="${styles.metric}"><strong>1M+</strong> Downloads</span>
              <span class="${styles.metric}"><strong>4.8★</strong> Rating</span>
            </div>
          </div>

          <div class="${styles.projectCard}">
            <div class="${styles.projectHeader}">
              <span class="${styles.projectIcon}">☁️</span>
              <span class="${styles.successBadge}">Success</span>
            </div>
            <h3>Cloud Migration</h3>
            <p class="${styles.client}">Healthcare Provider</p>
            <p class="${styles.description}">Full Azure migration with zero downtime and enhanced security compliance</p>
            <div class="${styles.metrics}">
              <span class="${styles.metric}"><strong>60%</strong> ↓ Costs</span>
              <span class="${styles.metric}"><strong>99.9%</strong> Uptime</span>
            </div>
          </div>
        </div>
      </div>

      <!-- Client Testimonials -->
      <div class="${styles.testimonialsSection}">
        <h2 class="${styles.sectionTitle}">💬 What Our Clients Say</h2>
        <div class="${styles.testimonialGrid}">
          <div class="${styles.testimonialCard}">
            <div class="${styles.quote}">"</div>
            <p>Optimum Partners delivered beyond our expectations. Their technical expertise and professionalism were outstanding.</p>
            <div class="${styles.author}">
              <strong>Sarah Johnson</strong>
              <span>CTO, Global Financial Corp</span>
            </div>
          </div>
          <div class="${styles.testimonialCard}">
            <div class="${styles.quote}">"</div>
            <p>The team's ability to handle complex requirements and deliver on time was impressive. Highly recommended!</p>
            <div class="${styles.author}">
              <strong>Michael Chen</strong>
              <span>VP Technology, National Bank</span>
            </div>
          </div>
        </div>
      </div>

      <!-- Footer -->
      <div class="${styles.footer}">
        <p>${escape(this.properties.description)}</p>
        <p class="${styles.env}">${this._environmentMessage}</p>
      </div>
    </section>`;
  }

  protected onInit(): Promise<void> {
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
