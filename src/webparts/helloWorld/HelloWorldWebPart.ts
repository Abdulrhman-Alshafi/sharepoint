import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloWorldWebPart.module.scss';
import * as strings from 'HelloWorldWebPartStrings';

export interface IHelloWorldWebPartProps {
  description: string;
}

export default class HelloWorldWebPart extends BaseClientSideWebPart<IHelloWorldWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.helloWorld} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <!-- Hero Section -->
      <div class="${styles.hero}">
        <div class="${styles.heroContent}">
          <h1 class="${styles.companyName}">✨ Optimum Partners</h1>
          <p class="${styles.tagline}">Transforming Ideas into Digital Excellence</p>
          <p class="${styles.greeting}">Welcome, <strong>${escape(this.context.pageContext.user.displayName)}</strong>!</p>
        </div>
      </div>

      <!-- Services Section -->
      <div class="${styles.services}">
        <h2 class="${styles.sectionTitle}">Our Expertise</h2>
        <div class="${styles.serviceGrid}">
          <div class="${styles.serviceCard}">
            <div class="${styles.icon}">💻</div>
            <h3>Web Development</h3>
            <p>Building responsive, scalable web applications using modern frameworks like React, Angular, and Vue</p>
          </div>
          <div class="${styles.serviceCard}">
            <div class="${styles.icon}">📱</div>
            <h3>Mobile Apps</h3>
            <p>Native and cross-platform mobile solutions for iOS and Android using React Native and Flutter</p>
          </div>
          <div class="${styles.serviceCard}">
            <div class="${styles.icon}">☁️</div>
            <h3>Cloud Solutions</h3>
            <p>Enterprise-grade cloud architecture and DevOps on Azure, AWS, and Microsoft 365</p>
          </div>
          <div class="${styles.serviceCard}">
            <div class="${styles.icon}">🤖</div>
            <h3>AI & Automation</h3>
            <p>Intelligent automation and AI-powered solutions to optimize your business processes</p>
          </div>
          <div class="${styles.serviceCard} ${styles.premium}">
            <div class="${styles.icon}">🔒</div>
            <h3>Cybersecurity</h3>
            <p>Advanced security solutions, penetration testing, and compliance management to protect your digital assets</p>
            <div class="${styles.premiumBadge}">Premium</div>
          </div>
          <div class="${styles.serviceCard} ${styles.premium}">
            <div class="${styles.icon}">📊</div>
            <h3>Data Analytics</h3>
            <p>Transform your data into actionable insights with advanced analytics, BI dashboards, and machine learning</p>
            <div class="${styles.premiumBadge}">Premium</div>
          </div>
        </div>
      </div>

      <!-- Stats Section -->
      <div class="${styles.stats}">
        <div class="${styles.statItem}">
          <div class="${styles.statNumber}">150+</div>
          <div class="${styles.statLabel}">Projects Delivered</div>
        </div>
        <div class="${styles.statItem}">
          <div class="${styles.statNumber}">50+</div>
          <div class="${styles.statLabel}">Happy Clients</div>
        </div>
        <div class="${styles.statItem}">
          <div class="${styles.statNumber}">10+</div>
          <div class="${styles.statLabel}">Years Experience</div>
        </div>
        <div class="${styles.statItem}">
          <div class="${styles.statNumber}">98%</div>
          <div class="${styles.statLabel}">Client Satisfaction</div>
        </div>
      </div>

      <!-- Technologies Section -->
      <div class="${styles.technologies}">
        <h2 class="${styles.sectionTitle}">Technologies We Master</h2>
        <div class="${styles.techBadges}">
          <span class="${styles.badge}">React</span>
          <span class="${styles.badge}">TypeScript</span>
          <span class="${styles.badge}">Node.js</span>
          <span class="${styles.badge}">Azure</span>
          <span class="${styles.badge}">SharePoint</span>
          <span class="${styles.badge}">Python</span>
          <span class="${styles.badge}">Docker</span>
          <span class="${styles.badge}">Kubernetes</span>
        </div>
      </div>

      <!-- Footer Info -->
      <div class="${styles.footer}">
        <p><strong>Description:</strong> ${escape(this.properties.description)}</p>
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
