import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './The3OneWebPart.module.scss';
import * as strings from 'The3OneWebPartStrings';

export interface IThe3OneWebPartProps {
  description: string;
}

export default class The3OneWebPart extends BaseClientSideWebPart<IThe3OneWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    this.domElement.innerHTML = `
    <section class="${styles.the3One} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <!-- Premium Header -->
      <div class="${styles.premiumHeader}">
        <div class="${styles.premiumBadge}">⭐ PREMIUM</div>
        <h1 class="${styles.title}">👥 Meet Our Team</h1>
        <p class="${styles.subtitle}">World-Class Developers & Innovators</p>
        <p class="${styles.greeting}">Welcome, <strong>${escape(this.context.pageContext.user.displayName)}</strong>!</p>
      </div>

      <!-- Team Members -->
      <div class="${styles.teamSection}">
        <div class="${styles.teamGrid}">
          <!-- Team Member 1 -->
          <div class="${styles.teamCard}">
            <div class="${styles.avatar}">👨‍💻</div>
            <h3>Ahmad Al-Shareef</h3>
            <p class="${styles.role}">Chief Technology Officer</p>
            <p class="${styles.bio}">15+ years leading enterprise architecture and cloud solutions. Azure MVP.</p>
            <div class="${styles.skills}">
              <span class="${styles.skill}">Azure</span>
              <span class="${styles.skill}">Architecture</span>
              <span class="${styles.skill}">DevOps</span>
            </div>
            <div class="${styles.contact}">
              <a href="mailto:ahmad@optimumpartners.jo">✉️ Email</a>
            </div>
          </div>

          <!-- Team Member 2 -->
          <div class="${styles.teamCard}">
            <div class="${styles.avatar}">👩‍💻</div>
            <h3>Layla Hassan</h3>
            <p class="${styles.role}">Lead Full-Stack Developer</p>
            <p class="${styles.bio}">Expert in React, Node.js, and microservices. Building scalable solutions.</p>
            <div class="${styles.skills}">
              <span class="${styles.skill}">React</span>
              <span class="${styles.skill}">Node.js</span>
              <span class="${styles.skill}">TypeScript</span>
            </div>
            <div class="${styles.contact}">
              <a href="mailto:layla@optimumpartners.jo">✉️ Email</a>
            </div>
          </div>

          <!-- Team Member 3 -->
          <div class="${styles.teamCard}">
            <div class="${styles.avatar}">👨‍💻</div>
            <h3>Omar Khalil</h3>
            <p class="${styles.role}">Mobile Development Lead</p>
            <p class="${styles.bio}">Creating stunning mobile experiences with React Native and Flutter.</p>
            <div class="${styles.skills}">
              <span class="${styles.skill}">React Native</span>
              <span class="${styles.skill}">Flutter</span>
              <span class="${styles.skill}">iOS/Android</span>
            </div>
            <div class="${styles.contact}">
              <a href="mailto:omar@optimumpartners.jo">✉️ Email</a>
            </div>
          </div>

          <!-- Team Member 4 -->
          <div class="${styles.teamCard}">
            <div class="${styles.avatar}">👩‍💻</div>
            <h3>Sara Nasser</h3>
            <p class="${styles.role}">AI & Data Science Lead</p>
            <p class="${styles.bio}">Pioneering AI solutions and machine learning models for business intelligence.</p>
            <div class="${styles.skills}">
              <span class="${styles.skill}">Python</span>
              <span class="${styles.skill}">ML/AI</span>
              <span class="${styles.skill}">TensorFlow</span>
            </div>
            <div class="${styles.contact}">
              <a href="mailto:sara@optimumpartners.jo">✉️ Email</a>
            </div>
          </div>
        </div>
      </div>

      <!-- Company Values -->
      <div class="${styles.valuesSection}">
        <h2 class="${styles.sectionTitle}">🌟 Our Core Values</h2>
        <div class="${styles.valuesGrid}">
          <div class="${styles.valueCard}">
            <div class="${styles.valueIcon}">🎯</div>
            <h4>Excellence</h4>
            <p>We deliver nothing but the best quality in everything we do</p>
          </div>
          <div class="${styles.valueCard}">
            <div class="${styles.valueIcon}">🚀</div>
            <h4>Innovation</h4>
            <p>Always pushing boundaries with cutting-edge technology</p>
          </div>
          <div class="${styles.valueCard}">
            <div class="${styles.valueIcon}">🤝</div>
            <h4>Partnership</h4>
            <p>Your success is our success - we're in this together</p>
          </div>
          <div class="${styles.valueCard}">
            <div class="${styles.valueIcon}">⚡</div>
            <h4>Agility</h4>
            <p>Fast, flexible, and responsive to your evolving needs</p>
          </div>
        </div>
      </div>

      <!-- Contact CTA -->
      <div class="${styles.ctaSection}">
        <h2>📧 Let's Work Together</h2>
        <p>Ready to transform your digital presence? Get in touch with our team today!</p>
        <div class="${styles.ctaButtons}">
          <a href="mailto:info@optimumpartners.jo" class="${styles.ctaButton}">Contact Us</a>
          <a href="tel:+962123456789" class="${styles.ctaButtonSecondary}">Call Us</a>
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
