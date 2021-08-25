import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './GloboSkeletonWebPart.module.scss';
import * as strings from 'GloboSkeletonWebPartStrings';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types'

export interface IGloboSkeletonWebPartProps {
  description: string;
  showStaffNumber: Boolean;
}

export default class GloboSkeletonWebPart extends BaseClientSideWebPart<IGloboSkeletonWebPartProps> {

  public render(): void {
    this.context.msGraphClientFactory
    .getClient()
    .then((client: MSGraphClient): void => {
      // get information about the current user from the Microsoft Graph
      client
      .api('/me')
      .get((error, userProfile: any, rawResponse?: any) => {
        this.domElement.innerHTML = `
          <div class="${styles.globoSkeleton}">
            <div class="${styles.container}">
              <div class="${styles.row}">
                <span class="${styles.title}">Welcome ${escape(this.context.pageContext.user.displayName)}!</span>
                <div class="${styles.subTitle}" id="spUserContainer"></div>
                <div class="${styles.rowTable}">
                  <div class="${styles.columnTable3}">
                    <h2>Manager</h2>
                    <div id="spManager"></div>
                  </div>
                  <div class="${styles.columnTable3}">
                    <h2>Colleagues</h2>
                    <div id="spColleagues"></div>
                  </div>
                  <div class="${styles.columnTable3}">
                    <h2>Direct Reports</h2>
                    <div id="spReports"></div>
                  </div>
                </div>
              </div>
            </div>
          </div>
          `;          

          this._renderJobTitle(userProfile);
          if(this.properties.showStaffNumber) {
            this._renderEmployeeId(client);
          }
          this._renderDirectReports(client);
          this._renderManagerAndColleagues(client, userProfile);
      })
    })
  }

  private _renderCollegues(client: MSGraphClient, userProfile: MicrosoftGraph.User, managerId: string): void {
    client
    .api(`/users/${managerId}/directReports`)
    .get((error, directReports: any, rawResponse?: any) => {
      const spColleagueContainer: Element = this.domElement.querySelector('#spColleagues');
      let html: string = spColleagueContainer.innerHTML;
      directReports.value.forEach((directReport: MicrosoftGraph.User) => {
        if (directReport.id != userProfile.id) {
          html += `<p class="${styles.description}">${escape(directReport.displayName)}</p>`;
        }
      });
      spColleagueContainer.innerHTML = html;
    })
  }

  private _renderManagerAndColleagues(client: MSGraphClient, userProfile: MicrosoftGraph.User): void {
    client
    .api('/me/manager')
    .get((error, manager: MicrosoftGraph.User, rawResponse?:any) => {
      const spUserContainer: Element = this.domElement.querySelector('#spManager');
      let html: string = spUserContainer.innerHTML;
      if (manager != null) {
        html += `<p class="${styles.description}">${escape(manager.displayName)}</p>`;
        spUserContainer.innerHTML = html;
      }
      this._renderCollegues(client, userProfile, manager.id);
    })
  }

  private _renderDirectReports(client: MSGraphClient): void {
    client
    .api('/me/directReports')
    .get((error, directReports: any, rawResponse?: any) => {
      const spUserContainer: Element = this.domElement.querySelector('#spReports');
      let html: string = spUserContainer.innerHTML;
      directReports.value.forEach((directReport: MicrosoftGraph.User) => {
        html += `<p class="${styles.description}">${escape(directReport.displayName)}</p>`;
      });
      spUserContainer.innerHTML = html;
    })
  }

  private _renderEmployeeId(client: MSGraphClient): void {
    client.api('/me/id/$value')
    .responseType('TEXT')
    .get((error, employeeId: any, rawResponse?:any) => {
      const spUserContainer: Element = this.domElement.querySelector('#spUserContainer');
      spUserContainer.innerHTML += `<p>${escape(employeeId)}</p>`
    })
  }

  private _renderJobTitle(userProfile: MicrosoftGraph.User): void {
    const spUserContainer: Element = this.domElement.querySelector('#spUserContainer');
    let html: string = spUserContainer.innerHTML;
    html += `<p>${escape(userProfile.jobTitle)}</p>`;
    spUserContainer.innerHTML = html;
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
                }),
                PropertyPaneToggle('showStaffNumber', {
                  label: "Show Staff Number",
                  checked: true
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
