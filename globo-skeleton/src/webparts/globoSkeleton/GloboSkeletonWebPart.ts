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
      })
    })
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
