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

export interface IGloboSkeletonWebPartProps {
  description: string;
  showStaffNumber: Boolean;
}

export default class GloboSkeletonWebPart extends BaseClientSideWebPart<IGloboSkeletonWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
      <div class="${styles.globoSkeleton}">
        <div class="${styles.container}">
          <div class="${styles.row}">
            <div class="${styles.column}">
              <span class="${styles.title}">Welcome to SharePoint!</span>
              <p class="${styles.subTitle}">Customize SharePoint experiences using Web Parts.</p>
              <p class="${styles.description}">${escape(this.properties.description)}</p>
              <p class="${styles.description}">Show Staff Number: ${this.properties.showStaffNumber}</p>
              <p class="${styles.description}">Name: ${escape(this.context.pageContext.user.displayName)}</p>
              <a href="https://aka.ms/spfx" class="${styles.button}">
                <span class="${styles.label}">Learn more</span>
              </a>
            </div>
          </div>
        </div>
      </div>`;
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
