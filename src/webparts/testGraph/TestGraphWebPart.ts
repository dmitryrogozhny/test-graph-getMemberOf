import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TestGraphWebPart.module.scss';
import * as strings from 'TestGraphWebPartStrings';

import { graph } from "@pnp/graph/presets/all";

export interface ITestGraphWebPartProps {
  description: string;
}

export default class TestGraphWebPart extends BaseClientSideWebPart<ITestGraphWebPartProps> {

  protected onInit(): Promise<void> {
    return super.onInit().then(_ => {
      graph.setup({ spfxContext: this.context });
    });
  }

  public render(): void {
    graph.me.getMemberGroups().then((groups) => {
      const expectedGroups: string[] = groups.value;
      const actualGroups: string[] = <string[]><unknown>groups;

      console.log('Raw: ', groups);
      console.log('Expected: ', expectedGroups);
      console.log('Actual:', actualGroups);
    });

    this.domElement.innerHTML = `<div class="${styles.testGraph}">Test Graph API call</div>`;
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
