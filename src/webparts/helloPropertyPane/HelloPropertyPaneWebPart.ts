import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './HelloPropertyPaneWebPart.module.scss';
import * as strings from 'HelloPropertyPaneWebPartStrings';

import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export interface IHelloPropertyPaneWebPartProps {
  description: string;
  myContinent: string;
  numContinentsVisited: number;
}

export default class HelloPropertyPaneWebPart extends BaseClientSideWebPart<IHelloPropertyPaneWebPartProps> {

  public render(): void {

    this.context.msGraphClientFactory.getClient().then((client: MSGraphClient ) : void => 
    {

      client.api('/me/messages').top(10).get((error, messages: any, rawResponse?: any) => {

        this.domElement.innerHTML = `
        <div class="${ styles.helloPropertyPane}">
        <div class="${ styles.container}">
          <div class="${ styles.row}">
            <div class="${ styles.column}">
              <span class="${ styles.title}">Welcome to SharePoint!</span>
              <p class="${ styles.subTitle}">Use Microsoft Graph in SharePoint Framework.</p>
              <div id="spListContainer" />
            </div>
          </div>
        </div>
        </div>`;
          // List the latest emails based on what we got from the Graph
        this._renderEmailList(messages.value);  
      });
    });
   
  }
  private _renderEmailList(messages: MicrosoftGraph.Message[]): void {
    let html: string = '';
    for (let index = 0; index < messages.length; index++) {
      html += `<p class="${styles.description}">Email ${index + 1} - ${escape(messages[index].subject)}</p>`;
    }
  
    // Add the emails to the placeholder
    const listContainer: Element = this.domElement.querySelector('#spListContainer');
    listContainer.innerHTML = html;
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
                PropertyPaneTextField('myContinent', {
                   label: 'Continent where I currently reside' 
                }),
                PropertyPaneSlider('numContinentsVisited', { 
                  label: 'Number of continents I\'ve visited',
                  min: 1,
                  max: 7,
                  showValue: true, 
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
