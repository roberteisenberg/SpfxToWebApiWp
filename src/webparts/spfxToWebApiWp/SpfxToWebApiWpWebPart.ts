import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { Version } from '@microsoft/sp-core-library';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './SpfxToWebApiWpWebPart.module.scss';
import * as strings from 'SpfxToWebApiWpWebPartStrings';

export interface ISpfxToWebApiWpWebPartProps {
  description: string;
}

export default class SpfxToWebApiWpWebPart extends BaseClientSideWebPart<ISpfxToWebApiWpWebPartProps> {

  private ordersClient: AadHttpClient;
  // private azureApiUrl = "https://apiusedbydegfrontend.azurewebsites.net/";
  private azureApiUrl = "https://testaadwithvision.azurewebsites.net/";
  private appId = "12fe8014-e4f8-4484-9fdf-bf631efbcde6";
  protected onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this.context.aadHttpClientFactory
        .getClient(this.appId)
        .then(client => { this.ordersClient = client; resolve(); });
    });
  }

  public render(): void {
    this.context.statusRenderer.displayLoadingIndicator(this.domElement, 'orders');

    this.ordersClient
      .get(this.azureApiUrl + 'test', AadHttpClient.configurations.v1)
      .then((res: HttpClientResponse): Promise<any> => (res.json()))
      .then((orders: any): void => {
        console.debug("Return from API:", orders);
        this.context.statusRenderer.clearLoadingIndicator(this.domElement);
        this.domElement.innerHTML = `
        <div class="${ styles.spfxToWebApiWp}">
          <div class="${ styles.container}">
            <div class="${ styles.row}">
              <div class="${ styles.column}">
                <p class="${ styles.description}">
                    the results: ${orders}
                </p>
                <button id="myRefreshButton">Refresh</button>
              </div>
            </div>
          </div>
        </div>`;

        const button = document.getElementById("myRefreshButton");
        button.addEventListener("click", (e: Event) => this.render());
      }, (err: any): void => {
        this.context.statusRenderer.renderError(this.domElement, err);
      });
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
