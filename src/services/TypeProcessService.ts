import { Text } from "@microsoft/sp-core-library";
import { ITypeProcessService } from "./ITypeProcessService";
import {
  ISPHttpClientOptions,
  SPHttpClient,
  SPHttpClientResponse,
} from "@microsoft/sp-http";
import * as strings from "servicesStrings";

export class TypeProcessService implements ITypeProcessService {
  private spHttpClient: SPHttpClient;

  constructor(spHttpClient: SPHttpClient) {
    this.spHttpClient = spHttpClient;
  }

  /**
   * Adds a new SharePoint list item to a list using the given data.
   *
   * @param webUrl The absolute Url to the SharePoint site.
   * @param listUrl The server-relative Url to the SharePoint list.
   * @param fieldsSchema The array of field schema for all relevant fields of this list.
   * @param data An object containing all the field values to set on creating item.
   * @returns Promise object represents the updated or erroneous form field values.
   */

  public createTypeProcess(
    webUrl: string,
    listUrl: any,
    data: any
  ): Promise<any> {
    return new Promise<any>((resolve, reject) => {
      const httpClientOptions: ISPHttpClientOptions = {
        headers: {
          Accept: "application/json;odata=verbose",
          "Content-type": "application/json;odata=verbose",
          "X-SP-REQUESTRESOURCES": "listUrl=" + encodeURIComponent(listUrl),
          "odata-version": "",
        },
        body: JSON.stringify({
          listItemCreateInfo: {
            __metadata: { type: "SP.ListItemCreationInformationUsingPath" },
            FolderPath: {
              __metadata: { type: "SP.ResourcePath" },
              DecodedUrl: listUrl,
            },
          },
          formValue: data.formValues,
          bNewDocumentUpdate: false,
          checkInComment: null,
        }),
      };

      const endpoint =
        `${webUrl}/_api/web/GetList(@listUrl)/AddValidateUpdateItemUsingPath` +
        `?@listUrl=${encodeURIComponent("'" + listUrl + "'")}`;

      this.spHttpClient
        .post(endpoint, SPHttpClient.configurations.v1, httpClientOptions)
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            return response.json();
          } else {
            reject(this.getErrorMessage(webUrl, response));
          }
        })
        .then((respData) => {
          resolve(respData.d.AddValidateUpdateItemUsingPath.results);
        })
        .catch((error) => {
          reject(this.getErrorMessage(webUrl, error));
        });
    });
  }

  /**
   * Returns an error message based on the specified error object
   * @param error : An error string/object
   */
  private getErrorMessage(webUrl: string, error: any): string {
    let errorMessage: string = error.statusText
      ? error.statusText
      : error.statusMessage
      ? error.statusMessage
      : error;
    const serverUrl = `{window.location.protocol}//{window.location.hostname}`;
    const webServerRelativeUrl = webUrl.replace(serverUrl, "");

    if (error.status === 403) {
      errorMessage = Text.format(
        strings.ErrorWebAccessDenied,
        webServerRelativeUrl
      );
    } else if (error.status === 404) {
      errorMessage = Text.format(
        strings.ErrorWebNotFound,
        webServerRelativeUrl
      );
    }
    return errorMessage;
  }
}
