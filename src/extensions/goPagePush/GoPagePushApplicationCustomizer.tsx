// this extension has been deployed to https://umaknowdev.sharepoint.com/sites/e2e 
import * as React from "react";
import { ReactElement } from "react";
import * as ReactDOM from "react-dom";
import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import { BaseApplicationCustomizer, PlaceholderContent, PlaceholderName } from "@microsoft/sp-application-base"; // used for the placeholder 
import { Dialog } from "@microsoft/sp-dialog";
import { Button } from "office-ui-fabric-react/lib/Button";

import * as strings from "GoPagePushApplicationCustomizerStrings";
import styles from "./AppCustomizer.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";

import { AadHttpClient, HttpClientResponse, IHttpClientOptions, HttpClient, AadTokenProvider} from "@microsoft/sp-http"; // used for the connection to API
const LOG_SOURCE: string = "GoPagePushApplicationCustomizer";

export interface IGoPagePushApplicationCustomizerProperties {
  Bottom: string;
}

export default class GoPagePushApplicationCustomizer extends BaseApplicationCustomizer<
  IGoPagePushApplicationCustomizerProperties
> {
  private _bottomPlaceholder: PlaceholderContent | undefined;
  private _token; // stores the bearer token

  @override
  public onInit(): Promise<void> {
    var clientID = '44e56dc9-0513-4445-9895-52ca527f85a9' // hard-coded for now
    // get the token and pass in the client id 
    this.context.aadTokenProviderFactory.getTokenProvider().then((value: AadTokenProvider) => {
      value.getToken(clientID).then(
        token => { this._token = token }
      ).catch(err => {
        console.log("printing the error: ", err); 
      })
    });

    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    // wait for the placeholders to be created (or handle them being changed) and then render
    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderPlaceHolders
    );
    // add the "connect to api" button to the buttom placeholder
    this._addButton();
    return Promise.resolve();
  }

  private _renderReactElement(component: ReactElement<any>, node: Element): void {
    ReactDOM.unmountComponentAtNode(node);
    ReactDOM.render(component, node);
  }

  // for the "connect to api" button
  private _addButton() {
    const button: React.ReactElement<IGoPagePushApplicationCustomizerProperties> = (
      <div className="button">
      <React.Fragment>
        <Button onClick={this._connect.bind(this)}> CLICK HERE TO CONNECT TO THE API</Button>
      </React.Fragment>
      </div>
    );
    this._renderReactElement(button, this._bottomPlaceholder.domElement);
  }

  // placeholder
  private _renderPlaceHolders(): void {
    // handling the bottom placeholder
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom
      );

      // the extension should not assume that the expected placeholder is available.
      if (!this._bottomPlaceholder) {
        console.error("The expected placeholder (Bottom) was not found.");
        return;
      }

      if (this.properties) {
        let bottomString: string = this.properties.Bottom;
        if (!bottomString) {
          bottomString = "(Bottom property was not defined.)";
        }

        if (this._bottomPlaceholder.domElement) {
          this._bottomPlaceholder.domElement.innerHTML = `
          <div class="${styles.app}">
            <div class="${styles.bottom}">
              <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape(
                bottomString
              )}
            </div>
          </div>`;
        }
      }
    }
  }

  private _connect(): void {
    let headers = new Headers();
    headers.append("authorization", "Bearer " + this._token);
    headers.append("accept", "application/json"); 

    this.context.httpClient.get('https://jq-webapp1.azurewebsites.net/api/Values', HttpClient.configurations.v1, { headers: headers })
      .then((response: HttpClientResponse): Promise<string> => {
        // output the response from the connection  
        alert("You have conncected to the API as: " + this.context.pageContext.user.email); 
        alert("The url of this page is: " + this.context.pageContext.site.absoluteUrl);
        return response.json();
      });
  }

  public render() {
  }
}
