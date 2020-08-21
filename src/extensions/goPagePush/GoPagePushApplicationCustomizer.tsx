import * as React from "react";
import { ReactElement } from "react";
import * as ReactDOM from "react-dom";
import { override } from "@microsoft/decorators";
import { Log } from "@microsoft/sp-core-library";
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName,
} from "@microsoft/sp-application-base";
import { Dialog } from "@microsoft/sp-dialog";
import { Button } from "office-ui-fabric-react/lib/Button";

import * as strings from "GoPagePushApplicationCustomizerStrings";
import styles from "./AppCustomizer.module.scss";
import { escape } from "@microsoft/sp-lodash-subset";

import { AadHttpClient, HttpClientResponse } from "@microsoft/sp-http"; // added
const LOG_SOURCE: string = "GoPagePushApplicationCustomizer";

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGoPagePushApplicationCustomizerProperties {
  testMessage: string;
  Top: string;
  Bottom: string;
  //ButtonLabel: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class GoPagePushApplicationCustomizer extends BaseApplicationCustomizer<
  IGoPagePushApplicationCustomizerProperties
> {
  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);
    // Wait for the placeholders to be created (or handle them being changed) and then
    // render.
    this.context.placeholderProvider.changedEvent.add(
      this,
      this._renderPlaceHolders
    );
    // this.aadHttpClient = await this.context.aadHttpClientFactory
    //   .getClient('https://jq-webapp1.azurewebsites.net');
    this._addButton();
    return Promise.resolve();
  }
  private _renderReactElement(
    component: ReactElement<any>,
    node: Element
  ): void {
    ReactDOM.unmountComponentAtNode(node);
    ReactDOM.render(component, node);
  }

  // for the api button
  private _addButton() {
    const button: React.ReactElement<IGoPagePushApplicationCustomizerProperties> = (
      <div className="button">
      <React.Fragment>
        <Button onClick={this._connect}> CLICK HERE TO CONNECT TO THE API</Button>
      </React.Fragment>
      </div>
    );
    this._renderReactElement(button, this._bottomPlaceholder.domElement);
  }

  // placeholder private method (added)
  private _renderPlaceHolders(): void {
    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Top,
        { onDispose: this._onDispose }
      );

      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error("The expected placeholder (Top) was not found.");
        return;
      }

      if (this.properties) {
        let topString: string = this.properties.Top;
        if (!topString) {
          topString = "(Top property was not defined.)";
        }

        if (this._topPlaceholder.domElement) {
          this._topPlaceholder.domElement.innerHTML = `
          <div class="${styles.app}">
            <div class="${styles.top}">
              <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i> ${escape(
                topString
              )}
            </div>
          </div>`;
        }
      }
    }

    // Handling the bottom placeholder
    if (!this._bottomPlaceholder) {
      this._bottomPlaceholder = this.context.placeholderProvider.tryCreateContent(
        PlaceholderName.Bottom,
        { onDispose: this._onDispose }
      );

      // The extension should not assume that the expected placeholder is available.
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

  // added for the placeholder
  private _onDispose(): void {
    console.log(
      "[GoPagePushApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders."
    );
  }

  private _connect(): void {
    console.log("connect called");
    this.context.aadHttpClientFactory
    .getClient('https://jq-webapp1.azurewebsites.net')
    .then((client: AadHttpClient): void => {
      // connect to the API
      // process data

    });
    //this.context.httpClient.get('https://jq-webapp1.azurewebsites.net', 
    // HttpClient.configurations.v1, {
    //   credentials:"include"
    // })
    // .then((response: HttpClientResponse): Promise<string> => {
    //   return response.json();
    // })
    // .then((greeting: string): void => {
    //   this.domElement.querySelector(".greeting").innerHTML = greeting;
    //})
  }

  public render() {
    // <React.Fragment>
    //   <div>
    //     <Button>click me</Button>
    //   </div>
    // </React.Fragment>;
  }
}
