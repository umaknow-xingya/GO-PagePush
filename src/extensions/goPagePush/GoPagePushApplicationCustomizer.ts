import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer
} from '@microsoft/sp-application-base';
import { Dialog } from '@microsoft/sp-dialog';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http'; // added 

import * as strings from 'GoPagePushApplicationCustomizerStrings';

const LOG_SOURCE: string = 'GoPagePushApplicationCustomizer';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGoPagePushApplicationCustomizerProperties {
  // This is an example; replace with your own property
  testMessage: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class GoPagePushApplicationCustomizer
  extends BaseApplicationCustomizer<IGoPagePushApplicationCustomizerProperties> {

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, `Initialized ${strings.Title}`);

    let message: string = this.properties.testMessage;
    if (!message) {
      message = '(No properties were provided.)';
    }

    Dialog.alert(`Hello from ${strings.Title}:\n\n${message}`);

    return Promise.resolve();
  }

  public render(): void {
    this.context.aadHttpClientFactory
      .getClient('https://contoso.azurewebsites.net')
      .then((client: AadHttpClient): void => {
        // connect to the API

      });
  }
}
