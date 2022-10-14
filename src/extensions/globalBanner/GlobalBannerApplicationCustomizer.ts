import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { Configuration } from "./banner-config";
import { ContextInfo } from "gd-sprest";

import * as strings from 'GlobalBannerApplicationCustomizerStrings';

/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface IGlobalBannerApplicationCustomizerProperties {
  fileUrl: string;
  webUrl: string;
}

/** A Custom Action which can be run during execution of a Client Side Application */
export default class GlobalBannerApplicationCustomizer
  extends BaseApplicationCustomizer<IGlobalBannerApplicationCustomizerProperties> {

  // Global Variable
  private _header: PlaceholderContent = null;

  public onInit(): Promise<void> {
    // Log
    Log.info(strings.LOG_KEY, `Initializing the ${strings.Title} solution.`);

    // Set the page context information in the library
    ContextInfo.setPageContext(this.context.pageContext);

    // Handle possible changes to the placeholders
    this.context.placeholderProvider.changedEvent.add(this, this.renderBanner);

    // Render the banner
    this.renderBanner();

    return Promise.resolve();
  }

  // Renders the banner
  private renderBanner() {
    // See if the header doesn't exist
    if (this._header === null) {
      // Log
      Log.info(strings.LOG_KEY, "Creating the banner.");

      // Create the header
      this._header = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top);

      // Log
      Log.info(strings.LOG_KEY, "Loading the configuration.");

      // Load the configuration file
      Configuration.load(this.properties.webUrl, this.properties.fileUrl).then(
        // Success
        () => {
          // See if this site is customized
          let cfg = Configuration.Sites[ContextInfo.webServerRelativeUrl];
          if (cfg) {
            // Render the banner
            this._header.domElement.innerHTML = cfg.html;
            this._header.domElement.style.backgroundColor = cfg.color;
          } else {
            // Render the default banner
            this._header.domElement.innerHTML = "This is the default Banner";
            this._header.domElement.style.backgroundColor = "lightBlue";
          }

          // Update the styling
          this._header.domElement.style.textAlign = "center";

          // Log
          Log.info(strings.LOG_KEY, "Banner rendered successfully.");
        }
      );
    }
  }
}