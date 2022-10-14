import { Log } from '@microsoft/sp-core-library';
import { Web } from "gd-sprest";
import * as strings from 'GlobalBannerApplicationCustomizerStrings';

// Site Configuration
export interface ISiteConfiguration {
    color?: string;
    html?: string;
}

/**
 * Configuration
 */
export class Configuration {
    private static _config: { [key: string]: ISiteConfiguration } = null;
    public static get Sites(): { [key: string]: ISiteConfiguration } { return this._config; }

    // Loads the configuration file
    public static load(webUrl: string, fileUrl: string) {
        // Return a promise
        return new Promise((resolve) => {
            // See if the cache has the information
            let cacheData = sessionStorage.getItem(strings.CACHE_KEY);
            if (cacheData) {
                try {
                    // Log
                    Log.info(strings.LOG_KEY, "Loading data from cache...");

                    // Try to convert the data
                    this._config = JSON.parse(cacheData);

                    // Log
                    Log.info(strings.LOG_KEY, "Configuration was set from cache...");

                    // Resolve the request
                    resolve(null);
                    return;
                }
                catch (ex) { this._config = {}; }
            }

            // Log
            Log.info(strings.LOG_KEY, `Loading the configuration file from ${webUrl}/${fileUrl}.`);

            // Load the web
            Web(webUrl).getFileByUrl(fileUrl).content().execute(
                // Success
                content => {
                    // Convert the string to a json object
                    try {
                        // Log
                        Log.info(strings.LOG_KEY, "Configuration file read. Converting it to a JSON object.");

                        // Convert the file content to a string
                        let cfgContent = String.fromCharCode.apply(null, new Uint8Array(content));
                        this._config = JSON.parse(cfgContent);

                        // Log
                        Log.info(strings.LOG_KEY, "Saving the configuration file to cache.");

                        // Save the content to cache
                        sessionStorage.setItem(strings.CACHE_KEY, cfgContent);

                        // Log
                        Log.info(strings.LOG_KEY, "Configuration file stored in cache successfully.");
                    }
                    catch (ex) { this._config = {}; }

                    // Resolve the request
                    resolve(null);
                },

                // Error
                err => {
                    // Default the configuration
                    this._config = {};

                    // Log
                    Log.info(strings.LOG_KEY, "Configuration file not found.");

                    // Log
                    Log.info(strings.LOG_KEY, "Updating the cache to store nothing.");

                    // Set the session storage
                    sessionStorage.setItem(strings.CACHE_KEY, JSON.stringify({}));

                    // Log
                    Log.info(strings.LOG_KEY, "Configuration file stored in cache successfully.");

                    // File may not exist
                    resolve(null);
                }
            );
        });
    }
}