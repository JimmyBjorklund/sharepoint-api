"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.SharepointClient = void 0;
const axios_1 = require("axios");
const qs = require("qs");
/**
 * This is a simple example of how to use the Microsoft Graph API to upload a file to a SharePoint site.
 *
 * For this to work you need to create a new app in Azure AD and give it the following permissions:
 * Microsoft Graph
 * - Directory.ReadWrite.All
 * - Files.ReadWrite.All
 *
 * URL: https://entra.microsoft.com/
 *
 * Sample code:
 *   // This is the name of the tenant in SharePoint, normally first part of the url.
 *   // sample: myName.sharepoint.com
 *   const tenantName: "myName";
 *   const tenantId: "00000000-0000-0000-0000-000000000000";
 *
 *   // This is the name of the site (group in teams) in SharePoint.
 *   const siteName: "MyGroup";
 *
 *   // This is the id and secret of the app you created in Azure AD.
 *   const clientId: "00000000-0000-0000-0000-000000000000";
 *   const clientSecret: "0000000000000000000000000000000000000000";
 *
 *   const client = new SharepointClient({
 *                             tenantId,
 *                             tenantName,
 *                             siteName,
 *                             clientId,
 *                             clientSecret
 *                           });
 *
 *   const token = await client.login();
 *   if (token === undefined) {
 *     console.log("Error getting token");
 *     return;
 *   }
 *   const site = await azure.getSite(token);
 *   if (site === undefined) {
 *     console.log("Error getting site");
 *     return;
 *   }
 *   const siteId = client.getSiteId(site);
 *   const driveName = "Dokument"; // Folder name in SharePoint.
 *   const driver = await client.getDrive(token, siteId, driveName);
 *   if (driver === undefined) {
 *     console.log("Error getting drive");
 *     return;
 *   }
 *   const path = "/TestUploadFiles";
 *   const fileName = "test.txt";
 *   const res = await client.upload(token, driver.id, path, fileName, "text/plain");
 *   console.log("Upload response", res);
 */
class SharepointClient {
    tenantId;
    tenantName;
    siteName;
    clientId;
    clientSecret;
    constructor(config) {
        this.tenantId = config.tenantId;
        this.tenantName = config.tenantName;
        this.siteName = config.siteName;
        this.clientId = config.clientId;
        this.clientSecret = config.clientSecret;
    }
    /**
     * Get a token from Azure AD.
     * @returns The token or undefined if an error occurred.
     */
    login = async () => {
        const tokenEndpoint = `https://login.microsoftonline.com/${this.tenantId}/oauth2/v2.0/token`;
        const options = {
            headers: {
                "Content-Type": `application/x-www-form-urlencoded`,
            },
        };
        const data = qs.stringify({
            grant_type: "client_credentials",
            client_id: this.clientId,
            client_secret: this.clientSecret,
            scope: "https://graph.microsoft.com/.default",
        });
        const tokenResponse = await axios_1.default
            .post(tokenEndpoint, data, options)
            .then((response) => {
            if (response.status !== 200) {
                console.log("Error getting token", response);
                return undefined;
            }
            return response.data;
        })
            .catch((error) => {
            console.log("Error getting token", error);
            return undefined;
        });
        return tokenResponse;
    };
    /**
     * Get a specific site from SharePoint.
     * @param token - The token from the login method.
     * @returns The site or undefined if an error occurred.
     */
    getSite = async (token) => {
        const url = `https://graph.microsoft.com/v1.0/sites/${this.tenantName}.sharepoint.com:/sites/${this.siteName}`;
        const options = {
            headers: {
                Authorization: `Bearer ${token.access_token}`,
            },
        };
        return axios_1.default
            .get(url, options)
            .then((response) => {
            return response.data;
        })
            .catch((error) => {
            console.log("Error getting site", error);
        });
    };
    /**
     * Get all drives in a site.
     * @param token - The token from the login method.
     * @param siteId - The id of the site.
     * @returns The drives or undefined if an error occurred.
     */
    getDrives = async (token, siteId) => {
        const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/Drives`;
        const options = {
            headers: {
                Authorization: `Bearer ${token.access_token}`,
            },
        };
        return axios_1.default
            .get(url, options)
            .then((response) => {
            return response.data?.value;
        })
            .catch((error) => {
            console.log("Error getting drives", error);
        });
    };
    /**
     * Get a specific drive from a site.
     * @param token - The token from the login method.
     * @param siteId - The id of the site.
     * @param driveName - The name of the drive to get.
     * @returns The drive or undefined if an error occurred.
     */
    getDrive = async (token, siteId, driveName) => {
        const drives = await this.getDrives(token, siteId);
        if (drives === undefined) {
            console.log("Error getting drives");
            return undefined;
        }
        const driver = drives.find((drive) => {
            return drive.name === driveName;
        });
        return driver;
    };
    /**
     * Get items (files, directories etc) in a drive.
     * @param token - The token from the login method.
     * @param driverId - The id of the drive to get items from.
     * @param path - The path to the items in from.
     * @returns The items or undefined if an error occurred.
     */
    getItems = async (token, driverId, path) => {
        const url = `https://graph.microsoft.com/v1.0/Drives/${driverId}/root:/${path}:/Children.`;
        const options = {
            headers: {
                Authorization: `Bearer ${token.access_token}`,
            },
        };
        return axios_1.default
            .get(url, options)
            .then((response) => {
            return response.data?.value;
        })
            .catch((error) => {
            console.log("Error getting items", error);
            return undefined;
        });
    };
    /**
     * Download a file from a SharePoint site.
     * @param token - The token from the login method.
     * @param driverId - The id of the drive to download from.
     * @param path - The path to the file in the drive.
     * @returns The file content or undefined if an error occurred.
     */
    downloadItem = async (token, driverId, path) => {
        const url = `https://graph.microsoft.com/v1.0/Drives/${driverId}/root:/${path}:/content`;
        const options = {
            headers: {
                Authorization: `Bearer ${token.access_token}`,
            },
        };
        return axios_1.default
            .get(url, options)
            .then((response) => {
            return response.data;
        })
            .catch((error) => {
            console.log("Error getting item", error);
            return undefined;
        });
    };
    /**
     * Upload or update a file in a SharePoint site.
     * @param token - The token from the login method.
     * @param driverId - The id of the drive to upload to.
     * @param path - The path to the file in the drive.
     * @param fileName - The name of the target file.
     * @param contentType - The content type of the file.
     * @param data - The file content as a buffer.
     * @returns The uploaded item or undefined if an error occurred.
     */
    upload = async (token, driverId, path, fileName, contentType, data) => {
        const url = `https://graph.microsoft.com/v1.0//drives/${driverId}/root:${path}/${fileName}:/content`;
        const options = {
            headers: {
                "Content-Type": contentType,
                Authorization: `Bearer ${token.access_token}`,
            },
        };
        //const binaryContent = new Blob(["Hello, World!"], { type: contentType });
        return axios_1.default
            .put(url, data, options)
            .then((response) => {
            console.log("Upload response", response.status, response.statusText);
            return response.data;
        })
            .catch((error) => {
            console.log("Error uploading", error.response);
            return undefined;
        });
    };
    getSiteId = (site) => {
        return site.id.split(",")[1];
    };
}
exports.SharepointClient = SharepointClient;
//# sourceMappingURL=sharepoint-client.js.map