import axios from "axios";
import * as qs from "qs";

export class SharepointApi {
  private tenantId: string;
  private tenantName: string;
  private siteName: string;
  private clientId: string;
  private clientSecret: string;
  constructor(config: { tenantId: string; tenantName: string; siteName: string; clientId: string; clientSecret: string }) {
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
  public login = async (): Promise<Azure.Token | undefined> => {
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
    const tokenResponse = await axios
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
  public getSite = async (token: Azure.Token): Promise<SharepointApi.Site | undefined> => {
    const url = `https://graph.microsoft.com/v1.0/sites/${this.tenantName}.sharepoint.com:/sites/${this.siteName}`;
    const options = {
      headers: {
        Authorization: `Bearer ${token.access_token}`,
      },
    };
    return axios
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
  public getDrives = async (token: Azure.Token, siteId: string): Promise<SharepointApi.Drive[] | undefined> => {
    const url = `https://graph.microsoft.com/v1.0/sites/${siteId}/Drives`;
    const options = {
      headers: {
        Authorization: `Bearer ${token.access_token}`,
      },
    };
    return axios
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
  public getDrive = async (token: Azure.Token, siteId: string, driveName: string): Promise<SharepointApi.Drive | undefined> => {
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
  public getItems = async (token: Azure.Token, driverId: string, path: string): Promise<SharepointApi.Item[] | undefined> => {
    const url = `https://graph.microsoft.com/v1.0/Drives/${driverId}/root:/${path}:/Children.`;
    const options = {
      headers: {
        Authorization: `Bearer ${token.access_token}`,
      },
    };
    return axios
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
  public downloadItem = async (token: Azure.Token, driverId: string, path: string): Promise<any | undefined> => {
    const url = `https://graph.microsoft.com/v1.0/Drives/${driverId}/root:/${path}:/content`;
    const options = {
      headers: {
        Authorization: `Bearer ${token.access_token}`,
      },
    };
    return axios
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
  public upload = async (token: Azure.Token, driverId: string, path: string, fileName: string, contentType: string, data: Buffer): Promise<SharepointApi.Item | undefined> => {
    const url = `https://graph.microsoft.com/v1.0//drives/${driverId}/root:${path}/${fileName}:/content`;
    const options = {
      headers: {
        "Content-Type": contentType,
        Authorization: `Bearer ${token.access_token}`,
      },
    };
    //const binaryContent = new Blob(["Hello, World!"], { type: contentType });
    return axios
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

  public getSiteId = (site: SharepointApi.Site): string => {
    return site.id.split(",")[1];
  };
}
export namespace Azure {
  export interface Token {
    token_type: "Bearer";
    expires_in: number;
    ext_expires_in: number;
    access_token: string;
  }
}

export namespace SharepointApi {
  export interface Site {
    "@odata.context": string;
    createdDateTime: string;
    description: string;
    id: string;
    lastModifiedDateTime: string;
    name: string;
    webUrl: string;
    displayName: string;
    root: any;
    siteCollection: any;
  }
  export interface Identity {
    user?: {
      email?: string;
      id?: string;
      displayName: string;
    };
    group?: {
      email?: string;
      id?: string;
      displayName: string;
    };
    application?: {
      id: string;
      displayName: string;
    };
  }
  export interface Drive {
    createdDateTime: string;
    description: string;
    id: string;
    lastModifiedDateTime: string;
    name: string;
    webUrl: string;
    driveType: string; // 'documentLibrary'
    createdBy: Identity;
    lastModifiedBy: Identity;
    owner: Identity;
    quota: any;
  }

  export interface Item {
    "@odata.context": string;
    "@microsoft.graph.downloadUrl": string;
    createdBy: Identity;
    createdDateTime: string;
    eTag: string;
    id: string;
    lastModifiedBy: Identity;
    lastModifiedDateTime: string;
    name: string;
    parentReference?: {
      driveType: string;
      driveId: string;
      id: string;
      name: string;
      path: string;
      siteId: string;
    };
    webUrl: string;
    cTag: string;
    file?: {
      mimeType: string;
      hashes: {
        quickXorHash: string;
      };
    };
    fileSystemInfo: {
      createdDateTime: string;
      lastModifiedDateTime: string;
    };
    folder?: {
      childCount: number;
    };
    shared: { scope: string };
    size: number;
  }
}
