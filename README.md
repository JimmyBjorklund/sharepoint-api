sharepoint-api
====

[![NPM](https://nodei.co/npm/@jimmybjorklund/sharepoint-api.png)](https://nodei.co/npm/@jimmybjorklund/sharepoint-api/)

## Introduction
This client makes it easier to upload/download and list files on your Microsoft SharePoint account.

It uses the Microsoft Graph api to handle the communication and you need to create an 
client access token for this api for this library to work.

## Installation

```
$ npm install sharepoint-client
```

## Create access.

### 1. First we need to create a app that is allowed to access the files. 
   
   Navigate to: https://entra.microsoft.com/

   Once you have logged in navigate in the menu.

        Application->App registrations.

   On this page select 

        + New registration.

   Give the integration app a good name and select single tenant and press Register.

   You should now be presented with the Application overview, on this page search for
   Application (client) ID <b>save this id for later as clientId.</b>

   On this page you should also be able to find the Directory (tenant) ID, <b>save this as
   tenantId for later.</b>


### 2. Create a client secret this is done by navigating to Certificates & secrets under you app view.

   On this page click

        + New client secret 

   Enter a good description and max lifetime of you tokens and press Add.

   In the list of secrets you should now see the Value of your secret key, <b>save this into 
   a variable called clientSecret.</b> NOTE: Do this directly as it will only be visible once.
   If you fail to do this you can remove it an recreated it to se a new value.

### 3. Setup Api permissions by going to API permissions.
   
   On this page click.

        + Add a permission

   You should select the api that say <b>Microsoft Graph.</b>
   You will be presented with two options delegated access or application.
    <b>Select application permissions.</b>

   You need to select:

        Directory.ReadWrite.All
        Files.ReadWrite.All
  
### Congratulations
You should now have a clientId and secret that will give you access.



## Sample
  This is a simple example of how to use the Microsoft Graph API to upload a file to a SharePoint site.
  
  ```ts
//This is the name of the tenant in SharePoint, normally first part of the url.
//sample: myName.sharepoint.com  
const tenantName: "myName";
// From App Overview page.
const tenantId: "00000000-0000-0000-0000-000000000000";
  
// This is the name of the site (group in teams) in SharePoint.
const siteName: "MyGroup";
   
// This is the id and secret of the app you created in Azure AD.
const clientId: "00000000-0000-0000-0000-000000000000";
const clientSecret: "0000000000000000000000000000000000000000";
  
const client = new SharepointApi({
                                   tenantId,
                                   tenantName,
                                   siteName,
                                   clientId,
                                   clientSecret
                                  });
// Get a Azure AD login token.                                    
const token = await client.login();
if (token === undefined) {
    console.log("Error getting token");
    return;
}

// Fetch the site to get the siteId.
const site = await azure.getSite(token);
if (site === undefined) {
    console.log("Error getting site");
    return;
}
const siteId = client.getSiteId(site);
const driveName = "Dokument"; // Folder name in SharePoint.

// Get the driver to find the correct driverId.
const driver = await client.getDrive(token, siteId, driveName);
if (driver === undefined) {
    console.log("Error getting drive");
    return;
}

const path = "/TestUploadFiles";
const fileName = "test.txt";
// Upload files.
const res = await client.upload(token, driver.id, path, fileName, "text/plain");
console.log("Upload response", res);
```

## LICENSE

MIT, see LICENSE.md file.