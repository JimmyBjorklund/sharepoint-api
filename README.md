# SharePoint API Client

[![NPM](https://img.shields.io/npm/v/@jimmybjorklund/sharepoint-api)](https://www.npmjs.com/package/@jimmybjorklund/sharepoint-api)
[![License](https://img.shields.io/npm/l/@jimmybjorklund/sharepoint-api)](LICENSE.md)

A lightweight Node.js client for interacting with Microsoft SharePoint via the Microsoft Graph API. This library simplifies file operations such as uploading, downloading, and listing files.

## Installation

```bash
npm install @jimmybjorklund/sharepoint-api
```

## Prerequisites

To use this library, you need to register an application in Microsoft Entra ID (formerly Azure AD) to obtain credentials.

### 1. Register an Application
1. Go to the [Microsoft Entra admin center](https://entra.microsoft.com/).
2. Navigate to **Applications** > **App registrations**.
3. Click **+ New registration**.
4. Enter a name for your app (e.g., "SharePoint Integration"), select **Single tenant**, and click **Register**.

### 2. Get Credentials
After registration, you will be taken to the application overview page.
- **Client ID**: Copy the `Application (client) ID`.
- **Tenant ID**: Copy the `Directory (tenant) ID`.

### 3. Create a Client Secret
1. In the app menu, go to **Certificates & secrets**.
2. Click **+ New client secret**.
3. Enter a description and expiration, then click **Add**.
4. **Important**: Copy the `Value` immediately. This is your `clientSecret`.

### 4. Configure API Permissions
1. In the app menu, go to **API permissions**.
2. Click **+ Add a permission** > **Microsoft Graph** > **Application permissions**.
3. Select and add the following permissions:
   - `Directory.ReadWrite.All`
   - `Files.ReadWrite.All`
4. Click **Grant admin consent** for your organization.

## Usage

Here is a complete example of how to upload a file to a SharePoint site.

```typescript
import { SharepointApi } from '@jimmybjorklund/sharepoint-api';
import * as fs from 'fs';

async function main() {
  // Configuration
  const config = {
    tenantName: "your-tenant-name", // e.g. "mycompany" from mycompany.sharepoint.com
    tenantId: "00000000-0000-0000-0000-000000000000",
    siteName: "MySharePointSite",
    clientId: "00000000-0000-0000-0000-000000000000",
    clientSecret: "your-client-secret-value"
  };

  const client = new SharepointApi(config);

  // 1. Authenticate
  const token = await client.login();
  if (!token) {
    console.error("Failed to authenticate.");
    return;
  }

  // 2. Get the Site ID using the site name
  const site = await client.getSite(token);
  if (!site) {
    console.error("Failed to get site information.");
    return;
  }
  const siteId = client.getSiteId(site);

  // 3. Get the Drive (Document Library)
  // 'Documents' is the default document library name in many locales, 
  // but it might be 'Dokument' or 'Shared Documents' depending on your setup.
  const driveName = "Documents"; 
  const drive = await client.getDrive(token, siteId, driveName);
  if (!drive) {
    console.error(`Failed to find drive with name: ${driveName}`);
    return;
  }

  // 4. Upload a File
  const remotePath = "/UploadedFiles"; // Folder path in SharePoint
  const fileName = "example.txt";
  const fileContent = Buffer.from("Hello SharePoint!", "utf-8"); // Or read from fs: fs.readFileSync('./local.txt')
  
  console.log(`Uploading ${fileName}...`);
  const uploadedItem = await client.upload(
    token,
    drive.id,
    remotePath,
    fileName,
    "text/plain",
    fileContent
  );

  if (uploadedItem) {
    console.log("Upload successful:", uploadedItem.webUrl);
  } else {
    console.log("Upload failed.");
  }
}

main();
```

## License

MIT. See [LICENSE.md](LICENSE.md) for details.

