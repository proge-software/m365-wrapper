
## Installation
### Via NPM:

    npm install m365-wrapper

### Via Latest unpkg CDN Version:

#### Latest compiled and minified JavaScript (US West region)
```html
<script type="text/javascript" src="https://unpkg.com/m365-wrapper/_bundles/m365-wrapper.js"></script>
```
```html
<script type="text/javascript" src="https://unpkg.com/m365-wrapper/_bundles/m365-wrapper.min.js"></script>
````

## What To Expect From This Library

This library is focused on wrapping Microsoft 365 API call merging togheder authentication features and Graph API call.
Authentication feature is driven by MSAL official library which this package depends on.

## OAuth 2.0 and the Implicit Flow

This library, like Msal, implements the [Implicit Grant Flow](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-implicit-grant-flow), as defined by the OAuth 2.0 protocol and is [OpenID](https://docs.microsoft.com/azure/active-directory/develop/v2-protocols-oidc) compliant.

## Usage

Client instance
```
var clientApplicationId = "<xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx>";
var organizationsClient = new M365Wrapper(clientApplicationId);
var allMicrosoftAccountClient = new M365Wrapper(clientId, "https://login.microsoftonline.com/common");
var singleTenantClient = new M365Wrapper(clientId, "https://login.microsoftonline.com/<tenant>/");
```

Login with Popup
```
const authResponse = await organizationsClient.loginPopup();
``` 

Get logged user detail
```
const userDetail = await organizationsClient.GetUserDetail();
``` 

