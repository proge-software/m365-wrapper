
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

Evaluate if the user has already logged and acquire token silently
```
await z.StatLoginPopupProcess();
``` 

Get logged user detail
```
const userDetail = await organizationsClient.GetUserDetail();
``` 

Get user joined teams
```
const joinedTeams = await organizationsClient.GetUserJoinedTeams();
``` 

Create online meeting
```
var meeting = {
    subject: "Online meeting subject",
    startDateTime: "2020-05-28T11:00:00.0000000-00:00",
    endDateTime: "2020-05-28T13:30:00.0000000-00:00",
    participants: {
        attendees: [                                // Note: fill with all attendees needed followed by a comma exept the last one.
            { upn: "name1@mydomain.com" }, 
            { upn: "name2@mydomain.com" }, 
            { upn: "name3@mydomain.com" }],   
        organizer: { upn: "name4@mydomain.com" }    // Note: organizer is optional (if not specified, the comma at the end of the above line also must be omitted).
    }
};
const onlineMeeting = await organizationsClient.CreateOnlineMeeting(meeting);
``` 

Create outlook calendar event
```
var outlCalEvent = {
    subject: "Outlook calendar event subject",
    body: {
        contentType: "HTML",
        content: "Some text message."               // Note: text message content.
    },
    start: {
        dateTime: "2020-05-29T16:00:00",
        timeZone: "W. Europe Standard Time"         // Possible admitted values can be found at https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/default-time-zones
    },
    end: {
        dateTime: "2020-05-29T17:30:00",
        timeZone: "W. Europe Standard Time"         // Possible admitted values can be found at https://docs.microsoft.com/en-us/windows-hardware/manufacture/desktop/default-time-zones
    },
    location: {
        displayName: "Online on Teams"
    },
    attendees: [                                    // Note: fill with all attendees needed followed by a comma exept the last one
        {
            emailAddress: {
                address: "name1@mydomain.com",
                name: "Name Surmane"
            },
            type: "required"                        // Note: Possible admitted values are required, optional, resource.
        },
        {
            emailAddress: {
                address: "name2@mydomain.com",
                name: "Name Surname"
            },
            type: "optional"                        // Note: Possible admitted values are required, optional, resource.
        }
    ],
    allowNewTimeProposals: true,
    isOnlineMeeting: true,
    onlineMeetingProvider: "teamsForBusiness"
}
const outCalEvent = await organizationsClient.CreateOutlookCalendarEvent(outlCalEvent);
``` 

