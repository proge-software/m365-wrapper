
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

This library is focused on wrapping Microsoft 365 API call merging together authentication features and Graph API call.
Authentication feature is driven by MSAL official library which this package depends on.

## OAuth 2.0 and the Implicit Flow

This library, like Msal, implements the [Implicit Grant Flow](https://docs.microsoft.com/azure/active-directory/develop/v2-oauth2-implicit-grant-flow), as defined by the OAuth 2.0 protocol and is [OpenID](https://docs.microsoft.com/azure/active-directory/develop/v2-protocols-oidc) compliant.

## Usage

### Authentication

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
await organizationsClient.StatLoginPopupProcess();
``` 

Logout (with account choice)
```
await organizationsClient.logout();
``` 

### User Info

Get logged user details (output type: MicrosoftGraph.User)
```
const userDetails = await organizationsClient.GetMyDetails();
``` 

Get data of a user (output type: MicrosoftGraph.User)
```
var userIdOrEmail = "<xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx> | <userEmail>";     // A valid user id or email (required).
const returnedUser = await organizationsClient.GetUserByIdOrEmail(userIdOrEmail);
``` 

Get logged user events (output type: collection of {subject, organizer, attendees, start, end, location, onlineMeeting, bodyPreview, webLink, body})
```
const userEvents = await organizationsClient.GetMyEvents();
``` 

Get logged user joined teams (output type: collection of MicrosoftGraph.Team)
```
const joinedTeams = await organizationsClient.GetMyJoinedTeams();
``` 

Get users from your organization (output type: collection of MicrosoftGraph.User)
```
const myOrgUsers = await organizationsClient.GetUsers();
``` 

### Teams info

Get data of the specified team (output type: MicrosoftGraph.Team)
```
var teamGroupId = "<xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx>";     // A valid Teams group unique id (required).
const returnedTeam = await organizationsClient.GetTeam(teamGroupId);
``` 

Get the list of the channels of a team (output type: collection of MicrosoftGraph.Channel)
```
var teamGroupId = "<xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx>";     // A valid Teams group unique id (required).
const teamChannelsList = await organizationsClient.GetTeamChannels(teamGroupId);
``` 

Get data of a team's channel (output type: MicrosoftGraph.Channel)
```
var teamGroupId = "<xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx>";     // A valid Teams group unique id (required).
var channelId = "<channelId>";                                  // A valid channel unique id (required).
const returnedChannel = await organizationsClient.GetTeamChannel();
``` 

Get a list of the group's direct members (output type: collection of MicrosoftGraph.DirectoryObject)
```
var teamGroupId = "<xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx>";     // A valid Teams group unique id (required).
const teamMembersList = await organizationsClient.GetTeamMembers(teamGroupId);
``` 

Get a list with the group's events (output type: collection of MicrosoftGraph.Event)
```
var teamGroupId = "<xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx>";     // A valid Teams group unique id (required).
const teamEventsList = await organizationsClient.GetTeamEvents(teamGroupId);
``` 

### Teams meeting

Create online meeting (note: the meeting does not show up on the user's calendar. Output type: MicrosoftGraph.OnlineMeeting)
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

Create outlook calendar event (output type: MicrosoftGraph.Event)
```
var outlCalEvent = {
    subject: "Outlook calendar event subject",
    body: {
        contentType: "HTML",
        content: "Some text message."               // Text message content.
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
    attendees: [                                    // Fill with all attendees needed followed by a comma exept the last one
        {
            emailAddress: {
                address: "name1@mydomain.com",
                name: "Name Surmane"
            },
            type: "required"                        // Possible admitted values are required, optional, resource.
        },
        {
            emailAddress: {
                address: "name2@mydomain.com",
                name: "Name Surname"
            },
            type: "optional"                        // Possible admitted values are required, optional, resource.
        }
    ],
    allowNewTimeProposals: true,                // Optional. True if the meeting organizer allows invitees to propose a new time when responding, false otherwise. Default is true.
    isOnlineMeeting: true,      // Optional. True if this event has online meeting information (that is, onlineMeeting points to an onlineMeetingInfo resource), false otherwise. 
                                // After you set isOnlineMeeting to true, onlineMeeting is initialized. Subsequently Outlook ignores any further changes to isOnlineMeeting, and the 
                                // meeting remains available online. Default is false (onlineMeeting is null).
    onlineMeetingProvider: "teamsForBusiness",      // Optional. Online meeting service provider. Possible values are unknown, teamsForBusiness, skypeForBusiness, and skypeForConsumer. 
                                                    // After you set onlineMeetingProvider, onlineMeeting is initialized. Subsequently you cannot change onlineMeetingProvider again, and 
                                                    // the meeting remains available online. Default is unknown.
    categories: [               // Optional. Displayed name of one or more Outlook categories (defined for the user) to associate with the event.
        "Orange Category", 
        "Purple Category",
        "Blue Category"
    ],
    importance: "normal",       // Optional. Importance of the event. Possible values are: low, normal, high. Default is normal.
    isAllDay: "false"           // Optional. Set to true if the event lasts all day. If true, regardless of whether it's a single-day or multi-day event, start 
                                // and end time must be set to midnight (period must be at least 24 hours long) and be in the same time zone. Default is false.    
};
const outCalEvent = await organizationsClient.CreateOutlookCalendarEvent(outlCalEvent);
``` 

Update outlook calendar event attendees (output type: MicrosoftGraph.Event)
```
var eventId = "<eventId>";                          // A valid outlook calendar event id (required).
var newAtteendees = {
    attendees: [                                    // Fill with all attendees needed for the event, followed by a comma exept the last one
        {
            emailAddress: {
                address: "name1@mydomain.com",
                name: "Name Surmane"
            },
            type: "required"                        // Possible admitted values are required, optional, resource.
        },
        {
            emailAddress: {
                address: "name2@mydomain.com",
                name: "Name Surname"
            },
            type: "optional"                        // Possible admitted values are required, optional, resource.
        }
    ]
};
const outCalEvent = await organizationsClient.UpdateOutlookCalendarEventAttendees(eventId, newAtteendees);
``` 

### One drive

Enumerate OneDrive resources available to the logged user (output type: collection of MicrosoftGraph.Drive)
```
const myDrives = await organizationsClient.GetMyDrives();
```

Search, within the drive of the logged user, the hierarchy of items for items matching a query (output type: collection of MicrosoftGraph.DriveItem)
```
var searchText = "<Text to search>";    // Optional. The query text used to search for items. Values may be matched
                                        // across several fields including filename, metadata, and file content.
const driveItems = await organizationsClient.GetMyDriveItemsByQuery(searchText);
```

Get DriveItems searching for items within both logged user drive and items shared with him (output type: collection of MicrosoftGraph.DriveItem)
```
var searchText = "<Text to search>";    // Optional. The query text used to search for items. Values may be matched
                                        // across several fields including filename, metadata, and file content.
const driveItems = await organizationsClient.GetMyDriveAndSharedItemsByQuery(searchText);
```

Retrieve a collection of DriveItem resources that have been shared with the logged user (output type: collection of MicrosoftGraph.DriveItem)
```
const driveItems = await organizationsClient.GetMySharedItems();
```

Get a DriveItem resource (also a shared one). To access a shared DriveItem resource, the request can be made using the 
parameters provided in 'remoteItem' facet returned by the GetMySharedItems() method (output type: MicrosoftGraph.DriveItem)
```
var driveId = "<driveId>";  // A valid drive unique id (required).
var itemId = "<itemId>";    // A valid DriveItem id (required).
const item = await organizationsClient.GetDriveItem(driveId, itemId);
```

Enumerate OneDrive resources available to the team group (output type: collection of MicrosoftGraph.Drive)
```
var teamGroupId = "<xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx>";     // A valid Teams group unique id (required).
const driveItems = await organizationsClient.GetTeamDrives(teamGroupId);
```

Enumerate the Drives (document libraries) under the given SharePoint site (output type: collection of MicrosoftGraph.Drive)
```
var siteIdOrName = "<siteIdOrName>";        // A valid sharepoint site name or id (required; site name example: contoso.sharepoint.com).
const driveItems = await organizationsClient.GetSiteDrives(siteIdOrName);
```

Search, within the drive of the given SharePoint site, the hierarchy of items for items matching a query (output type: collection of MicrosoftGraph.DriveItem)
```
var siteIdOrName = "<siteIdOrName>";    // A valid sharepoint site name or id (required; site name example: contoso.sharepoint.com).
var searchText = "<Text to search>";    // Optional. The query text used to search for items. Values may be matched
                                        // across several fields including filename, metadata, and file content.
const driveItems = await organizationsClient.GetSiteDriveItemsByQuery(siteIdOrName, searchText);
```

Enumerate the DriveItem resources in the root of a specific OneDrive resource (output type: collection of MicrosoftGraph.DriveItem)
```
var driveId = "<driveId>";      // A valid drive unique id (required).
const driveItems = await organizationsClient.GetDriveItems(driveId);
```

Search, within the given OneDrive resource, the hierarchy of items for items matching a query (output type: collection of MicrosoftGraph.DriveItem)
```
var driveId = "<driveId>";              // A valid drive unique id (required).
var searchText = "<Text to search>";    // Optional. The query text used to search for items. Values may be matched
                                        // across several fields including filename, metadata, and file content.
const driveItems = await organizationsClient.GetDriveItemsByQuery(driveId, searchText);
```

Enumerate the DriveItems resources in the folder of a specific OneDrive resource (output type: collection of MicrosoftGraph.DriveItem)
```
var driveId = "<driveId>";      // A valid drive unique id (required).
var folderId = "<folderId>";    // A valid folder id (required).
const folderItems = await organizationsClient.GetDriveFolderItems(driveId, folderId);
```

Access a Teams group default document library and get the list of the children of a DriveItem by root relative path (output type: collection of MicrosoftGraph.DriveItem)
```
var teamGroupId = "<xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx>";     // A valid Teams group unique id (required).
var relPath = "/General/MySpecificFolder";                      // Optional. Relative path (the slash ("/") at the beginning and/or at the end can be specified or omitted).
const driveItemContentList = await organizationsClient.GetTeamDefaultDriveItems(teamGroupId, relPath);
``` 

Search, within the given Teams group, the hierarchy of items for items matching a query (output type: collection of MicrosoftGraph.DriveItem)
```
var teamGroupId = "<xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx>";     // A valid Teams group unique id (required).
var searchText = "<Text to search>";                            // Optional. The query text used to search for items. Values may be matched
                                                                // across several fields including filename, metadata, and file content.
const driveItems = await organizationsClient.GetTeamDriveItemsByQuery(teamGroupId, searchText);
```

Get the list of the effective sharing permissions on a driveItem (among the ones of the driveItems of the currently logged in user).
(output type: MicrosoftGraph.Permission)
```
var itemId = "<itemId>";    // Valid id of a driveItem of the currently logged in user (required).
const item = await organizationsClient.GetMyDriveItemSharingPermissions(itemId);
```
