import "isomorphic-fetch";
import { UserAgentApplication, Configuration, AuthenticationParameters, AuthResponse, Account } from 'msal';
import { Client, ClientOptions, AuthenticationProvider } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { ImplicitMSALAuthenticationProvider, MSALAuthenticationProviderOptions } from "@microsoft/microsoft-graph-client/lib/src/browser";
import UserSearchRequest from "./Types/UserSearchRequest";

//import { ImplicitMSALAuthenticationProvider as NodeImplicitMSALAuthenticationProvider } from "@microsoft/microsoft-graph-client/lib/src/ImplicitMSALAuthenticationProvider";

class M365Wrapper {
  protected authPar: AuthenticationParameters = {
    scopes: ['User.Read', 'User.ReadBasic.All', 
      'Calendars.ReadWrite', 'Calendars.Read.Shared',
      'email', 'Team.ReadBasic.All',  'OnlineMeetings.ReadWrite', 
      'Files.Read.All', 'Group.Read.All', 'Reports.Read.All'],
    prompt: 'select_account',
  };
  protected configuration: Configuration = {
    auth: {
      clientId: '',
      authority: 'https://login.microsoftonline.com/organizations',
    },
    cache: {
      cacheLocation: 'sessionStorage'
    }
  };
  protected GraphScopes: string[] = [...this.authPar.scopes!];
  protected providerOptions: MSALAuthenticationProviderOptions;
  protected msalApplication: UserAgentApplication;
  protected authProvider: AuthenticationProvider;
  protected options: ClientOptions;
  protected client: Client;


  constructor(clientId: string);
  constructor(clientId: string, authority?: string) {
    if (clientId)
      this.configuration.auth.clientId = clientId;

    if (authority)
      this.configuration.auth.authority = authority;

    this.msalApplication = new UserAgentApplication(this.configuration);
    this.providerOptions = new MSALAuthenticationProviderOptions(this.GraphScopes);
    this.authProvider = new ImplicitMSALAuthenticationProvider(this.msalApplication, this.providerOptions);

    this.options = {
      authProvider: this.authProvider
    };

    this.client = Client.initWithMiddleware(this.options);
  }

  public async loginPopup(): Promise<AuthResponse> {
    return await this.msalApplication.loginPopup(this.authPar);
  }

  public async acquireTokenSilent(): Promise<AuthResponse> {
    return await this.msalApplication.acquireTokenSilent(this.authPar);
  }

  public async acquireTokenPopup(): Promise<AuthResponse> {
    return await this.msalApplication.acquireTokenPopup(this.authPar);
  }

  public async getLoginInProgress(): Promise<boolean> {
    return await this.msalApplication.getLoginInProgress();
  }

  public async StatLoginPopupProcess() {
    let account = this.getAccount();
    if (account) {
      await this.acquireTokenSilent().then(async response => {
        //const account = thatMsal.getAccount();
        // this.SET_ACCOUNT(account);
        // this.SET_ID_TOKEN(response);
        // this.SET_LOGIN_STATE(true);       
      }).catch(async error => {
        if (error.errorMessage.indexOf("interaction_required") !== -1) {
          await this.acquireTokenPopup()
            .then(async response => {
              const account = this.getAccount();
              // this.SET_ACCOUNT(account);
              // this.SET_ID_TOKEN(response);
              // this.SET_LOGIN_STATE(true);              
            })
            .catch(err => {
              console.log(err);
            });
        }
        else
          console.log(error);
      })
    }
    else {
      await this.loginPopup()
        .then(async response => {
          await this.acquireTokenSilent().then(async response => {
            account = this.getAccount();
            // this.SET_ACCOUNT(account);
            // this.SET_ID_TOKEN(response);
            // this.SET_LOGIN_STATE(true);            
          })
            .catch(err => {
              console.log(err);
            });
        })
        .catch(err => {
          console.log(err);
        });

    }
  }

  public getAccount(): Account {
    return this.msalApplication.getAccount();
  }

  public logout() {
    this.msalApplication.logout();
  }

  public async GetMyDetails(): Promise<MicrosoftGraph.User> {
    try {
      const userDetails: MicrosoftGraph.User = await this.client.api("/me")
      .get();
      return userDetails;
    } catch (error) {
      throw error;
    }
  }

  public async GetMyEvents(): Promise<[MicrosoftGraph.Event]> {
    try {
      const events = await this.client.api("/me/calendar/events")
        .select('subject,organizer,attendees,start,end,location,onlineMeeting,bodyPreview,webLink,body')
        .get();
      return events;
    } catch (error) {
      throw error;
    }
  }

  public async GetUsers(UserSearchRequest: UserSearchRequest): Promise<MicrosoftGraph.User[]> {
    let query = this.client.api('/users');

    if (UserSearchRequest && UserSearchRequest.issuer && UserSearchRequest.mail) {
      query = query.filter(`identities/any(c:c/issuerAssignedId eq '${UserSearchRequest.mail}' and c/issuer eq '${UserSearchRequest.issuer}')`);
    }

    let res: MicrosoftGraph.User[] = await query.select('displayName,givenName,postalCode,mail,surname,userPrincipalName')
      .get();

    return res;
  }

  public async GetMyJoinedTeams(): Promise<[MicrosoftGraph.Team]> {
    try {
      const teams = await this.client.api("/me/joinedTeams")
        .select('Id,displayName,description')
        .get();
      return teams;
    }
    catch (error) {
      throw error;
    }
  }

  public async CreateOnlineMeeting(onlineMeeting: MicrosoftGraph.OnlineMeeting): Promise<[MicrosoftGraph.OnlineMeeting]> {

    let res: [MicrosoftGraph.OnlineMeeting] = await this.client.api('/me/onlineMeetings')
      .post(onlineMeeting);

    return res;
  }

  public async CreateOutlookCalendarEvent(userEvent: MicrosoftGraph.Event): Promise<[MicrosoftGraph.Event]> {
    //POST /users/{id | userPrincipalName}/calendar/events   <<< Da provare

    let res: [MicrosoftGraph.Event] = await this.client.api('/me/events')
      .post(userEvent);

    return res;
  }

  public async UpdateOutlookCalendarEventAttendees(eventId: string, newAtteendees: string): Promise<MicrosoftGraph.Event> {    
    try {
      let res: MicrosoftGraph.Event = await this.client.api(`/me/events/${eventId}`)
        .patch(newAtteendees);
  
      return res;
      }
    catch (error) {
      throw error;
    }
  }

  public async GetMyDrives(): Promise<[MicrosoftGraph.Drive]> {
    try {
      const items = await this.client.api("/me/drives")
        .get();
      
      return items;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetMyDriveItemsByQuery(queryText: string): Promise<[MicrosoftGraph.DriveItem]> {
    try {
      const items = await this.client.api(`/me/drive/root/search(q='${queryText}')`) 
        .get();
      
      return items;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetMyDriveAndSharedItemsByQuery(queryText: string): Promise<[MicrosoftGraph.DriveItem]> {
    try {
      const items = await this.client.api(`/me/drive/search(q='${queryText}')`) 
        .get();
      
      return items;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetMySharedItems(): Promise<[MicrosoftGraph.DriveItem]> {
    try {
      const items = await this.client.api(`/me/drive/sharedWithMe`) 
        .get();
      
      return items;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetTeamDrives(teamGroupId: string): Promise<[MicrosoftGraph.Drive]> {
    try {
      const items = await this.client.api(`/groups/${teamGroupId}/drives`)
        .get();
      
      return items;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetSiteDrives(siteIdOrName: string): Promise<[MicrosoftGraph.Drive]> {
    try {
      const items = await this.client.api(`/sites/${siteIdOrName}/drives`)
        .get();
      
      return items;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetSiteDriveItemsByQuery(siteIdOrName: string, queryText: string): Promise<[MicrosoftGraph.DriveItem]> {
    try {
      const items = await this.client.api(`/sites/${siteIdOrName}/drive/root/search(q='${queryText}')`) 
        .get();
      
      return items;
    }
    catch (error) {
      throw error;
    }
  }
  
  public async GetDriveItems(driveId: string): Promise<[MicrosoftGraph.DriveItem]> {
    try {
      const items = await this.client.api(`/drives/${driveId}/root/children`)
        .get();
      
      return items;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetDriveItemsByQuery(driveId: string, queryText: string): Promise<[MicrosoftGraph.DriveItem]> {
    try {
      const items = await this.client.api(`/drives/${driveId}/root/search(q='${queryText}')`) 
        .get();
      
      return items;
    }
    catch (error) {
      throw error;
    }
  }  

  public async GetDriveFolderItems(driveId: string, folderId: string): Promise<[MicrosoftGraph.DriveItem]> {
    try {
      const items = await this.client.api(`/drives/${driveId}/items/${folderId}/children`) 
        .get();
      
      return items;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetDriveItem(driveId: string, itemId: string): Promise<MicrosoftGraph.DriveItem> {
    try {
      const items = await this.client.api(`/drives/${driveId}/items/${itemId}`) 
        .get();
      
      return items;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetTeamDefaultDriveItems(teamGroupId: string, relativePath: string): Promise<[MicrosoftGraph.DriveItem]> {
    try {
      var items = null;

      if (relativePath.length > 0 && relativePath != "/") {  
        if (!relativePath.startsWith("/")) {
          relativePath = `/${relativePath}`;
        }
        if (relativePath.endsWith("/")) {
          relativePath = relativePath.slice(0, -1);
        }
        items = await this.client.api(`/groups/${teamGroupId}/drive/root:${relativePath}:/children`)
        .get();
      } 
      else {
        items = await this.client.api(`/groups/${teamGroupId}/drive/root/children`)
        .get();
      }
      
      return items;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetTeamDriveItemsByQuery(teamGroupId: string, queryText: string): Promise<[MicrosoftGraph.DriveItem]> {
    try {
      const items = await this.client.api(`/groups/${teamGroupId}/drive/root/search(q='${queryText}')`) 
        .get();
      
      return items;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetTeam(teamId: string): Promise<MicrosoftGraph.Team> {
    try {
      const retTeam = await this.client.api(`/teams/${teamId}`)
        .get();
      return retTeam;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetTeamChannels(teamId: string): Promise<[MicrosoftGraph.Channel]> {
    try {
      const retChannels = await this.client.api(`/teams/${teamId}/channels`)
        .get();
      return retChannels;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetTeamChannel(teamId: string, channelId: string): Promise<MicrosoftGraph.Channel> {
    try {
      const retChannel = await this.client.api(`/teams/${teamId}/channels/${channelId}`)
        .get();
      return retChannel;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetTeamMembers(teamId: string): Promise<[MicrosoftGraph.DirectoryObject]> {
    try {
      const retMembers = await this.client.api(`/groups/${teamId}/members`)
        .get();
      return retMembers;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetTeamEvents(teamId: string): Promise<[MicrosoftGraph.Event]> {
    try {
      const retEvents = await this.client.api(`/groups/${teamId}/events`)
        .get();
      return retEvents;
    }
    catch (error) {
      throw error;
    }
  }

  public async GetUserByIdOrEmail(userIdOrEmail: string): Promise<[MicrosoftGraph.User]> {
    try {
      const retUser = await this.client.api(`/users/${userIdOrEmail}`)
        .get();
      return retUser;
    }
    catch (error) {
      throw error;
    }
  }

  // GetMyApplications: Permissions problems (output 403: Forbidden)
  public async GetMyApplications(): Promise<any> {
    try {
      // const retReport = await this.client.api("/reports/getOffice365ActivationsUserDetail(period='D7')")
      // const retReport = await this.client.api("/reports/getOffice365ActivationsUserDetail")
      const retReport = await this.client.api("/reports/getOffice365ActiveUserDetail(period='D7')")
        .get();
      return retReport;
    }
    catch (error) {
      throw error;
    }
  }






  // Not working (nb: beta)
  // public async GetUserPresence(userId: string): Promise<any> {
  //   try {
  //     const members = await this.client.api("/beta/users/" + userId + "/presence")
  //       .get();
  //     return members;
  //   }
  //   catch (error) {
  //     throw error;
  //   }
  // }

}




export = M365Wrapper;
