import "isomorphic-fetch";
import { PublicClientApplication, Configuration, AuthenticationResult, PopupRequest, AccountInfo } from '@azure/msal-browser';
import { Client, ClientOptions, AuthenticationProvider } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { AuthCodeMSALBrowserAuthenticationProvider, AuthCodeMSALBrowserAuthenticationProviderOptions } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import UserSearchRequest from "./models/requests/user-search-request";
import PopupRequestConstants from "./constants/popup-request-constants";
import UserHandler from "./handlers/user-handler";

class M365Wrapper {

  private readonly authPar: PopupRequest = {
    scopes: PopupRequestConstants.DEFAULT_SCOPES,
    prompt: PopupRequestConstants.PROMPT_SELECT_ACCOUNT
  };

  private readonly configuration: Configuration = {
    auth: {
      clientId: '',
      authority: 'https://login.microsoftonline.com/organizations',
    },
    cache: {
      cacheLocation: 'sessionStorage'
    }
  };

  protected GraphScopes: string[] = [...this.authPar.scopes!];
  protected providerOptions: AuthCodeMSALBrowserAuthenticationProviderOptions;
  protected msalApplication: PublicClientApplication;
  protected authProvider: AuthenticationProvider;
  protected options: ClientOptions;
  protected client: Client;

  public user: UserHandler;

  constructor(clientId: string);
  constructor(clientId: string, authority?: string) {
    if (clientId)
      this.configuration.auth.clientId = clientId;

    if (authority)
      this.configuration.auth.authority = authority;

    this.msalApplication = new PublicClientApplication(this.configuration);
    this.providerOptions = this.authPar as AuthCodeMSALBrowserAuthenticationProviderOptions;
    this.authProvider = new AuthCodeMSALBrowserAuthenticationProvider(this.msalApplication, this.providerOptions);

    this.options = {
      authProvider: this.authProvider
    };

    this.client = Client.initWithMiddleware(this.options);

    this.user = new UserHandler(this.msalApplication, this.client);
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

  public async IsTeamsInMyLicenses(): Promise<boolean> {
    try {

      var bFound = false;
      var teamsSkuPartNumbers: string[] = ['ENTERPRISEPACK_FACULTY',
        'STANDARDWOFFPACK_FACULTY',
        'STANDARDWOFFPACK_IW_FACULTY',
        'ENTERPRISEPREMIUM_FACULTY',
        'ENTERPRISEPREMIUM_NOPSTNCONF_FACULTY',
        'STANDARDPACK_FACULTY',
        'ENTERPRISEPACK_EDULRG',
        'ENTERPRISEWITHSCAL_FACULTY',
        'M365EDU_A3_FACULTY',
        'M365EDU_A5_FACULTY',
        'M365EDU_A5_NOPSTNCONF_FACULTY',
        'STANDARDWOFFPACK_HOMESCHOOL_FAC',
        'STANDARDWOFFPACK_FACULTY_DEVICE',
        'ENTERPRISEPACK_STUDENT',
        'STANDARDWOFFPACK_IW_STUDENT',
        'ENTERPRISEPREMIUM_STUDENT',
        'ENTERPRISEPREMIUM_NOPSTNCONF_STUDENT',
        'STANDARDPACK_STUDENT',
        'ENTERPRISEWITHSCAL_STUDENT',
        'M365EDU_A3_STUDENT',
        'M365EDU_A3_STUUSEBNFT',
        'M365EDU_A5_STUDENT',
        'M365EDU_A5_STUUSEBNFT',
        'M365EDU_A5_NOPSTNCONF_STUDENT',
        'M365EDU_A5_NOPSTNCONF_STUUSEBNFT',
        'ENTERPRISEPACKPLUS_STUDENT',
        'ENTERPRISEPACKPLUS_STUUSEBNFT',
        'ENTERPRISEPREMIUM_STUUSEBNFT',
        'ENTERPRISEPREMIUM_NOPSTNCONF_STUUSEBNFT',
        'STANDARDWOFFPACK_HOMESCHOOL_STU',
        'STANDARDWOFFPACK_STUDENT_DEVICE',
        'STANDARDWOFFPACK_IW_STUDENT']

      var licenses;
      try {
        licenses = await this.client.api(`/me/licenseDetails`)
          .get();
      } catch (error) {
        return false;
      }

      for (var i = 0; i < licenses.value.length; i++) {
        if (teamsSkuPartNumbers.includes(licenses.value[i].skuPartNumber)) {
          bFound = true;
          break;
        }
      }

      return bFound;
    }
    catch (error) {
      throw error;
    }
  }

  public async IsOneDriveInMyLicenses(): Promise<boolean> {
    try {

      var bFound = false;
      var teamsSkuPartNumbers: string[] = ['O365_BUSINESS',
        'SMB_BUSINESS',
        'OFFICESUBSCRIPTION',
        'WACONEDRIVESTANDARD',
        'WACONEDRIVEENTERPRISE',
        'VISIOONLINE_PLAN1',
        'VISIOCLIENT']

      var licenses;
      try {
        licenses = await this.client.api(`/me/licenseDetails`)
          .get();
      } catch (error) {
        return false;
      }

      for (var i = 0; i < licenses.value.length; i++) {
        if (teamsSkuPartNumbers.includes(licenses.value[i].skuPartNumber)) {
          bFound = true;
          break;
        }
      }

      if (!bFound) {
        bFound = await this.IsOfficeInMyLicenses();
      }

      return bFound;
    }
    catch (error) {
      throw error;
    }
  }

  public async IsOfficeInMyLicenses(): Promise<boolean> {
    try {

      var bFound = false;
      var teamsSkuPartNumbers: string[] = ['M365EDU_A3_FACULTY',
        'M365EDU_A3_STUDENT',
        'M365EDU_A5_FACULTY',
        'M365EDU_A5_STUDENT',
        'O365_BUSINESS',
        'SMB_BUSINESS',
        'OFFICESUBSCRIPTION',
        'O365_BUSINESS_ESSENTIALS', // Mobile
        'SMB_BUSINESS_ESSENTIALS', // Mobile
        'O365_BUSINESS_PREMIUM',
        'SMB_BUSINESS_PREMIUM',
        'SPB',
        'SPE_E3',
        'SPE_E5',
        'SPE_E3_USGOV_DOD',
        'SPE_E3_USGOV_GCCHIGH',
        'SPE_F1', // Mobile
        'ENTERPRISEPREMIUM_FACULTY',
        'ENTERPRISEPREMIUM_STUDENT',
        'STANDARDPACK', // Mobile
        'ENTERPRISEPACK',
        'DEVELOPERPACK',
        'ENTERPRISEPACK_USGOV_DOD',
        'ENTERPRISEPACK_USGOV_GCCHIGH',
        'ENTERPRISEWITHSCAL',
        'ENTERPRISEPREMIUM',
        'ENTERPRISEPREMIUM_NOPSTNCONF',
        'DESKLESSPACK', // Mobile
        'MIDSIZEPACK',
        'LITEPACK_P2']

      var licenses;
      try {
        licenses = await this.client.api(`/me/licenseDetails`)
          .get();
      } catch (error) {
        return false;
      }

      for (var i = 0; i < licenses.value.length; i++) {
        if (teamsSkuPartNumbers.includes(licenses.value[i].skuPartNumber)) {
          bFound = true;
          break;
        }
      }

      return bFound;
    }
    catch (error) {
      throw error;
    }
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

  public async GetMyDriveItemSharingPermissions(itemId: string): Promise<[MicrosoftGraph.Permission]> {
    try {
      const items = await this.client.api(`/me/drive/items/${itemId}/permissions`)
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