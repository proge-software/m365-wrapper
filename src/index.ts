import "isomorphic-fetch";
import { PublicClientApplication, Configuration, AuthenticationResult, PopupRequest, AccountInfo } from '@azure/msal-browser';
import { Client, ClientOptions, AuthenticationProvider } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { AuthCodeMSALBrowserAuthenticationProvider, AuthCodeMSALBrowserAuthenticationProviderOptions } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import PopupRequestConstants from "./constants/popup-request-constants";
import UserHandler from "./handlers/user-handler";
import UsersHandler from "./handlers/users-handler";
import CalendarHandler from "./handlers/calendar-handler";
import TeamsHandler from "./handlers/teams-handler";
import OfficeHandler from "./handlers/office-handler";
import DriveHandler from "./handlers/drive-handler";

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

  public office: OfficeHandler;
  public user: UserHandler;
  public users: UsersHandler;
  public calendar: CalendarHandler;
  public teams: TeamsHandler;
  public drive: DriveHandler;

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

    this.office = new OfficeHandler(this.client);
    this.user = new UserHandler(this.msalApplication, this.client);
    this.users = new UsersHandler(this.client);
    this.calendar = new CalendarHandler(this.client);
    this.teams = new TeamsHandler(this.client);
    this.drive = new DriveHandler(this.client, this.office);
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