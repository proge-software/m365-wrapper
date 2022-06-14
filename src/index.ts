import "isomorphic-fetch";
import { PublicClientApplication, Configuration, PopupRequest } from '@azure/msal-browser';
import { Client } from '@microsoft/microsoft-graph-client';
import { AuthCodeMSALBrowserAuthenticationProvider, AuthCodeMSALBrowserAuthenticationProviderOptions } from "@microsoft/microsoft-graph-client/authProviders/authCodeMsalBrowser";
import PopupRequestConstants from "./constants/popup-request-constants";
import ConfigurationsConstants from "./constants/configurations-constants";
import UserHandler from "./handlers/user-handler";
import UsersHandler from "./handlers/users-handler";
import CalendarHandler from "./handlers/calendar-handler";
import TeamsHandler from "./handlers/teams-handler";
import OfficeHandler from "./handlers/office-handler";
import DriveHandler from "./handlers/drive-handler";
import SitesHandler from "./handlers/sites-handler";

class M365Wrapper {

  public office: OfficeHandler;
  public user: UserHandler;
  public users: UsersHandler;
  public calendar: CalendarHandler;
  public teams: TeamsHandler;
  public drive: DriveHandler;
  public sites: SitesHandler;

  constructor(clientId: string);
  constructor(clientId: string, authority?: string) {

    let configuration: Configuration = {
      auth: {
        clientId: clientId ? clientId : '',
        authority: authority? authority : ConfigurationsConstants.DEFAULT_AUTHORITY,
      },
      cache: {
        cacheLocation: ConfigurationsConstants.CACHE_LOCATION_SESSION_STORAGE
      }
    };

    let popupRequest: PopupRequest = {
      scopes: PopupRequestConstants.DEFAULT_SCOPES,
      prompt: PopupRequestConstants.PROMPT_SELECT_ACCOUNT
    };

    let msalApplication = new PublicClientApplication(configuration);
    let providerOptions = popupRequest as AuthCodeMSALBrowserAuthenticationProviderOptions;
    let authProvider = new AuthCodeMSALBrowserAuthenticationProvider(msalApplication, providerOptions);

    let client = Client.initWithMiddleware({
      authProvider: authProvider
    });

    this.office = new OfficeHandler(client);
    this.user = new UserHandler(msalApplication, client);
    this.users = new UsersHandler(client);
    this.calendar = new CalendarHandler(client);
    this.teams = new TeamsHandler(client);
    this.drive = new DriveHandler(client, this.office);
    this.sites = new SitesHandler(client);
  }

}

export = M365Wrapper;