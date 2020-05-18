import "isomorphic-fetch";
import { UserAgentApplication, Configuration, AuthenticationParameters, AuthResponse, Account } from 'msal';
import { Client, ClientOptions } from '@microsoft/microsoft-graph-client';
//import { } from '@microsoft/microsoft-graph-types';
import { ImplicitMSALAuthenticationProvider, MSALAuthenticationProviderOptions } from "@microsoft/microsoft-graph-client/lib/src/browser";

export class M365Wrapper {
  protected authPar: AuthenticationParameters = {
    scopes: ['user.read', 'Calendars.ReadWrite'],
    prompt: 'select_account',
  };
  protected configuration: Configuration = {
    auth: {
      clientId: '271d15ae-ec0e-42c9-bfaa-fd0b325e96d2',
      authority: 'https://login.microsoftonline.com/common',
    },
    cache: {
      cacheLocation: 'sessionStorage'
    }
  };
  protected GraphScopes: string[] = [...this.authPar.scopes!];
  protected providerOptions = new MSALAuthenticationProviderOptions(this.GraphScopes);
  protected userAgentApplication: UserAgentApplication = new UserAgentApplication(this.configuration);
  protected authProvider = new ImplicitMSALAuthenticationProvider(this.userAgentApplication, this.providerOptions);
  protected options: ClientOptions = {
    authProvider: this.authProvider, // An instance created from previous step
  };
  protected client: Client = Client.initWithMiddleware(this.options);
  constructor() {

  }

  public TestStartup(): boolean {
    return true;
  };

  public static start() {
    let x = new M365Wrapper();
    
  }

}

M365Wrapper.start();