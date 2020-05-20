import "isomorphic-fetch";
import { UserAgentApplication, Configuration, AuthenticationParameters, AuthResponse, Account } from 'msal';
import { Client, ClientOptions, AuthenticationProvider } from '@microsoft/microsoft-graph-client';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { ImplicitMSALAuthenticationProvider, MSALAuthenticationProviderOptions } from "@microsoft/microsoft-graph-client/lib/src/browser";
//import { ImplicitMSALAuthenticationProvider as NodeImplicitMSALAuthenticationProvider } from "@microsoft/microsoft-graph-client/lib/src/ImplicitMSALAuthenticationProvider";

class M365Wrapper {
  protected authPar: AuthenticationParameters = {
    scopes: ['user.read', 'Calendars.ReadWrite'],
    prompt: 'select_account',
  };
  protected configuration: Configuration = {
    auth: {
      clientId: '9f43a6bd-9b42-4cf9-82f8-d9f1960596cc',
      authority: 'https://login.microsoftonline.com/common',
    },
    // cache: {
    //   cacheLocation: 'sessionStorage'
    // }
  };
  protected GraphScopes: string[] = [...this.authPar.scopes!];
  protected providerOptions: MSALAuthenticationProviderOptions;
  protected msalApplication: UserAgentApplication;
  protected authProvider: AuthenticationProvider;
  protected options: ClientOptions;
  protected client: Client;

  constructor();
  constructor(clientId?: string);
  constructor(clientId?: string, authority?: string) {
    if (clientId)
      this.configuration.auth.clientId = clientId;

    if (authority)
      this.configuration.auth.authority = authority;

    this.msalApplication = new UserAgentApplication(this.configuration);
    this.providerOptions = new MSALAuthenticationProviderOptions(this.GraphScopes);

    if (true)
      this.authProvider = new ImplicitMSALAuthenticationProvider(this.msalApplication, this.providerOptions);
    //else
    //  this.authProvider = new ImplicitMSALAuthenticationProvider(this.msalApplication, this.providerOptions);

    this.options = {
      authProvider: this.authProvider, // An instance created from previous step
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
        // await thatGraph.GetAvatar().then(a => {
        //     this.SET_AVATAR_IMAGE_URL(a);
        // })
      }).catch(async error => {
        if (error.errorMessage.indexOf("interaction_required") !== -1) {
          await this.acquireTokenPopup()
            .then(async response => {
              const account = this.getAccount();
              // this.SET_ACCOUNT(account);
              // this.SET_ID_TOKEN(response);
              // this.SET_LOGIN_STATE(true);
              // await thatGraph.GetAvatar().then(a => {
              //     this.SET_AVATAR_IMAGE_URL(a);
              // })
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
            // await thatGraph.GetAvatar().then(a => {
            //     this.SET_AVATAR_IMAGE_URL(a);
            // })
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

  public async GetUserDetail(): Promise<[MicrosoftGraph.User]> {
    try {
      let userDetails = await this.client.api("/me").get();
      let x: [MicrosoftGraph.User] = userDetails;
      return x;
      console.log(userDetails);
    } catch (error) {
      throw error;
    }

  }


  public TestStartup(): boolean {
    return true;
  };
}

// if(typeof exports != "undefined"){    
//   exports.M365Client = new M365Wrapper(); 
// }else{    
//   var M365Client =new M365Wrapper();
// }

export = M365Wrapper;

new M365Wrapper().TestStartup();