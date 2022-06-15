import { PublicClientApplication, AuthenticationResult, AccountInfo, PopupRequest } from "@azure/msal-browser";
import { Client } from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import PopupRequestConstants from "../constants/popup-request-constants";

export default class UserHandler {

    private readonly popupRequest: PopupRequest = {
        scopes: PopupRequestConstants.DEFAULT_SCOPES,
        prompt: PopupRequestConstants.PROMPT_SELECT_ACCOUNT
    };

    constructor(private readonly msalApplication: PublicClientApplication, private readonly client: Client) { }

    public async loginPopup(): Promise<AuthenticationResult> {
        let loginResult = await this.msalApplication.loginPopup(this.popupRequest);
        this.msalApplication.setActiveAccount(loginResult.account);
        return loginResult;
    }

    public async acquireTokenSilent(): Promise<AuthenticationResult> {
        return await this.msalApplication.acquireTokenSilent(this.popupRequest);
    }

    public async acquireTokenPopup(): Promise<AuthenticationResult> {
        return await this.msalApplication.acquireTokenPopup(this.popupRequest);
    }

    public async statLoginPopupProcess() {
        let account = this.getMyAccountInfo();
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
                            const account = this.getMyAccountInfo();
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
                        account = this.getMyAccountInfo();
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

    public getMyAccountInfo(): AccountInfo {
        return this.msalApplication.getActiveAccount();
    }

    public logoutPopup() {
        this.msalApplication.logoutPopup(this.popupRequest);
    }

    public async getMyDetails(): Promise<MicrosoftGraph.User> {
        return (await this.client.api("/me").get()) as MicrosoftGraph.User;
    }

    // GetMyApplications: Permissions problems (output 403: Forbidden)
    public async getMyApplications(): Promise<any> {
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
}