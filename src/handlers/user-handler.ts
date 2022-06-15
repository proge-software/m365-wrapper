import { PublicClientApplication, AuthenticationResult, AccountInfo, PopupRequest } from "@azure/msal-browser";
import { Client } from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import PopupRequestConstants from "../constants/popup-request-constants";
import M365WrapperDataResult from "../models/results/m365-wrapper-data-result";
import M365WrapperResult from "../models/results/m365-wrapper-result";
import ErrorsHandler from "./errors-handler";

export default class UserHandler {

    private readonly popupRequest: PopupRequest = {
        scopes: PopupRequestConstants.DEFAULT_SCOPES,
        prompt: PopupRequestConstants.PROMPT_SELECT_ACCOUNT
    };

    constructor(private readonly msalApplication: PublicClientApplication, private readonly client: Client) { }

    public async loginPopup(): Promise<M365WrapperDataResult<AuthenticationResult>> {

        try {
            let loginResult = await this.msalApplication.loginPopup(this.popupRequest);
            this.msalApplication.setActiveAccount(loginResult.account);
            return M365WrapperDataResult.createSuccess(loginResult);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async acquireTokenSilent(): Promise<M365WrapperDataResult<AuthenticationResult>> {

        try {
            return M365WrapperDataResult.createSuccess(await this.msalApplication.acquireTokenSilent(this.popupRequest));
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async acquireTokenPopup(): Promise<M365WrapperDataResult<AuthenticationResult>> {

        try {
            return M365WrapperDataResult.createSuccess(await this.msalApplication.acquireTokenPopup(this.popupRequest));
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async statLoginPopupProcess() {

        try {
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
                                let account = this.getMyAccountInfo();
                                // this.SET_ACCOUNT(account);
                                // this.SET_ID_TOKEN(response);
                                // this.SET_LOGIN_STATE(true);              
                            })
                            .catch(err => {
                                return ErrorsHandler.getErrorDataResult(err);
                            });
                    }
                    else
                        return ErrorsHandler.getErrorDataResult(error);
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
                                return ErrorsHandler.getErrorDataResult(err);
                            });
                    })
                    .catch(err => {
                        return ErrorsHandler.getErrorDataResult(err);
                    });

            }
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public getMyAccountInfo(): M365WrapperDataResult<AccountInfo> {
        try {
            return M365WrapperDataResult.createSuccess(this.msalApplication.getActiveAccount());
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async logoutPopup(): Promise<M365WrapperResult> {
        try {
            await this.msalApplication.logoutPopup(this.popupRequest);
            return new M365WrapperResult();
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getMyDetails(): Promise<M365WrapperDataResult<MicrosoftGraph.User>> {
        try {
            let result: MicrosoftGraph.User = await this.client.api("/me").get();
            return M365WrapperDataResult.createSuccess(result);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    // GetMyApplications: Permissions problems (output 403: Forbidden)
    public async getMyApplications(): Promise<M365WrapperDataResult<any>> {
        try {
            // const retReport = await this.client.api("/reports/getOffice365ActivationsUserDetail(period='D7')")
            // const retReport = await this.client.api("/reports/getOffice365ActivationsUserDetail")
            let retReport = await this.client.api("/reports/getOffice365ActiveUserDetail(period='D7')")
                .get();
            return M365WrapperDataResult.createSuccess(retReport);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }
}