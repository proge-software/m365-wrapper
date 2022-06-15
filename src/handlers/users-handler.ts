import { Client } from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import UserSearchRequest from "../models/requests/user-search-request";

export default class UsersHandler {

    constructor(private readonly client: Client) { }

    public async getUsers(UserSearchRequest: UserSearchRequest): Promise<MicrosoftGraph.User[]> {
        let query = this.client.api('/users');

        if (UserSearchRequest && UserSearchRequest.issuer && UserSearchRequest.mail) {
            query = query.filter(`identities/any(c:c/issuerAssignedId eq '${UserSearchRequest.mail}' and c/issuer eq '${UserSearchRequest.issuer}')`);
        }

        let res: MicrosoftGraph.User[] = await query.select('displayName,givenName,postalCode,mail,surname,userPrincipalName')
            .get();

        return res;
    }

    public async getUserByIdOrEmail(userIdOrEmail: string): Promise<[MicrosoftGraph.User]> {
        try {
            const retUser = await this.client.api(`/users/${userIdOrEmail}`)
                .get();
            return retUser;
        }
        catch (error) {
            throw error;
        }
    }

    // Not working (nb: beta)
    // public async getUserPresence(userId: string): Promise<any> {
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