import { Client } from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import M365App from "../models/results/m365-app";
import M365WrapperDataResult from "../models/results/m365-wrapper-data-result";
import M365WrapperResult from "../models/results/m365-wrapper-result";
import ErrorsHandler from "./errors-handler";

export default class TeamsHandler {

    constructor(private readonly client: Client) { }

    public async isInMyLicenses(): Promise<M365WrapperResult> {
        try {

            let result: M365WrapperResult = { isSuccess: false } as M365WrapperResult;
            let teamsSkuPartNumbers: string[] = ['ENTERPRISEPACK_FACULTY',
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

            let licenses = await this.client.api(`/me/licenseDetails`)
                .get();

            for (let i = 0; i < licenses.value.length; i++) {
                if (teamsSkuPartNumbers.includes(licenses.value[i].skuPartNumber)) {
                    result.isSuccess = true;
                    break;
                }
            }

            return result;
        }
        catch (error) {
            return ErrorsHandler.getErrorResult(error);
        }
    }

    public async getMyJoinedTeams(): Promise<M365WrapperDataResult<[MicrosoftGraph.Team]>> {
        try {
            let teams: [MicrosoftGraph.Team] = await this.client.api("/me/joinedTeams")
                .select('Id,displayName,description')
                .get();

            return M365WrapperDataResult.createSuccess(teams);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async createOnlineMeeting(onlineMeeting: MicrosoftGraph.OnlineMeeting): Promise<M365WrapperDataResult<[MicrosoftGraph.OnlineMeeting]>> {

        try {
            let res: [MicrosoftGraph.OnlineMeeting] = await this.client.api('/me/onlineMeetings')
                .post(onlineMeeting);

            return M365WrapperDataResult.createSuccess(res);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getTeam(teamId: string): Promise<M365WrapperDataResult<MicrosoftGraph.Team>> {
        try {
            let retTeam: MicrosoftGraph.Team = await this.client.api(`/teams/${teamId}`)
                .get();
            return M365WrapperDataResult.createSuccess(retTeam);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getTeamChannels(teamId: string): Promise<M365WrapperDataResult<[MicrosoftGraph.Channel]>> {
        try {
            let retChannels: [MicrosoftGraph.Channel] = await this.client.api(`/teams/${teamId}/channels`)
                .get();
            return M365WrapperDataResult.createSuccess(retChannels);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getTeamChannel(teamId: string, channelId: string): Promise<M365WrapperDataResult<MicrosoftGraph.Channel>> {
        try {
            let retChannel: MicrosoftGraph.Channel = await this.client.api(`/teams/${teamId}/channels/${channelId}`)
                .get();
            return M365WrapperDataResult.createSuccess(retChannel);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getTeamMembers(teamId: string): Promise<M365WrapperDataResult<[MicrosoftGraph.DirectoryObject]>> {
        try {
            let retMembers: [MicrosoftGraph.DirectoryObject] = await this.client.api(`/groups/${teamId}/members`)
                .get();
            return M365WrapperDataResult.createSuccess(retMembers);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getTeamEvents(teamId: string): Promise<M365WrapperDataResult<[MicrosoftGraph.Event]>> {
        try {
            let retEvents: [MicrosoftGraph.Event] = await this.client.api(`/groups/${teamId}/events`)
                .get();
            return M365WrapperDataResult.createSuccess(retEvents);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getTeamDrives(teamGroupId: string): Promise<M365WrapperDataResult<[MicrosoftGraph.Drive]>> {
        try {
            let items: [MicrosoftGraph.Drive] = await this.client.api(`/groups/${teamGroupId}/drives`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getTeamDefaultDriveItems(teamGroupId: string, relativePath: string): Promise<M365WrapperDataResult<[MicrosoftGraph.DriveItem]>> {
        try {
            let items: [MicrosoftGraph.DriveItem] = null;

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

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getTeamDriveItemsByQuery(teamGroupId: string, queryText: string): Promise<M365WrapperDataResult<[MicrosoftGraph.DriveItem]>> {
        try {
            let items: [MicrosoftGraph.DriveItem] = await this.client.api(`/groups/${teamGroupId}/drive/root/search(q='${queryText}')`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }
    
    public getApps(): M365WrapperDataResult<M365App[]>{
        return new M365WrapperDataResult(null, [{
            name: 'Teams',
            link: 'https://teams.microsoft.com',
            icon: 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCI+PGRlZnM+PHN0eWxlPi5jbHMtMXtvcGFjaXR5OjAuMTt9LmNscy0ye2ZpbGw6bm9uZTt9LmNscy0ze2ZpbGw6IzUwNTljOTt9LmNscy00e2ZpbGw6IzdiODNlYjt9LmNscy01e29wYWNpdHk6MC41O30uY2xzLTZ7ZmlsbDojNGI1M2JjO30uY2xzLTd7ZmlsbDojZmZmO308L3N0eWxlPjwvZGVmcz48dGl0bGU+VGVhbXNfMjR4PC90aXRsZT48ZyBpZD0iVGVhbXMiPjxwYXRoIGNsYXNzPSJjbHMtMSIgZD0iTTgsMTAuMTJWMTcuNUE1LjQsNS40LDAsMCwwLDguNjEsMjBoNUExLjUsMS41LDAsMCwwLDE1LDE4LjY1VjlsLS4yOCwwSDkuMTJBMS4xMiwxLjEyLDAsMCwwLDgsMTAuMTJaIi8+PHBhdGggY2xhc3M9ImNscy0xIiBkPSJNMTMuODMsNmgtM0EzLjI5LDMuMjksMCwwLDAsMTQsOC4zMWEzLjMzLDMuMzMsMCwwLDAsMS0uMTd2LTFBMS4xOCwxLjE4LDAsMCwwLDEzLjgzLDZaIi8+PHJlY3QgY2xhc3M9ImNscy0yIiB3aWR0aD0iMjQiIGhlaWdodD0iMjQiLz48cGF0aCBjbGFzcz0iY2xzLTMiIGQ9Ik0yMi44Nyw5aC01bC0xLjM5LDEuMTN2NS41OWEzLjc2LDMuNzYsMCwwLDAsNy41MSwwVjEwLjEzQTEuMTIsMS4xMiwwLDAsMCwyMi44Nyw5WiIvPjxjaXJjbGUgaWQ9IkhlYWQiIGNsYXNzPSJjbHMtMyIgY3g9IjIwLjUiIGN5PSI1LjUiIHI9IjIuNSIvPjxwYXRoIGNsYXNzPSJjbHMtNCIgZD0iTTkuMTIsOWg4Ljc2QTEuMTIsMS4xMiwwLDAsMSwxOSwxMC4xMlYxNy41QTUuNSw1LjUsMCwwLDEsMTMuNSwyM2gwQTUuNSw1LjUsMCwwLDEsOCwxNy41VjEwLjEyQTEuMTIsMS4xMiwwLDAsMSw5LjEyLDlaIi8+PGNpcmNsZSBpZD0iSGVhZC0yIiBkYXRhLW5hbWU9IkhlYWQiIGNsYXNzPSJjbHMtNCIgY3g9IjE0IiBjeT0iNSIgcj0iMy4zMSIvPjxwYXRoIGNsYXNzPSJjbHMtNSIgZD0iTTgsMTAuMTJWMTcuNUE1LjQsNS40LDAsMCwwLDguNjEsMjBoNUExLjUsMS41LDAsMCwwLDE1LDE4LjY1VjlsLS4yOCwwSDkuMTJBMS4xMiwxLjEyLDAsMCwwLDgsMTAuMTJaIi8+PHBhdGggY2xhc3M9ImNscy01IiBkPSJNMTMuODMsNmgtM0EzLjI5LDMuMjksMCwwLDAsMTQsOC4zMWEzLjMzLDMuMzMsMCwwLDAsMS0uMTd2LTFBMS4xOCwxLjE4LDAsMCwwLDEzLjgzLDZaIi8+PHJlY3QgaWQ9IkJhY2tfUGxhdGUiIGRhdGEtbmFtZT0iQmFjayBQbGF0ZSIgY2xhc3M9ImNscy02IiB5PSI1IiB3aWR0aD0iMTQiIGhlaWdodD0iMTQiIHJ4PSIxLjE3Ii8+PHBhdGggY2xhc3M9ImNscy03IiBkPSJNMTAuMTgsOS41OEg3Ljc5VjE2SDYuMjJWOS41OEgzLjgyVjhoNi4zNloiLz48L2c+PC9zdmc+'
        }]);
    }
}