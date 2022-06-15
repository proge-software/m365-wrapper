import { Client } from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export default class TeamsHandler {

    constructor(private readonly client: Client) { }

    public async IsInMyLicenses(): Promise<boolean> {
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
}