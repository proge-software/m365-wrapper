import { Client } from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';

export default class SitesHandler {

    constructor(private readonly client: Client) { }

    public async getSiteDrives(siteIdOrName: string): Promise<[MicrosoftGraph.Drive]> {
        try {
            const items = await this.client.api(`/sites/${siteIdOrName}/drives`)
                .get();

            return items;
        }
        catch (error) {
            throw error;
        }
    }

    public async getSiteDriveItemsByQuery(siteIdOrName: string, queryText: string): Promise<[MicrosoftGraph.DriveItem]> {
        try {
            const items = await this.client.api(`/sites/${siteIdOrName}/drive/root/search(q='${queryText}')`)
                .get();

            return items;
        }
        catch (error) {
            throw error;
        }
    }
}