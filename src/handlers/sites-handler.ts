import { Client } from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import M365WrapperDataResult from "../models/results/m365-wrapper-data-result";
import ErrorsHandler from "./errors-handler";

export default class SitesHandler {

    constructor(private readonly client: Client) { }

    public async getSiteDrives(siteIdOrName: string): Promise<M365WrapperDataResult<[MicrosoftGraph.Drive]>> {
        try {
            let items: [MicrosoftGraph.Drive] = await this.client.api(`/sites/${siteIdOrName}/drives`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getSiteDriveItemsByQuery(siteIdOrName: string, queryText: string): Promise<M365WrapperDataResult<[MicrosoftGraph.DriveItem]>> {
        try {
            let items: [MicrosoftGraph.DriveItem] = await this.client.api(`/sites/${siteIdOrName}/drive/root/search(q='${queryText}')`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }
}