import { Client } from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import M365App from "../models/results/m365-app";
import M365WrapperDataResult from "../models/results/m365-wrapper-data-result";
import ErrorsHandler from "./errors-handler";

export default class SitesHandler {

    constructor(private readonly client: Client) { }

    public async getRootSite(): Promise<M365WrapperDataResult<MicrosoftGraph.Site>> {
        try {
            let item: MicrosoftGraph.Site = await this.client.api(`/sites/root`)
                .get();

            return M365WrapperDataResult.createSuccess(item);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

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
    
    public async getApps(): Promise<M365WrapperDataResult<M365App[]>> {

        let rootSiteUrl: string = (await this.getRootSite()).data.webUrl;

        return new M365WrapperDataResult(null, [{
            name: 'SharePoint',
            link: rootSiteUrl,
            icon: ''
        }]);
    }
}