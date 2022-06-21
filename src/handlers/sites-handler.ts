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
            icon: 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCI+PGRlZnM+PHN0eWxlPi5jbHMtMXtmaWxsOm5vbmU7fS5jbHMtMntmaWxsOiMwMzZjNzA7fS5jbHMtM3tmaWxsOiMxYTliYTE7fS5jbHMtNHtmaWxsOiMzN2M2ZDA7fS5jbHMtNXtvcGFjaXR5OjAuNTt9LmNscy02e2ZpbGw6IzAzODM4Nzt9LmNscy03e2ZpbGw6I2ZmZjt9PC9zdHlsZT48L2RlZnM+PHRpdGxlPlNoYXJlcG9pbnRfMjR4PC90aXRsZT48ZyBpZD0iU2hhcmVwb2ludCI+PGcgaWQ9Il8yNCIgZGF0YS1uYW1lPSIyNCI+PHJlY3QgY2xhc3M9ImNscy0xIiB3aWR0aD0iMjQiIGhlaWdodD0iMjQiLz48Y2lyY2xlIGNsYXNzPSJjbHMtMiIgY3g9IjExIiBjeT0iNyIgcj0iNyIvPjxjaXJjbGUgY2xhc3M9ImNscy0zIiBjeD0iMTgiIGN5PSIxMyIgcj0iNiIvPjxjaXJjbGUgY2xhc3M9ImNscy00IiBjeD0iMTIiIGN5PSIxOSIgcj0iNSIvPjxwYXRoIGNsYXNzPSJjbHMtNSIgZD0iTTEzLjgzLDZINC4wN0E2LjYzLDYuNjMsMCwwLDAsNCw3YTcsNywwLDAsMCw3LDcsNy41OSw3LjU5LDAsMCwwLDEuMDctLjA4czAsLjA1LDAsLjA4SDEyYTUsNSwwLDAsMC0xLjU1LjI1bC40MiwwdjBsLS42NC4wNkE1LDUsMCwwLDAsNywxOWE0LjcxLDQuNzEsMCwwLDAsLjEsMWg2LjVBMS41LDEuNSwwLDAsMCwxNSwxOC42NVY3LjE3QTEuMTgsMS4xOCwwLDAsMCwxMy44Myw2WiIvPjxyZWN0IGlkPSJCYWNrX1BsYXRlIiBkYXRhLW5hbWU9IkJhY2sgUGxhdGUiIGNsYXNzPSJjbHMtNiIgeT0iNSIgd2lkdGg9IjE0IiBoZWlnaHQ9IjE0IiByeD0iMS4xNyIvPjxwYXRoIGNsYXNzPSJjbHMtNyIgZD0iTTUsMTEuODVhMi4zNywyLjM3LDAsMCwxLS43MS0uNzQsMi4wOCwyLjA4LDAsMCwxLS4yNC0xLDIsMiwwLDAsMSwuNDUtMS4zMkEyLjY5LDIuNjksMCwwLDEsNS43Miw4YTUuMiw1LjIsMCwwLDEsMS42Ni0uMjZBNi4xMSw2LjExLDAsMCwxLDkuNTYsOFY5LjU3YTMuNDgsMy40OCwwLDAsMC0xLS40MUE1LDUsMCwwLDAsNy40Miw5YTIuNDMsMi40MywwLDAsMC0xLjE4LjI2Ljc3Ljc3LDAsMCwwLS40OC43MS43Mi43MiwwLDAsMCwuMi41LDEuODgsMS44OCwwLDAsMCwuNTQuMzljLjIzLjExLjU2LjI2LDEsLjQ0bC4xNC4wNkE3Ljg3LDcuODcsMCwwLDEsOC45MiwxMmEyLjI0LDIuMjQsMCwwLDEsLjc1Ljc1LDIuMTksMi4xOSwwLDAsMSwuMjcsMS4xNCwyLjE4LDIuMTgsMCwwLDEtLjQyLDEuMzhBMi4zOCwyLjM4LDAsMCwxLDguMzcsMTZhNS4wNiw1LjA2LDAsMCwxLTEuNjIuMjQsOC43LDguNywwLDAsMS0xLjQ4LS4xMiw1LjE3LDUuMTcsMCwwLDEtMS4yLS4zNVYxNC4xOGEzLjY0LDMuNjQsMCwwLDAsMS4yMi41OEE0Ljc5LDQuNzksMCwwLDAsNi42MiwxNWEyLjMxLDIuMzEsMCwwLDAsMS4yMS0uMjZBLjgxLjgxLDAsMCwwLDguMjQsMTQsLjc4Ljc4LDAsMCwwLDgsMTMuNDQsMi4zMSwyLjMxLDAsMCwwLDcuMzgsMTNjLS4yNy0uMTMtLjY2LS4zMS0xLjE5LS41M0E2LjM3LDYuMzcsMCwwLDEsNSwxMS44NVoiLz48L2c+PC9nPjwvc3ZnPg=='
        }]);
    }
}