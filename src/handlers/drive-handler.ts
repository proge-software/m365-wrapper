import { Client } from "@microsoft/microsoft-graph-client";
import { Drive, DriveItem, Permission } from '@microsoft/microsoft-graph-types';
import M365App from "../models/results/m365-app";
import M365WrapperDataResult from "../models/results/m365-wrapper-data-result";
import M365WrapperResult from "../models/results/m365-wrapper-result";
import ErrorsHandler from "./errors-handler";
import OfficeHandler from "./office-handler";

export default class DriveHandler {

    constructor(private readonly client: Client, private readonly office: OfficeHandler) { }

    public async isOneDriveInMyLicenses(): Promise<M365WrapperResult> {
        try {

            let result: M365WrapperResult = { isSuccess: false } as M365WrapperResult;
            let teamsSkuPartNumbers: string[] = ['O365_BUSINESS',
                'SMB_BUSINESS',
                'OFFICESUBSCRIPTION',
                'WACONEDRIVESTANDARD',
                'WACONEDRIVEENTERPRISE',
                'VISIOONLINE_PLAN1',
                'VISIOCLIENT']

            let licenses = await this.client.api(`/me/licenseDetails`)
                .get();

            for (let i = 0; i < licenses.value.length; i++) {
                if (teamsSkuPartNumbers.includes(licenses.value[i].skuPartNumber)) {
                    result.isSuccess = true;
                    break;
                }
            }

            if (!result) {
                result.isSuccess = (await this.office.isInMyLicenses()).isSuccess;
            }

            return result;
        }
        catch (error) {
            return ErrorsHandler.getErrorResult(error);
        }
    }

    public async getMyDrives(): Promise<M365WrapperDataResult<[Drive]>> {
        try {
            let items: [Drive] = await this.client.api("/me/drives")
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getMyDriveItemsByQuery(queryText: string): Promise<M365WrapperDataResult<[DriveItem]>> {
        try {
            let items = await this.client.api(`/me/drive/root/search(q='${queryText}')`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getMyDriveAndSharedItemsByQuery(queryText: string): Promise<M365WrapperDataResult<[DriveItem]>> {
        try {
            let items = await this.client.api(`/me/drive/search(q='${queryText}')`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getMySharedItems(): Promise<M365WrapperDataResult<[DriveItem]>> {
        try {
            let items = await this.client.api(`/me/drive/sharedWithMe`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getDriveItems(driveId: string): Promise<M365WrapperDataResult<[DriveItem]>> {
        try {
            let items = await this.client.api(`/drives/${driveId}/root/children`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getDriveItemsByQuery(driveId: string, queryText: string): Promise<M365WrapperDataResult<[DriveItem]>> {
        try {
            let items = await this.client.api(`/drives/${driveId}/root/search(q='${queryText}')`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getDriveFolderItems(driveId: string, folderId: string): Promise<M365WrapperDataResult<DriveItem[]>> {
        try {
            let items = await this.client.api(`/drives/${driveId}/items/${folderId}/children`)
                .get();

            return M365WrapperDataResult.createSuccess(items.value);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getDriveItem(driveId: string, itemId: string): Promise<M365WrapperDataResult<DriveItem>> {
        try {
            let items = await this.client.api(`/drives/${driveId}/items/${itemId}`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getMyDriveItemSharingPermissions(itemId: string): Promise<M365WrapperDataResult<[Permission]>> {
        try {
            let items = await this.client.api(`/me/drive/items/${itemId}/permissions`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public getApps(): M365WrapperDataResult<M365App[]> {
        return new M365WrapperDataResult(null, [{
            name: 'OneDrive',
            link: 'https://onedrive.live.com',
            icon: 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCI+PGRlZnM+PHN0eWxlPi5jbHMtMXtmaWxsOm5vbmU7fS5jbHMtMntmaWxsOiMwMzY0Yjg7fS5jbHMtM3tmaWxsOiMwMDc4ZDQ7fS5jbHMtNHtmaWxsOiMxNDkwZGY7fS5jbHMtNXtmaWxsOiMyOGE4ZWE7fTwvc3R5bGU+PC9kZWZzPjx0aXRsZT5PbmVEcml2ZV8yNHg8L3RpdGxlPjxnIGlkPSJPbmVEcml2ZSI+PHJlY3QgY2xhc3M9ImNscy0xIiB3aWR0aD0iMjQiIGhlaWdodD0iMjQiLz48cGF0aCBjbGFzcz0iY2xzLTIiIGQ9Ik0xNC41LDE1bDQuOTUtNC43NEE3LjUsNy41LDAsMCwwLDUuOTIsOEM2LDgsMTQuNSwxNSwxNC41LDE1WiIvPjxwYXRoIGNsYXNzPSJjbHMtMyIgZD0iTTkuMTUsOC44OWgwQTYsNiwwLDAsMCw2LDhINS45MmE2LDYsMCwwLDAtNC44NCw5LjQzTDguNSwxNi41bDUuNjktNC41OVoiLz48cGF0aCBjbGFzcz0iY2xzLTQiIGQ9Ik0xOS40NSwxMC4yNmgtLjMyYTQuODQsNC44NCwwLDAsMC0xLjk0LjRoMGwtMywxLjI2TDE3LjUsMTZsNS45MiwxLjQ0YTQuODgsNC44OCwwLDAsMC00LTcuMThaIi8+PHBhdGggY2xhc3M9ImNscy01IiBkPSJNMS4wOCwxNy40M0E2LDYsMCwwLDAsNiwyMEgxOS4xM2E0Ljg5LDQuODksMCwwLDAsNC4yOS0yLjU2bC05LjIzLTUuNTNaIi8+PC9nPjwvc3ZnPg=='
        }]);
    }

    public async createFolder(driveId: string, parentItemId: string, folder: string | DriveItem): Promise<M365WrapperDataResult<DriveItem>> {
        try {

            let driveItem: DriveItem;

            if (typeof folder === 'string') {
                driveItem = {
                    name: folder,
                    folder: {},
                    '@microsoft.graph.conflictBehavior': 'rename'
                } as DriveItem;
            }
            else {
                driveItem = folder;
            }

            let result: DriveItem = await this.client.api(`/drives/${driveId}/items/${parentItemId}/children`)
                .post(driveItem);

            return M365WrapperDataResult.createSuccess(result);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async uploadSmallFile(driveId: string, parentItemId: string, filename: string, stream: any): Promise<M365WrapperDataResult<DriveItem>> {
        try {
            let folderItems = await this.getDriveFolderItems(driveId, parentItemId);
            let folderItem = folderItems.data.find(x => x.name == filename);

            let result: DriveItem;
            if(folderItem != undefined) {
                result = await this.client.api(`/drives/${driveId}/items/${folderItem.id}/content`)
                .put(stream);
            }
            else {
                result = await this.client.api(`/drives/${driveId}/items/${parentItemId}:/${filename}:/content`)
                .put(stream);
            }

            return M365WrapperDataResult.createSuccess(result);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }
    
    public async uploadLargeFile(driveId: string, parentItemId: string, filename: string, stream: any): Promise<M365WrapperDataResult<DriveItem>> {
        try {

            let result: DriveItem = await this.client.api(`/drives/${driveId}/items/${parentItemId}:/${filename}:/content`)
                .put(stream);

            return M365WrapperDataResult.createSuccess(result);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }
}