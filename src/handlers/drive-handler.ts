import { Client } from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
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

    public async getMyDrives(): Promise<M365WrapperDataResult<[MicrosoftGraph.Drive]>> {
        try {
            let items: [MicrosoftGraph.Drive] = await this.client.api("/me/drives")
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getMyDriveItemsByQuery(queryText: string): Promise<M365WrapperDataResult<[MicrosoftGraph.DriveItem]>> {
        try {
            let items = await this.client.api(`/me/drive/root/search(q='${queryText}')`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getMyDriveAndSharedItemsByQuery(queryText: string): Promise<M365WrapperDataResult<[MicrosoftGraph.DriveItem]>> {
        try {
            let items = await this.client.api(`/me/drive/search(q='${queryText}')`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getMySharedItems(): Promise<M365WrapperDataResult<[MicrosoftGraph.DriveItem]>> {
        try {
            let items = await this.client.api(`/me/drive/sharedWithMe`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getDriveItems(driveId: string): Promise<M365WrapperDataResult<[MicrosoftGraph.DriveItem]>> {
        try {
            let items = await this.client.api(`/drives/${driveId}/root/children`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getDriveItemsByQuery(driveId: string, queryText: string): Promise<M365WrapperDataResult<[MicrosoftGraph.DriveItem]>> {
        try {
            let items = await this.client.api(`/drives/${driveId}/root/search(q='${queryText}')`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getDriveFolderItems(driveId: string, folderId: string): Promise<M365WrapperDataResult<[MicrosoftGraph.DriveItem]>> {
        try {
            let items = await this.client.api(`/drives/${driveId}/items/${folderId}/children`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getDriveItem(driveId: string, itemId: string): Promise<M365WrapperDataResult<MicrosoftGraph.DriveItem>> {
        try {
            let items = await this.client.api(`/drives/${driveId}/items/${itemId}`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }

    public async getMyDriveItemSharingPermissions(itemId: string): Promise<M365WrapperDataResult<[MicrosoftGraph.Permission]>> {
        try {
            let items = await this.client.api(`/me/drive/items/${itemId}/permissions`)
                .get();

            return M365WrapperDataResult.createSuccess(items);
        }
        catch (error) {
            return ErrorsHandler.getErrorDataResult(error);
        }
    }
}