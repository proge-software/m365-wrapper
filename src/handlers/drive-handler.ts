import { Client } from "@microsoft/microsoft-graph-client";
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import OfficeHandler from "./office-handler";

export default class DriveHandler {

    constructor(private readonly client: Client, private readonly office: OfficeHandler) { }

    public async isOneDriveInMyLicenses(): Promise<boolean> {
        try {

            var bFound = false;
            var teamsSkuPartNumbers: string[] = ['O365_BUSINESS',
                'SMB_BUSINESS',
                'OFFICESUBSCRIPTION',
                'WACONEDRIVESTANDARD',
                'WACONEDRIVEENTERPRISE',
                'VISIOONLINE_PLAN1',
                'VISIOCLIENT']

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

            if (!bFound) {
                bFound = await this.office.isInMyLicenses();
            }

            return bFound;
        }
        catch (error) {
            throw error;
        }
    }

    public async getMyDrives(): Promise<[MicrosoftGraph.Drive]> {
        try {
            const items = await this.client.api("/me/drives")
                .get();

            return items;
        }
        catch (error) {
            throw error;
        }
    }

    public async getMyDriveItemsByQuery(queryText: string): Promise<[MicrosoftGraph.DriveItem]> {
        try {
            const items = await this.client.api(`/me/drive/root/search(q='${queryText}')`)
                .get();

            return items;
        }
        catch (error) {
            throw error;
        }
    }

    public async getMyDriveAndSharedItemsByQuery(queryText: string): Promise<[MicrosoftGraph.DriveItem]> {
        try {
            const items = await this.client.api(`/me/drive/search(q='${queryText}')`)
                .get();

            return items;
        }
        catch (error) {
            throw error;
        }
    }

    public async getMySharedItems(): Promise<[MicrosoftGraph.DriveItem]> {
        try {
            const items = await this.client.api(`/me/drive/sharedWithMe`)
                .get();

            return items;
        }
        catch (error) {
            throw error;
        }
    }

    public async getDriveItems(driveId: string): Promise<[MicrosoftGraph.DriveItem]> {
        try {
            const items = await this.client.api(`/drives/${driveId}/root/children`)
                .get();

            return items;
        }
        catch (error) {
            throw error;
        }
    }

    public async getDriveItemsByQuery(driveId: string, queryText: string): Promise<[MicrosoftGraph.DriveItem]> {
        try {
            const items = await this.client.api(`/drives/${driveId}/root/search(q='${queryText}')`)
                .get();

            return items;
        }
        catch (error) {
            throw error;
        }
    }

    public async getDriveFolderItems(driveId: string, folderId: string): Promise<[MicrosoftGraph.DriveItem]> {
        try {
            const items = await this.client.api(`/drives/${driveId}/items/${folderId}/children`)
                .get();

            return items;
        }
        catch (error) {
            throw error;
        }
    }

    public async getDriveItem(driveId: string, itemId: string): Promise<MicrosoftGraph.DriveItem> {
        try {
            const items = await this.client.api(`/drives/${driveId}/items/${itemId}`)
                .get();

            return items;
        }
        catch (error) {
            throw error;
        }
    }

    public async getMyDriveItemSharingPermissions(itemId: string): Promise<[MicrosoftGraph.Permission]> {
        try {
            const items = await this.client.api(`/me/drive/items/${itemId}/permissions`)
                .get();

            return items;
        }
        catch (error) {
            throw error;
        }
    }
}