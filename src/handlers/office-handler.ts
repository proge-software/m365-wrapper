import { Client } from "@microsoft/microsoft-graph-client";

export default class OfficeHandler {

    constructor(private readonly client: Client) { }

    public async IsOfficeInMyLicenses(): Promise<boolean> {
        try {

            var bFound = false;
            var teamsSkuPartNumbers: string[] = ['M365EDU_A3_FACULTY',
                'M365EDU_A3_STUDENT',
                'M365EDU_A5_FACULTY',
                'M365EDU_A5_STUDENT',
                'O365_BUSINESS',
                'SMB_BUSINESS',
                'OFFICESUBSCRIPTION',
                'O365_BUSINESS_ESSENTIALS', // Mobile
                'SMB_BUSINESS_ESSENTIALS', // Mobile
                'O365_BUSINESS_PREMIUM',
                'SMB_BUSINESS_PREMIUM',
                'SPB',
                'SPE_E3',
                'SPE_E5',
                'SPE_E3_USGOV_DOD',
                'SPE_E3_USGOV_GCCHIGH',
                'SPE_F1', // Mobile
                'ENTERPRISEPREMIUM_FACULTY',
                'ENTERPRISEPREMIUM_STUDENT',
                'STANDARDPACK', // Mobile
                'ENTERPRISEPACK',
                'DEVELOPERPACK',
                'ENTERPRISEPACK_USGOV_DOD',
                'ENTERPRISEPACK_USGOV_GCCHIGH',
                'ENTERPRISEWITHSCAL',
                'ENTERPRISEPREMIUM',
                'ENTERPRISEPREMIUM_NOPSTNCONF',
                'DESKLESSPACK', // Mobile
                'MIDSIZEPACK',
                'LITEPACK_P2']

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
}