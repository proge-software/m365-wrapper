import { Client } from "@microsoft/microsoft-graph-client";
import M365App from "../models/results/m365-app";
import M365WrapperDataResult from "../models/results/m365-wrapper-data-result";
import M365WrapperResult from "../models/results/m365-wrapper-result";
import ErrorsHandler from "./errors-handler";

export default class OfficeHandler {

    constructor(private readonly client: Client) { }

    public async isInMyLicenses(): Promise<M365WrapperResult> {
        try {

            let result: M365WrapperResult = { isSuccess: false } as M365WrapperResult;
            let teamsSkuPartNumbers: string[] = ['M365EDU_A3_FACULTY',
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

    public getApps(): M365WrapperDataResult<M365App[]>{
        return new M365WrapperDataResult(null, [{
            name: 'Office',
            link: 'https://www.office.com',
            icon: ''
        },{
            name: 'Word',
            link: 'https://www.office.com/launch/word',
            icon: ''
        },{
            name: 'Excel',
            link: 'https://www.office.com/launch/excel',
            icon: ''
        },{
            name: 'PowerPoint',
            link: 'https://www.office.com/launch/powerpoint',
            icon: ''
        }]);
    }
}