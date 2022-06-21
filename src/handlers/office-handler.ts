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
            icon: 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHhtbG5zOnhsaW5rPSJodHRwOi8vd3d3LnczLm9yZy8xOTk5L3hsaW5rIiB2aWV3Qm94PSIwIDAgNjQgNjQiPgoJPGRlZnM+CgkJPHN0eWxlPgogICAgICAgIAkJLmNscy0xewoJCQkJCWlzb2xhdGlvbjppc29sYXRlOwoJCQkJfQoKCQkJCS5jbHMtMnsKCQkJCQlvcGFjaXR5OjAuMjsKCQkJCX0KCgkJCQkuY2xzLTIsLmNscy0zLC5jbHMtNXsKCQkJCQltaXgtYmxlbmQtbW9kZTptdWx0aXBseTsKCQkJCX0KCgkJCQkuY2xzLTR7CgkJCQkJZmlsbDojZmZmOwoJCQkJfQoKCQkJCS5jbHMtNXsKCQkJCQlvcGFjaXR5OjAuMTI7CgkJCQl9CgoJCQkJLmNscy02ewoJCQkJCWZpbGw6dXJsKCNsaW5lYXItZ3JhZGllbnQpOwoJCQkJfQoKCQkJCS5jbHMtN3sKCQkJCQlmaWxsOnVybCgjbGluZWFyLWdyYWRpZW50LTIpOwoJCQkJfQoKCQkJCS5jbHMtOHsKCQkJCQlmaWxsOnVybCgjbGluZWFyLWdyYWRpZW50LTMpOwoJCQkJfQoKCQkJCS5jbHMtOXsKCQkJCQlmaWxsOnVybCgjbGluZWFyLWdyYWRpZW50LTQpOwoJCQkJfQoKCQkJCS5jbHMtMTB7CgkJCQkJZmlsbDp1cmwoI2xpbmVhci1ncmFkaWVudC01KTsKCQkJCX0KCgkJCQkuY2xzLTExewoJCQkJCWZpbGw6bm9uZTsKCQkJCX0KCQk8L3N0eWxlPgoKCQk8bGluZWFyR3JhZGllbnQgaWQ9ImxpbmVhci1ncmFkaWVudCIgeDE9IjQ1LjIiIHkxPSItMS40MiIgeDI9IjQ1LjIiIHkyPSI1Ny44IiBncmFkaWVudFVuaXRzPSJ1c2VyU3BhY2VPblVzZSI+PHN0b3Agb2Zmc2V0PSIwIiBzdG9wLWNvbG9yPSIjZmZiOTAwIi8+PHN0b3Agb2Zmc2V0PSIwLjE3IiBzdG9wLWNvbG9yPSIjZWY4NDAwIi8+PHN0b3Agb2Zmc2V0PSIwLjMxIiBzdG9wLWNvbG9yPSIjZTI1YzAxIi8+PHN0b3Agb2Zmc2V0PSIwLjQzIiBzdG9wLWNvbG9yPSIjZGI0NDAxIi8+PHN0b3Agb2Zmc2V0PSIwLjUiIHN0b3AtY29sb3I9IiNkODNiMDEiLz48L2xpbmVhckdyYWRpZW50PgoJCTxsaW5lYXJHcmFkaWVudCBpZD0ibGluZWFyLWdyYWRpZW50LTIiIHgxPSIzNC41MiIgeTE9IjAuNjciIHgyPSIzLjE2IiB5Mj0iNDUuNDUiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIj48c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiM4MDA2MDAiLz48c3RvcCBvZmZzZXQ9IjAuNiIgc3RvcC1jb2xvcj0iI2M3MjEyNyIvPjxzdG9wIG9mZnNldD0iMC43MyIgc3RvcC1jb2xvcj0iI2MxMzk1OSIvPjxzdG9wIG9mZnNldD0iMC44NSIgc3RvcC1jb2xvcj0iI2JjNGI4MSIvPjxzdG9wIG9mZnNldD0iMC45NCIgc3RvcC1jb2xvcj0iI2I5NTc5OSIvPjxzdG9wIG9mZnNldD0iMSIgc3RvcC1jb2xvcj0iI2I4NWJhMiIvPjwvbGluZWFyR3JhZGllbnQ+CgkJPGxpbmVhckdyYWRpZW50IGlkPSJsaW5lYXItZ3JhZGllbnQtMyIgeDE9IjE4LjUiIHkxPSI1NS42MyIgeDI9IjU5LjQ0IiB5Mj0iNTUuNjMiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIj48c3RvcCBvZmZzZXQ9IjAiIHN0b3AtY29sb3I9IiNmMzJiNDQiLz48c3RvcCBvZmZzZXQ9IjAuNiIgc3RvcC1jb2xvcj0iI2E0MDcwYSIvPjwvbGluZWFyR3JhZGllbnQ+CgkJPGxpbmVhckdyYWRpZW50IGlkPSJsaW5lYXItZ3JhZGllbnQtNCIgeDE9IjM1LjE2IiB5MT0iLTAuMjQiIHgyPSIyOC41MiIgeTI9IjkuMjQiIGdyYWRpZW50VW5pdHM9InVzZXJTcGFjZU9uVXNlIj48c3RvcCBvZmZzZXQ9IjAiIHN0b3Atb3BhY2l0eT0iMC40Ii8+PHN0b3Agb2Zmc2V0PSIxIiBzdG9wLW9wYWNpdHk9IjAiLz48L2xpbmVhckdyYWRpZW50PgoJCTxsaW5lYXJHcmFkaWVudCBpZD0ibGluZWFyLWdyYWRpZW50LTUiIHgxPSI0Ni4zMiIgeTE9IjU2LjU1IiB4Mj0iMjcuOTkiIHkyPSI1NC45NSIgZ3JhZGllbnRVbml0cz0idXNlclNwYWNlT25Vc2UiPjxzdG9wIG9mZnNldD0iMCIgc3RvcC1vcGFjaXR5PSIwLjQiLz48c3RvcCBvZmZzZXQ9IjEiIHN0b3Atb3BhY2l0eT0iMCIvPjwvbGluZWFyR3JhZGllbnQ+CiAgCTwvZGVmcz4KICAKIDxnIGNsYXNzPSJjbHMtMSI+Cgk8ZyBpZD0iSWNvbnNfLV9Db2xvciIgZGF0YS1uYW1lPSJJY29ucyAtIENvbG9yIj4KCQk8ZyBpZD0iRGVza3RvcF8tX0Z1bGxfQmxlZWQiIGRhdGEtbmFtZT0iRGVza3RvcCAtIEZ1bGwgQmxlZWQiPgoJCQk8ZyBjbGFzcz0iY2xzLTIiPgoJCQkJPHBhdGggY2xhc3M9ImNscy00IiBkPSJNMTkuOTMsNDlhMy4yMiwzLjIyLDAsMCwwLTEuNTksNkwyOS43LDYxLjQ0YTYuMiw2LjIsMCwwLDAsMy4wNy44MUE2LDYsMCwwLDAsMzQuNDgsNjJsMTcuMDktNC44N0E2LjEyLDYuMTIsMCwwLDAsNTYsNTEuMjZWNDlaIi8+CiAgICAgICAgICAJPC9nPgoKICAgICAgICAgIAk8ZyBjbGFzcz0iY2xzLTUiPgoJCQkJPHBhdGggY2xhc3M9ImNscy00IiBkPSJNMTkuOTMsNDlhMy4yMiwzLjIyLDAsMCwwLTEuNTksNkwyOS43LDYxLjQ0YTYuMiw2LjIsMCwwLDAsMy4wNy44MUE2LDYsMCwwLDAsMzQuNDgsNjJsMTcuMDktNC44N0E2LjEyLDYuMTIsMCwwLDAsNTYsNTEuMjZWNDlaIi8+CiAgICAgICAgIAk8L2c+CgkJCQk8cGF0aCBjbGFzcz0iY2xzLTYiIGQ9Ik0zNC40MSwyLDM5LDEyLjVWNDlMMzQuNDgsNjJsMTcuMDktNC44N0E2LjEyLDYuMTIsMCwwLDAsNTYsNTEuMjZWMTIuNzRhNi4xMSw2LjExLDAsMCwwLTQuNDQtNS44OFoiLz4KCQkJCTxwYXRoIGNsYXNzPSJjbHMtNyIgZD0iTTEyLjc0LDQ4LjYxbDUtMi43QTQuMzYsNC4zNiwwLDAsMCwyMCw0Mi4wOFYyMi40M2E0LjM3LDQuMzcsMCwwLDEsMi44Ny00LjFMMzksMTIuNVY4LjA3QTYuMzIsNi4zMiwwLDAsMCwzNC40MSwyYTYuMTgsNi4xOCwwLDAsMC0xLjczLS4yNGgwYTYuNDEsNi40MSwwLDAsMC0zLjE0LjgzTDExLjA4LDEzLjEyQTYuMSw2LjEsMCwwLDAsOCwxOC40MlY0NS43OEEzLjIxLDMuMjEsMCwwLDAsMTIuNzQsNDguNjFaIi8+CgkJCQk8cGF0aCBjbGFzcz0iY2xzLTgiIGQ9Ik0zOSw0OUgxOS45M2EzLjIyLDMuMjIsMCwwLDAtMS41OSw2TDI5LjcsNjEuNDRhNi4yLDYuMiwwLDAsMCwzLjA3LjgxaDBBNiw2LDAsMCwwLDM0LjQ4LDYyLDYuMjIsNi4yMiwwLDAsMCwzOSw1NloiLz4KCQkJCTxwYXRoIGNsYXNzPSJjbHMtOSIgZD0iTTEyLjc0LDQ4LjYxbDUtMi43QTQuMzYsNC4zNiwwLDAsMCwyMCw0Mi4wOFYyMi40M2E0LjM3LDQuMzcsMCwwLDEsMi44Ny00LjFMMzksMTIuNVY4LjA3QTYuMzIsNi4zMiwwLDAsMCwzNC40MSwyYTYuMTgsNi4xOCwwLDAsMC0xLjczLS4yNGgwYTYuNDEsNi40MSwwLDAsMC0zLjE0LjgzTDExLjA4LDEzLjEyQTYuMSw2LjEsMCwwLDAsOCwxOC40MlY0NS43OEEzLjIxLDMuMjEsMCwwLDAsMTIuNzQsNDguNjFaIi8+CgkJCQk8cGF0aCBjbGFzcz0iY2xzLTEwIiBkPSJNMzksNDlIMTkuOTNhMy4yMiwzLjIyLDAsMCwwLTEuNTksNkwyOS43LDYxLjQ0YTYuMiw2LjIsMCwwLDAsMy4wNy44MWgwQTYsNiwwLDAsMCwzNC40OCw2Miw2LjIyLDYuMjIsMCwwLDAsMzksNTZaIi8+CgkJCQk8cmVjdCBjbGFzcz0iY2xzLTExIiB3aWR0aD0iNjQiIGhlaWdodD0iNjQiLz4KCQk8L2c+Cgk8L2c+CjwvZz4KPC9zdmc+Cg=='
        },{
            name: 'Word',
            link: 'https://www.office.com/launch/word',
            icon: 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCI+PGRlZnM+PHN0eWxlPi5jbHMtMXtmaWxsOm5vbmU7fS5jbHMtMntmaWxsOiM0MWE1ZWU7fS5jbHMtM3tmaWxsOiMyYjdjZDM7fS5jbHMtNHtmaWxsOiMxODVhYmQ7fS5jbHMtNXtmaWxsOiMxMDNmOTE7fS5jbHMtNntvcGFjaXR5OjAuNTt9LmNscy03e2ZpbGw6I2ZmZjt9PC9zdHlsZT48L2RlZnM+PHRpdGxlPldvcmRfMjR4PC90aXRsZT48ZyBpZD0iV29yZCI+PGcgaWQ9Il8yNCIgZGF0YS1uYW1lPSIyNCI+PHJlY3QgY2xhc3M9ImNscy0xIiB3aWR0aD0iMjQiIGhlaWdodD0iMjQiLz48cGF0aCBjbGFzcz0iY2xzLTIiIGQ9Ik0yNCw3VjJhMSwxLDAsMCwwLTEtMUg3QTEsMSwwLDAsMCw2LDJWN2w5LDJaIi8+PHBvbHlnb24gY2xhc3M9ImNscy0zIiBwb2ludHM9IjI0IDcgNiA3IDYgMTIgMTUuNSAxNCAyNCAxMiAyNCA3Ii8+PHBvbHlnb24gY2xhc3M9ImNscy00IiBwb2ludHM9IjI0IDEyIDYgMTIgNiAxNyAxNSAxOC41IDI0IDE3IDI0IDEyIi8+PHBhdGggY2xhc3M9ImNscy01IiBkPSJNNiwxN0gyNGEwLDAsMCwwLDEsMCwwdjVhMSwxLDAsMCwxLTEsMUg3YTEsMSwwLDAsMS0xLTFWMTdhMCwwLDAsMCwxLDAsMFoiLz48cGF0aCBjbGFzcz0iY2xzLTYiIGQ9Ik0xMy44Myw2SDZWMjBoNy42QTEuNSwxLjUsMCwwLDAsMTUsMTguNjVWNy4xN0ExLjE4LDEuMTgsMCwwLDAsMTMuODMsNloiLz48cmVjdCBpZD0iQmFja19QbGF0ZSIgZGF0YS1uYW1lPSJCYWNrIFBsYXRlIiBjbGFzcz0iY2xzLTQiIHk9IjUiIHdpZHRoPSIxNCIgaGVpZ2h0PSIxNCIgcng9IjEuMTciLz48cGF0aCBpZD0iTGV0dGVyIiBjbGFzcz0iY2xzLTciIGQ9Ik0xMC4xNiwxNkg4LjcyTDcsMTAuNDgsNS4yOCwxNkgzLjg0TDIuMjQsOEgzLjY4TDQuOCwxMy42LDYuNDgsOC4xNmgxLjJsMS42LDUuNDRMMTAuNCw4aDEuMzZaIi8+PC9nPjwvZz48L3N2Zz4='
        },{
            name: 'Excel',
            link: 'https://www.office.com/launch/excel',
            icon: 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCI+PGRlZnM+PHN0eWxlPi5jbHMtMXtmaWxsOiMyMWEzNjY7fS5jbHMtMntmaWxsOm5vbmU7fS5jbHMtM3tmaWxsOiMxMDdjNDE7fS5jbHMtNHtmaWxsOiMzM2M0ODE7fS5jbHMtNXtmaWxsOiMxODVjMzc7fS5jbHMtNntvcGFjaXR5OjAuNTt9LmNscy03e2ZpbGw6I2ZmZjt9PC9zdHlsZT48L2RlZnM+PHRpdGxlPkV4Y2VsXzI0eDwvdGl0bGU+PGcgaWQ9IkV4Y2VsIj48ZyBpZD0iXzI0IiBkYXRhLW5hbWU9IjI0Ij48cGF0aCBjbGFzcz0iY2xzLTEiIGQ9Ik0xNiwxSDdBMSwxLDAsMCwwLDYsMlY3bDEwLDUsNCwxLjVMMjQsMTJWN1oiLz48cmVjdCBjbGFzcz0iY2xzLTIiIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIvPjxyZWN0IGNsYXNzPSJjbHMtMyIgeD0iNiIgeT0iNy4wMiIgd2lkdGg9IjEwIiBoZWlnaHQ9IjQuOTgiLz48cGF0aCBjbGFzcz0iY2xzLTQiIGQ9Ik0yNCwyVjdIMTZWMWg3QTEsMSwwLDAsMSwyNCwyWiIvPjxwYXRoIGNsYXNzPSJjbHMtNSIgZD0iTTE2LDEySDZWMjJhMSwxLDAsMCwwLDEsMUgyM2ExLDEsMCwwLDAsMS0xVjE3WiIvPjxwYXRoIGNsYXNzPSJjbHMtNiIgZD0iTTEzLjgzLDZINlYyMGg3LjZBMS41LDEuNSwwLDAsMCwxNSwxOC42NVY3LjE3QTEuMTgsMS4xOCwwLDAsMCwxMy44Myw2WiIvPjxyZWN0IGlkPSJCYWNrX1BsYXRlIiBkYXRhLW5hbWU9IkJhY2sgUGxhdGUiIGNsYXNzPSJjbHMtMyIgeT0iNSIgd2lkdGg9IjE0IiBoZWlnaHQ9IjE0IiByeD0iMS4xNyIvPjxwYXRoIGNsYXNzPSJjbHMtNyIgZD0iTTMuNDMsMTYsNiwxMiwzLjY0LDhINS41NWwxLjMsMi41NWE0LjYzLDQuNjMsMCwwLDEsLjI0LjU0aDBhNS43Nyw1Ljc3LDAsMCwxLC4yNy0uNTZMOC43Niw4aDEuNzVMOC4wOCwxMmwyLjQ5LDRIOC43MWwtMS41LTIuOEEyLjE0LDIuMTQsMCwwLDEsNywxMi44M0g3YTEuNTQsMS41NCwwLDAsMS0uMTcuMzZMNS4zLDE2WiIvPjxyZWN0IGNsYXNzPSJjbHMtMyIgeD0iMTYiIHk9IjEyIiB3aWR0aD0iOCIgaGVpZ2h0PSI1Ii8+PC9nPjwvZz48L3N2Zz4='
        },{
            name: 'PowerPoint',
            link: 'https://www.office.com/launch/powerpoint',
            icon: 'data:image/svg+xml;base64,PHN2ZyB4bWxucz0iaHR0cDovL3d3dy53My5vcmcvMjAwMC9zdmciIHZpZXdCb3g9IjAgMCAyNCAyNCI+PGRlZnM+PHN0eWxlPi5jbHMtMXtmaWxsOm5vbmU7fS5jbHMtMntmaWxsOiNlZDZjNDc7fS5jbHMtM3tmaWxsOiNmZjhmNmI7fS5jbHMtNHtmaWxsOiNkMzUyMzA7fS5jbHMtNXtvcGFjaXR5OjAuNTt9LmNscy02e29wYWNpdHk6MC4xO30uY2xzLTd7ZmlsbDojYzQzZTFjO30uY2xzLTh7ZmlsbDojZmZmO308L3N0eWxlPjwvZGVmcz48dGl0bGU+UG93ZXJwb2ludF8yNHg8L3RpdGxlPjxnIGlkPSJQb3dlcnBvaW50Ij48cmVjdCBjbGFzcz0iY2xzLTEiIHdpZHRoPSIyNCIgaGVpZ2h0PSIyNCIvPjxwYXRoIGNsYXNzPSJjbHMtMiIgZD0iTTEzLDFBMTEsMTEsMCwwLDAsMiwxMmwxNC44NCwzLjg0WiIvPjxwYXRoIGNsYXNzPSJjbHMtMyIgZD0iTTEzLDFBMTEsMTEsMCwwLDEsMjQsMTJMMTguNSwxNSwxMywxMloiLz48cGF0aCBjbGFzcz0iY2xzLTQiIGQ9Ik0yLDEyYTExLDExLDAsMCwwLDIyLDBaIi8+PHBhdGggY2xhc3M9ImNscy01IiBkPSJNMTUsMTguNjVWNy4xN0ExLjE4LDEuMTgsMCwwLDAsMTMuODMsNkgzLjhBMTAuOTEsMTAuOTEsMCwwLDAsNS40OSwyMEgxMy42QTEuNSwxLjUsMCwwLDAsMTUsMTguNjVaIi8+PHBhdGggY2xhc3M9ImNscy02IiBkPSJNMTUsMTguNjVWNy4xN0ExLjE4LDEuMTgsMCwwLDAsMTMuODMsNkgzLjhBMTAuOTEsMTAuOTEsMCwwLDAsNS40OSwyMEgxMy42QTEuNSwxLjUsMCwwLDAsMTUsMTguNjVaIi8+PHJlY3QgaWQ9IkJhY2tfUGxhdGUiIGRhdGEtbmFtZT0iQmFjayBQbGF0ZSIgY2xhc3M9ImNscy03IiB5PSI1IiB3aWR0aD0iMTQiIGhlaWdodD0iMTQiIHJ4PSIxLjE3Ii8+PHBhdGggY2xhc3M9ImNscy04IiBkPSJNNy40LDhhMy4zMiwzLjMyLDAsMCwxLDIuMi42NCwyLjMyLDIuMzIsMCwwLDEsLjc2LDEuODZBMy40MiwzLjQyLDAsMCwxLDEwLDEyLjExYTIuNTQsMi41NCwwLDAsMS0xLjA3LDEsMy43LDMuNywwLDAsMS0xLjYxLjM0SDUuNzhWMTZINC4yMlY4Wk01Ljc4LDEySDcuMTJhMS43OCwxLjc4LDAsMCwwLDEuMTktLjM1LDEuNDYsMS40NiwwLDAsMCwuNC0xLjFjMC0uODgtLjUxLTEuMzItMS41NC0xLjMySDUuNzhaIi8+PC9nPjwvc3ZnPg=='
        }]);
    }
}