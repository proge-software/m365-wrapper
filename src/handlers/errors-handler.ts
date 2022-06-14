import { AuthError } from "@azure/msal-browser";
import M365WrapperDataResult from "../models/results/m365-wrapper-data-result";
import M365WrapperError  from "../models/results/m365-wrapper-error";
import M365WrapperResult from "../models/results/m365-wrapper-result";

export default class ErrorsHandler {

    public static GetErrorResult(catchedError: any): M365WrapperResult {
        let error: M365WrapperError;

        if (catchedError instanceof AuthError) {
            error = new M365WrapperError(catchedError, catchedError.errorCode);
        }
        else if (catchedError instanceof Error) {
            error = new M365WrapperError(catchedError);
        }
        else {
            throw error;
        }

        return new M365WrapperResult(error);
    }

    public static GetErrorDataResult<TData>(catchedError: any): M365WrapperDataResult<TData> {
        return new M365WrapperDataResult<TData>(this.GetErrorResult(catchedError));
    }
}