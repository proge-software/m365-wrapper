import { AuthError } from "@azure/msal-browser";
import Result from "../models/results/result";

export default class ErrorsHandler {

    public static GetErrorResult(error: any): Result {
        let result: Result = { isSuccess: false } as Result;

        if (error instanceof AuthError) {
            result.error.code = error.errorCode;
            result.error.message = error.errorMessage;
            result.error.stack = error.stack;
        }
        else if (error instanceof Error) {
            result.error.code = error.name;
            result.error.message = error.message;
            result.error.stack = error.stack;
        }
        else {
            throw error;
        }

        return result;
    }
}