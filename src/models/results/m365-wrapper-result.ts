import M365WrapperError from "./m365-wrapper-error";

export default class M365WrapperResult {
    isSuccess: boolean;
    error: M365WrapperError;

    constructor(error?: M365WrapperError) {
        if(error) {
            this.isSuccess = false;
            this.error = error;
        }
        else {
            this.isSuccess = true;
        }
    }
}