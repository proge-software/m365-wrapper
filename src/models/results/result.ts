import Error from "./error-result";

export default interface Result {
    isSuccess: boolean;
    error: Error;
}