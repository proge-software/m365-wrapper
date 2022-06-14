import Result from "./result";

export default interface DataResult<TData> extends Result {
    data: TData;
}