import M365WrapperResult from "./m365-wrapper-result";

export default class M365WrapperDataResult<TData> extends M365WrapperResult {
    data: TData;

    constructor(result?: M365WrapperResult, data?: TData) {
        super(result?.error);

        if (data)
            this.data = data;
    }

    static createSuccess<TData>(data: TData): M365WrapperDataResult<TData> {
        return new M365WrapperDataResult<TData>(undefined, data);
    }
}