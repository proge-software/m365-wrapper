export default class M365WrapperError extends Error {
    code: string;

    constructor(error?: Error, code?: string) {
        super(error?.message);

        this.code = code ?? undefined;
        this.name = error?.name ?? undefined;
        this.stack = error?.stack ?? undefined;
    }
}