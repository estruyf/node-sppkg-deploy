export interface IOptions {
    username?: string;
    password?: string;
    tenant?: string;
    site?: string;
    absoluteUrl?: string;
    filename?: string;
    skipFeatureDeployment?: boolean;
    sp2016?: boolean;
    verbose?: boolean;
}
export interface IWebAndList {
    webId: string;
    listId: string;
}
export interface IFileInfo {
    id: number;
    version: number;
}
