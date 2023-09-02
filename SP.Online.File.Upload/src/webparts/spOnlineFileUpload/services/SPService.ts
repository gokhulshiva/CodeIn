import { SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { HttpRequestError } from "@pnp/queryable";
import { fileFromServerRelativePath, IFileInfo, IFileUploadProgressData, Version, Versions } from "@pnp/sp/files";
import { CheckinType, IFileAddResult, IItem, IRoleAssignmentInfo, ISiteUserInfo, PermissionKind, RoleDefinition, folderFromServerRelativePath } from "@pnp/sp/presets/all";

import { LIBRARY_PROPOSAL, Result } from "../common/constants";
import { IFileUploadResult } from "../models/IFileUploadResult";

export class SPService {

    private _userID: number;
    private _context: any;
    private _sp;

    constructor(context: any) {
        this._context = context;
        this._sp = spfi().using(SPFx(context));
        this._userID = context.pageContext.legacyPageContext.userId;
    }

    public async getAgencyFolders(): Promise<any> {
        const libraryName = LIBRARY_PROPOSAL;
        let folders = (await this._sp.web.lists.getByTitle(libraryName).rootFolder.folders());
        folders = folders.sort((a, b) => {
            const nameA = a.Name.toLowerCase();
            const nameB = b.Name.toLowerCase();
            if (nameA < nameB) {
                return -1;
            }
            if (nameA > nameB) {
                return 1;
            }
            return 0;
        });
        return folders;
    }


    public async getCategories(): Promise<any> {
        const libraryName = LIBRARY_PROPOSAL;
        return this._sp.web.lists.getByTitle(libraryName).fields.getByTitle('Category')().then((data) => {
            return data.Choices
        }).catch(() => {
            return []
        });
    }

    public async getDomains(): Promise<any> {
        const libraryName = LIBRARY_PROPOSAL;
        return this._sp.web.lists.getByTitle(libraryName).fields.getByTitle('Subject Domain')().then((data) => {
            return data.Choices
        }).catch(() => {
            return []
        });
    }

    public async checkIfFileExists(fileRelativePath: string): Promise<any> {
        return new Promise<any>((resolve) => {
            this._sp.web.getFileByServerRelativePath(fileRelativePath).select('Exists')().then((file: IFileInfo) => {
                if (file.Exists) {
                    fileFromServerRelativePath(this._sp.web, fileRelativePath).getItem().then((filetItem: any) => {
                        const authorId = filetItem.AuthorId;
                        resolve({ isFileExists: true, authorId: authorId });
                    });
                }
                else {
                    resolve({ isFileExists: false, authorId: -1 });
                }
            }).catch((error) => {
                resolve({ isFileExists: false, authorId: -1 });
            });
        });
    }

    public async uploadFile(file: any, libraryPath: string): Promise<IFileUploadResult> {
        const fileNamePath = file.name; //encodeURI(file.name);
        const context = this._context;
        const hostUrl = context.pageContext.legacyPageContext.siteAbsoluteUrl.replace(context.pageContext.legacyPageContext.siteServerRelativeUrl, "");
        return new Promise<IFileUploadResult>((resolve, reject) => {
            folderFromServerRelativePath(this._sp.web, libraryPath).files
                .addUsingPath(fileNamePath, file, { Overwrite: true })
                .then((spFile: IFileAddResult) => {
                    console.log('spfile', spFile);
                    spFile.file.getItem().then((fileItem: any) => {
                        const fileID = fileItem.ID;
                        const fileUrl = spFile.data && spFile.data.ServerRelativeUrl ?
                            `${hostUrl}${spFile.data.ServerRelativeUrl}` : "";
                        console.log('file data');
                        console.log(spFile);
                        resolve({ result: Result.SUCCESS, fileID, fileUrl, error: "" });
                    })
                }).catch((error: any) => {
                    this.getErrorMessage(error).then((errorMessage) => {
                        resolve({ result: Result.ERROR, fileID: -1, fileUrl: "", error: errorMessage });
                    });
                });
        });
    }

    public async updateFile(libraryName: string, fileID: number, body: any): Promise<string> {
        return this._sp.web.lists.getByTitle(libraryName).items
            .getById(fileID)
            .update(body)
            .then((data) => {
                return Result.SUCCESS;
            }).catch((error) => {
                return Result.ERROR
            });
    }

    public async getFileById(fileID: number): Promise<any> {
        const libraryName = LIBRARY_PROPOSAL;
        return this._sp.web.lists.getByTitle(libraryName)
            .items
            .getById(fileID)
            .select('*,FileRef,FileLeafRef')
            .expand('File')()
            .then((data) => {
                return data;
            }).catch((error) => {
                console.log('get file item error: ');
                console.log(error);
                return undefined;
            });
    }

    private async getErrorMessage(e: any): Promise<string> {
        let message: string = "";
        // are we dealing with an HttpRequestError?
        if (e?.isHttpRequestError) {

            // we can read the json from the response
            const json = await (<HttpRequestError>e).response.json();

            // if we have a value property we can show it
            message = typeof json["odata.error"] === "object" ? json["odata.error"].message.value : e.message;

            // add of course you have access to the other properties and can make choices on how to act
            // if ((<HttpRequestError>e).status === 404) {
            //     console.error((<HttpRequestError>e).statusText);
            //     // maybe create the resource, or redirect, or fallback to a secondary data source
            //     // just ideas, handle any of the status codes uniquely as needed
            // }

        } else {
            // not an HttpRequestError so we just log message
            message = e.message;
        }

        return message;
    }
}