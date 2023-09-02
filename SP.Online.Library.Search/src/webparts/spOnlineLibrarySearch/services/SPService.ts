import { SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/fields";
import "@pnp/sp/items";
import "@pnp/sp/folders";
import "@pnp/sp/files";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IProposalItem } from "../models/IProposalItem";
import { FILETYPE, FieldType, LIBRARY_PROPOSAL, LIST_DATE_FILTER_CONFIG, Properties } from "../common/constants";
import { IDateConfigItem } from "../models/IDateConfigItem";
import { IMetaFieldInfo } from "../models/IMetaFieldInfo";
import { IStore } from "../models/IStore";
import  { TermStore, TermSets, TermSet, Terms } from "@pnp/sp/taxonomy";
import { ISearchResult } from "../models/ISearchResult";
import { ISharePointSearchResponse } from "../models/ISharePointSearchResponse";
import { ICellValue, ICells } from "./ISearchService";
import * as mammoth from "mammoth";

// import { taxonomy, ITermStore, ITermSet } from "@pnp/sp-taxonomy";

export class SPService {

    private _userID: number;
    private _context: any;
    private _sp;
    constructor(context: any) {
        this._context = context;
        this._sp = spfi().using(SPFx(context));
        this._userID = context.pageContext.legacyPageContext.userId;
    }

    public async getDateFilterConfig(): Promise<IDateConfigItem[]> {
        const listName = LIST_DATE_FILTER_CONFIG;
        return this._sp.web.lists.getByTitle(listName).items().then((data) => {
            return data;
        }).catch((error) => {
            return [];
        })
    }

    public async getAgencyFolders(): Promise<any> {
        const libraryName = LIBRARY_PROPOSAL;
        let folders = (await this._sp.web.lists.getByTitle(libraryName).rootFolder.folders());
        folders = folders.sort((a,b) => {
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

    public async getProposalData(searchQuery: string): Promise<IProposalItem[]> {
        const libraryName = 'Proposal Folder';
        const proposalLibrary = this._sp.web.lists.getByTitle(libraryName);
        return proposalLibrary.renderListDataAsStream({
            ViewXml: searchQuery
        }).then((data) => {
            return data.Row;
        }).catch((error) => {
            return [];
        })
    }


    public async getTermSetIdByFieldName(fieldName: string): Promise<any> {
        const libraryName = 'Proposal Folder';
        return this._sp.web.lists.getByTitle(libraryName).fields.getByTitle(fieldName).select('SspId, TermSetId')().then((data) => {
            return data;
        });
    }

    public async getTermGroups(): Promise<any> {
        return this._sp.termStore.groups().then((data) => {
            return data;
        }).catch(() => {
            return [];
        });
    }

    public async getTermStores(): Promise<any> {
        return this._sp.termStore().then((data) => {
            return data;
        }).catch((error) => {
            console.log('error in getting termstore');
            return [];
        })
    }


    public async getTermSetById(id: string): Promise<any> {
      return this._sp.termStore.sets.getById(id).getAllChildrenAsOrderedTree().then((terms) => {
        return terms;
      }).catch((error) => {
        console.log('error in getting terms');
        return [];
      })
    }   


    public getSearchResults(query: string): Promise<ISearchResult[]> {
        const properties = Properties.join(',');
        const siteUrl = this._context.pageContext.web.absoluteUrl;
        const libraryPath = `${siteUrl}/${LIBRARY_PROPOSAL}`;

        const path1 = `${libraryPath}/DMS`;
        const path2 = `${libraryPath}/Sample`;
        //path + path
        let url: string =  siteUrl + 
                          `/_api/search/query?querytext='${query}+site:"${libraryPath}"'`+ 
                          `&$selectproperties=${properties}` +
                          //`&processbestbets=true&enablequeryrules=true` +
                          `&querytemplate='{searchTerms} IsDocument:true'`+
                          `&rowlimit=100`;
        
        return new Promise<ISearchResult[]>((resolve, reject) => {
            // Do an Ajax call to receive the search results
            this._getSearchData(url).then((res: ISharePointSearchResponse | any) => {
                let searchResp: ISearchResult[] = [];
                console.log('search response');
                console.log(res);

                // Check if there was an error
                if (typeof res["odata.error"] !== "undefined") {
                    if (typeof res["odata.error"]["message"] !== "undefined") {
                        Promise.reject(res["odata.error"]["message"].value);
                        return;
                    }
                }

                if (!this._isNull(res)) {
                    const fields: string = properties;

                    // Retrieve all the table rows
                    if (typeof res.PrimaryQueryResult.RelevantResults.Table !== 'undefined') {
                        if (typeof res.PrimaryQueryResult.RelevantResults.Table.Rows !== 'undefined') {
                            searchResp = this._setSearchResults(res.PrimaryQueryResult.RelevantResults.Table.Rows, fields);
                        }
                    }
                }

                // Return the retrieved result set
                resolve(searchResp);
            });
        });
    }

     /**
     * Retrieve the results from the search API
     *
     * @param url
     */
     private _getSearchData(url: string): Promise<ISharePointSearchResponse> {
        return this._context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
            headers: {
                'odata-version': '3.0'
            }
        }).then((res: SPHttpClientResponse) => {
            return res.json();
        }).catch((error: any) => {
            return Promise.reject(JSON.stringify(error));
        });
    }

     /**
     * Set the current set of search results
     *
     * @param crntResults
     * @param fields
     */
     private _setSearchResults(crntResults: ICells[], fields: string): any[] {
        const temp: any[] = [];

        if (crntResults.length > 0) {
            const flds: string[] = fields.toLowerCase().split(',');

            crntResults.forEach((result) => {
                // Create a temp value
                var val: any = {}

                result.Cells.forEach((cell: ICellValue) => {
                    if (flds.indexOf(cell.Key.toLowerCase()) !== -1) {
                        // Add key and value to temp value
                        val[cell.Key] = cell.Value;
                    }
                });

                // Push this to the temp array
                temp.push(val);
            });
        }

        return temp;
    }

        /**
     * Check if the value is null or undefined
     *
     * @param value
     */
        private _isNull(value: any): boolean {
            return value === null || typeof value === "undefined";
        }

    public async getFileContent(fileRelativeUrl: string, fileType: string): Promise<any> {
        const arrayBuffer = await this._getFileAsBufferArray(fileRelativeUrl);
        const fileContent = await mammoth.convertToHtml({arrayBuffer: arrayBuffer});
        return fileContent;
    }    

    private async _getFileAsBufferArray(fileRelativeUrl: string): Promise<any> {
        return this._sp.web.getFileByServerRelativePath(fileRelativeUrl).getBuffer();
    }

}