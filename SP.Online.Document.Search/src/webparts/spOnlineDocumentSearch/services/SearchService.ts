import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import { ISearchResult } from '../../../models/ISearchResult';
import { ISearchResults, ICells, ICellValue, ISearchResponse } from './ISearchService';
import { cloneDeep, escape, isEmpty } from '@microsoft/sp-lodash-subset';
import { ISharePointSearchResponse } from '../../../models/search/ISharePointSearchResponse';
import { ISharePointSearchQuery } from '../../../models/search/ISharePointSearchQuery';
import { ISharePointSearchPromotedResult, ISharePointSearchResult, ISharePointSearchResultBlock, ISharePointSearchResults } from '../../../models/search/ISharePointSearchResults';
import { Constants } from '../../../common/Constants';
import { FilterComparisonOperator, IDataFilterResult, IDataFilterResultValue } from '@pnp/modern-search-extensibility';
import { Properties } from '../../../common/WPConstants';

export default class SearchService {

    /**
     * The SharePoint search service endpoint REST URL
     */
    private searchEndpointUrl: string;

    /**
     * The SPHttpClient instance
     */
    private spHttpClient: SPHttpClient;

    /**
     * The SPHttpClient instance
     */
    private context: any;
    
    constructor(_context: any) {
        this.context = _context;
        this.spHttpClient = _context.spHttpClient;
        this.searchEndpointUrl = `${this.context.pageContext.web.absoluteUrl}/_api/search/postquery`;
    }

    public getSearchResults1(query: string): Promise<ISearchResult[]> {
        const properties = Properties.join(',');
        let url: string = this.context.pageContext.web.absoluteUrl + 
                          `/_api/search/query?querytext='${query}'`+ 
                          `&$selectproperties=${properties}` +
                          `&processbestbets=true&enablequeryrules=true` +
                          "&querytemplate='{searchterms} IsDocument:true'";
        
        return new Promise<ISearchResult[]>((resolve, reject) => {
            // Do an Ajax call to receive the search results
            this._getSearchData(url).then((res: ISharePointSearchResponse) => {
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
        return this.context.spHttpClient.get(url, SPHttpClient.configurations.v1, {
            headers: {
                'odata-version': '3.0'
            }
        }).then((res: SPHttpClientResponse) => {
            return res.json();
        }).catch(error => {
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
                var val: Object = {}

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


      /**
     * Performs a search query against SharePoint
     * @param searchQuery The search query in KQL format
     * @return The search results
     */
      public async search(searchQuery: ISharePointSearchQuery): Promise<ISharePointSearchResults> {

        let results: ISharePointSearchResults = {
            queryKeywords: searchQuery.Querytext,
            refinementResults: [],
            relevantResults: [],
            secondaryResults: [],
            totalRows: 0
        };

        try {

            const response = await this.spHttpClient.post(this.searchEndpointUrl, SPHttpClient.configurations.v1, {
                body: this.getRequestPayload(searchQuery),
                headers: {
                    'odata-version': '3.0',
                    'accept': 'application/json;odata=nometadata',
                    'X-ClientService-ClientTag': Constants.X_CLIENTSERVICE_CLIENTTAG,
                    'UserAgent': Constants.X_CLIENTSERVICE_CLIENTTAG
                }
            });

            if (response.ok) {
                const searchResponse: ISharePointSearchResponse = await response.json();

                if (searchResponse.PrimaryQueryResult) {

                    let refinementResults: IDataFilterResult[] = [];

                    // Get the transformed query submitted to SharePoint
                    const properties = searchResponse.PrimaryQueryResult.RelevantResults.Properties.filter((property) => {
                        return property.Key === 'QueryModification';
                    });

                    if (properties.length === 1) {
                        results.queryModification = properties[0].Value;
                    }

                    const resultRows = searchResponse.PrimaryQueryResult.RelevantResults.Table.Rows;
                    let refinementResultsRows = searchResponse.PrimaryQueryResult.RefinementResults;

                    const refinementRows: any = refinementResultsRows ? refinementResultsRows.Refiners : [];

                    // Map search results
                    let searchResults: ISharePointSearchResult[] = this.getSearchResults(resultRows);

                    // Map refinement results
                    refinementRows.forEach((refiner) => {

                        let values: IDataFilterResultValue[] = [];
                        refiner.Entries.forEach((item) => {
                            values.push({
                                count: parseInt(item.RefinementCount, 10),
                                name: item.RefinementValue.replace("string;#", ""), // Replace string;# for calculated columns https://github.com/SharePoint/sp-dev-solutions/issues/304
                                value: item.RefinementToken,
                                operator: FilterComparisonOperator.Contains
                            } as IDataFilterResultValue);
                        });

                        refinementResults.push({
                            filterName: refiner.Name,
                            values: values
                        });
                    });

                    results.relevantResults = searchResults;
                    results.refinementResults = refinementResults;
                    results.totalRows = searchResponse.PrimaryQueryResult.RelevantResults.TotalRows;

                    if (!isEmpty(searchResponse.SpellingSuggestion)) {
                        results.spellingSuggestion = searchResponse.SpellingSuggestion;
                    }
                }

                // Query rules handling
                if (searchResponse.SecondaryQueryResults) {

                    const secondaryQueryResults = searchResponse.SecondaryQueryResults;

                    if (Array.isArray(secondaryQueryResults) && secondaryQueryResults.length > 0) {

                        let promotedResults: ISharePointSearchPromotedResult[] = [];
                        let secondaryResults: ISharePointSearchResultBlock[] = [];

                        secondaryQueryResults.forEach((e) => {

                            // Best bets are mapped through the "SpecialTermResults" https://msdn.microsoft.com/en-us/library/dd907265(v=office.12).aspx
                            if (e.SpecialTermResults) {
                                // Casting as pnpjs has an incorrect mapping of SpecialTermResults
                                (e.SpecialTermResults).Results.forEach((result) => {
                                    promotedResults.push({
                                        title: result.Title,
                                        url: result.Url,
                                        description: result.Description
                                    } as ISharePointSearchPromotedResult);
                                });
                            }

                            // Secondary/Query Rule results are mapped through SecondaryQueryResults.RelevantResults
                            if (e.RelevantResults) {
                                const secondaryResultItems = this.getSearchResults(e.RelevantResults.Table.Rows);

                                const secondaryResultBlock: ISharePointSearchResultBlock = {
                                    title: e.RelevantResults.ResultTitle,
                                    results: secondaryResultItems
                                };

                                // Only keep secondary result blocks which have items
                                if (secondaryResultBlock.results.length > 0) {
                                    secondaryResults.push(secondaryResultBlock);
                                }
                            }
                        });

                        results.promotedResults = promotedResults;
                        results.secondaryResults = secondaryResults;
                    }
                }
                return results;
            } else {
                throw new Error(`${response['statusMessage']}`);
            }

        } catch (error) {
            console.log('error', error);
        }
    }

    /**
     * Extracts search results from search response rows
     * @param resultRows the search result rows
     */
    private getSearchResults(resultRows: any): ISharePointSearchResult[] {

        // Map search results
        let searchResults: ISharePointSearchResult[] = resultRows.map((elt) => {

            // Build item result dynamically
            // We can't type the response here because search results are by definition too heterogeneous so we treat them as key-value object
            let result: ISharePointSearchResult = {};

            elt.Cells.map((item) => {
                if (item.Key === "HtmlFileType" && item.Value) {
                    result["FileType"] = item.Value;
                }
                else if (!result[item.Key]) {
                    result[item.Key] = item.Value;
                }
            });

            return result;
        });

        return searchResults;
    }

    private getRequestPayload(searchQuery: ISharePointSearchQuery): string {

        let queryPayload: any = cloneDeep(searchQuery);

        queryPayload.HitHighlightedProperties = this.fixArrProp(queryPayload.HitHighlightedProperties);
        queryPayload.Properties = this.fixArrProp(queryPayload.Properties);
        queryPayload.RefinementFilters = this.fixArrProp(queryPayload.RefinementFilters);
        queryPayload.ReorderingRules = this.fixArrProp(queryPayload.ReorderingRules);
        queryPayload.SelectProperties = this.fixArrProp(queryPayload.SelectProperties);
        queryPayload.SortList = this.fixArrProp(queryPayload.SortList);

        const postBody = {
            request: {
                '__metadata': {
                    'type': 'Microsoft.Office.Server.Search.REST.SearchRequest'
                },
                ...queryPayload
            }
        };

        return JSON.stringify(postBody);
    }

    /**
     * Fix array property
     *
     * @param prop property to fix for container struct
     */
    private fixArrProp(prop: any): { results: any[] } {
        if (typeof prop === "undefined") {
            return ({ results: [] });
        }
        return { results: Array.isArray(prop) ? prop : [prop] };
    }
    
}