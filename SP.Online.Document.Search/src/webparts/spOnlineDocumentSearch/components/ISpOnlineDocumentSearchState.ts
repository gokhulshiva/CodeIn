import { ISearchResult } from "../../../models/ISearchResult";
import { ISharePointSearchResult, ISharePointSearchResults } from "../../../models/search/ISharePointSearchResults";
import { ISearchResults } from "../services/ISearchService";

export interface ISpOnlineDocumentSearchState {
    isDataLoaded: boolean;
    inProgress: boolean;
    // searchResults: ISharePointSearchResults;
    searchResults: ISearchResult[];
}