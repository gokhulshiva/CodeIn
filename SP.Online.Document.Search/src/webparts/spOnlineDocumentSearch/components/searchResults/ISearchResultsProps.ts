import { ISearchResult } from "../../../../models/ISearchResult";
import { ISharePointSearchResult, ISharePointSearchResults } from "../../../../models/search/ISharePointSearchResults";

export interface ISearchResultsProps {
    InProgress: boolean;
    SearchResults: ISearchResult[];
}