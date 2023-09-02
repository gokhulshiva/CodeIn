import { ISearchResult } from "../../../../models/ISearchResult";
import { ISharePointSearchResults } from "../../../../models/search/ISharePointSearchResults";
import { SPHttpClient } from "@microsoft/sp-http";
import { ISearchResults } from "../../services/ISearchService";

export interface ISearchControlsProps {
    Context: any;
    SiteUrl: string;
    InitiateSearch: () => void;
    UpdateContext: (searchResults: ISharePointSearchResults) => void;
    UpdateContext1: (searchResults: ISearchResult[]) => void;
    OnBasicSearch: (query: string) => void;
}