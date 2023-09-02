import { IColumn } from "office-ui-fabric-react";

export interface ISearchResultsState {
    columns: IColumn[];
    orderByColumn: string;
    orderDirection: string;
    isPanelOpen: boolean;
    selectedDocumentName: string;
    selectedDocumentUrl: string;
    fileContent: string;
    isLoading: boolean;
    searchQuery: string;
    searchResults: any;
}