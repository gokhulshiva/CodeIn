import { IFolderInfo } from "@pnp/sp/folders";
import { IDateConfigItem } from "../../models/IDateConfigItem";
import { IFileItem } from "../../models/IFileItem";
import { IProposalItem } from "../../models/IProposalItem";
import { ISearchResult } from "../../models/ISearchResult";

export interface ILibrarySearchState {
    dateFilterType: string;
    selectedDatePeriod: number;
    selectedAgencies: string[];
    selectedCategory: string;
    selectedSubjectDomains: string[];
    fromDate: Date;
    toDate: Date;
    dateFilterItems: IDateConfigItem[];
    documents: IFileItem[];
    allLibraryResults: IProposalItem[];
    libraryFilterResults: IProposalItem[];
    inProgress: boolean;
    termSetIdAgency: string;
    termSetIdCategory: string;
    agencies: IFolderInfo[];
    categories: string[];
    searchQuery: string;
    searchPanelText: string;
    querySearchResults: ISearchResult[];
    searchResults: IProposalItem[];
    fileContent: string;
}