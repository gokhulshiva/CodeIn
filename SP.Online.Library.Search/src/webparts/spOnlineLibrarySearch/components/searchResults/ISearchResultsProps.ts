import { IProposalItem } from "../../models/IProposalItem";

export interface ISearchResultsProps {
    Context: any;
    IsLoading: boolean;
    SearchPanelText: string;
    SearchResults: IProposalItem[];
    OnSearchTextChange: (searchText: string) => void;
}