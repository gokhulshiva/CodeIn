import { SearchType } from "../../../../common/WPConstants";

export interface ISearchControlsState {
    searchType: string;
    searchQuery: string;
    allWords: string;
    exactPhrase: string;
    anyWords: string;
}