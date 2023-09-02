import * as React from "react";
import { ISearchControlsProps } from "./ISearchControlsProps";
import { ISearchControlsState } from "./ISearchControlsState";
import { ChoiceGroup, IChoiceGroupOption, PrimaryButton, SearchBox, Stack, StackItem, TextField } from "office-ui-fabric-react";
import { stackStyles } from "../../../../common/fabricStyles";
import { BuiltinSourceIds, SearchType, SearchTypes } from "../../../../common/WPConstants";
import SearchService from "../../services/SearchService";
import { IDataContext } from "@pnp/modern-search-extensibility";
import { ISharePointSearchQuery } from "../../../../models/search/ISharePointSearchQuery";
import { SharePointSearchService } from "../../../../services/searchService/SharePointSearchService";

export default class SearchControls extends React.Component<ISearchControlsProps, ISearchControlsState> {
    private _searchService: SearchService;
    private _sharePointSearchService: SharePointSearchService;

    constructor(props: ISearchControlsProps) {
        super(props);
        this._searchService = new SearchService(this.props.Context);
        //this._sharePointSearchService = serviceScope.consume<ISharePointSearchService>(SharePointSearchService.ServiceKey);
        this.state = {
            searchType: SearchType.Basic,
            searchQuery: "",
            allWords: "",
            exactPhrase: "",
            anyWords: ""
        }
    }

    private handleSearchTypeChange(ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption) {
        const searchType = option.text;
        this.setState({
            searchType
        });
    }

    private handleAllWordsChange(e: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string) {
        this.setState({
            allWords: newText
        });
    }

    private handleExactPhraseChange(e: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string) {
        this.setState({
            exactPhrase: newText
        });
    }

    private handleAnyWordsChange(e: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string) {
        this.setState({
            anyWords: newText
        });
    }

    private handleSearch(searchQuery: string) {
        this.setState({
            searchQuery
        });
    }

    private handleAdvanceSearch() {

    }

    private clearSearch() {
        this.setState({
            searchQuery: ""
        });
    }

    private async searchQuery(query: string): Promise<void> {
        const dataContext: IDataContext = { inputQueryText: query, itemsCountPerPage: 10, originalInputQueryText: query, pageNumber: 1 };
        const searchQuery: ISharePointSearchQuery = { Querytext: query };
        searchQuery.TrimDuplicates = true;
        searchQuery.SourceId = BuiltinSourceIds.Documents;
        //searchQuery.RefinementFilters = [`Path:equals("${this.props.SiteUrl}")`]
        //const results = await this._searchService.search(searchQuery);
        // const results = await this._searchService.getSearchResults1(query);
        // console.log('results', results);
        this.props.OnBasicSearch(query);
    }



    public render(): React.ReactElement<ISearchControlsProps> {
        const optionsSearchType: IChoiceGroupOption[] = SearchTypes.map((t) => ({ key: t, text: t }));
        const searchType = this.state.searchType;
        const { allWords, exactPhrase, anyWords } = this.state;
        return (
            <div className="searchControlsContainer">
                <Stack styles={stackStyles}>
                    <StackItem>
                        <ChoiceGroup
                            options={optionsSearchType}
                            selectedKey={searchType}
                            className="searchType"
                            onChange={(e, o) => this.handleSearchTypeChange(e, o)} />
                    </StackItem>
                </Stack>
                {
                    searchType == SearchType.Basic ?

                        <div className="basicSearch">
                            <Stack styles={stackStyles} >
                                <StackItem className="col50">
                                    <SearchBox placeholder="Search"
                                        value={this.state.searchQuery}
                                        onChange={(e, value) => this.handleSearch(value)}
                                        onSearch={(value) => this.searchQuery(value)}
                                        onClear={() => this.clearSearch()} />
                                </StackItem>
                            </Stack>
                        </div>

                        :

                        <div className="advancedSearch">
                            <Stack styles={stackStyles} horizontal className="width100">
                                <StackItem className="col10">
                                    Find documents that have
                                </StackItem>
                                <StackItem className="col50">
                                    <Stack>
                                        <StackItem>
                                            <Stack>
                                                All of these words:
                                                <TextField value={allWords} onChange={(e, value) => this.handleAllWordsChange(e, value)} />
                                            </Stack>
                                            <Stack>
                                                The exact phrase:
                                                <TextField value={allWords} onChange={(e, value) => this.handleExactPhraseChange(e, value)} />
                                            </Stack>
                                            <Stack>
                                                Any of these words:
                                                <TextField value={allWords} onChange={(e, value) => this.handleAnyWordsChange(e, value)} />
                                            </Stack>
                                        </StackItem>
                                    </Stack>
                                </StackItem>
                            </Stack>
                            <Stack>
                                <StackItem className="width100 btnRow" >
                                    <PrimaryButton text="Search" onClick={() => this.handleAdvanceSearch()} />
                                </StackItem>
                            </Stack>
                        </div>
                }


            </div>
        )
    }
}