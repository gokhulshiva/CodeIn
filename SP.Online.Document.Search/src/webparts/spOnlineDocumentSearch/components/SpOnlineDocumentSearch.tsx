import * as React from 'react';
import styles from './SpOnlineDocumentSearch.module.scss';
import { ISpOnlineDocumentSearchProps } from './ISpOnlineDocumentSearchProps';
import { ISpOnlineDocumentSearchState } from './ISpOnlineDocumentSearchState';
import { escape } from '@microsoft/sp-lodash-subset';
import SearchControls from './searchControls/SearchControls';
import { ISharePointSearchResults } from '../../../models/search/ISharePointSearchResults';
import SearchResults from './searchResults/SearchResults';
import { ISearchResult } from '../../../models/ISearchResult';
import { ISearchResults } from '../services/ISearchService';
import SearchService from '../services/SearchService';
import LibrarySearch from './librarySearch/LibrarySearch';

export default class SpOnlineDocumentSearch extends React.Component<ISpOnlineDocumentSearchProps, ISpOnlineDocumentSearchState> {
  private _searchService: SearchService;
  constructor(props: ISpOnlineDocumentSearchProps) {
    super(props);
    this._searchService = new SearchService(this.props.context);
    this.state = {
      isDataLoaded: false,
      inProgress: false,
      searchResults: undefined
    }
    this.updateSearchContext = this.updateSearchContext.bind(this);
    this.updateSearchContex1 = this.updateSearchContex1.bind(this);
    this.handleBasicSearch = this.handleBasicSearch.bind(this);
  }

  private initiateSearch() {
    this.setState({
      inProgress: true
    });
  }

  private updateSearchContext(searchResults: ISharePointSearchResults) {
    console.log('ISharePointSearchResults');
    console.log(searchResults);
    // this.setState({
    //   searchResults
    // })
  }

  private updateSearchContex1(searchResults: ISearchResult[]) {
    console.log('ISharePointSearchResults');
    console.log(searchResults);
    this.setState({
      searchResults,
      isDataLoaded: true,
      inProgress: false
    });
  }

  private async handleBasicSearch(query: string): Promise<void> {
    this.setState({
      inProgress: true
    });
    const results = await this._searchService.getSearchResults1(query);
    console.log('results', results);
    this.setState({
      searchResults: results,
      inProgress: false
    });
  }

  public render(): React.ReactElement<ISpOnlineDocumentSearchProps> {
    const {
      description,
    } = this.props;

    return (
      <section className={`${styles.spOnlineDocumentSearch}}`}>
        <div className='spOnlineDocumentSearchContainer'> 
        <LibrarySearch  
          Context={this.props.context}
          SiteUrl={this.props.siteUrl}/>
          {/* <SearchControls
            Context={this.props.context}
            SiteUrl={this.props.siteUrl}
            InitiateSearch={this.initiateSearch}
            UpdateContext={this.updateSearchContext}
            UpdateContext1={this.updateSearchContex1}
            OnBasicSearch={this.handleBasicSearch} />
          <SearchResults
            InProgress={this.state.inProgress}
            SearchResults={this.state.searchResults} /> */}
        </div>
      </section>
    );
  }
}
