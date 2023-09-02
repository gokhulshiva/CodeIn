import * as React from 'react';
import styles from './SpOnlineLibrarySearch.module.scss';
import { ISpOnlineLibrarySearchProps } from './ISpOnlineLibrarySearchProps';
import { ISpOnlineLibrarySearchState } from './ISpOnlineLibrarySearchState';
import { escape } from '@microsoft/sp-lodash-subset';
import LibrarySearch from './librarySearch/LibrarySearch';
import { SPService } from '../services/SPService';
import { loadTheme } from 'office-ui-fabric-react';

export default class SpOnlineLibrarySearch extends React.Component<ISpOnlineLibrarySearchProps, ISpOnlineLibrarySearchState> {
  private _spService: SPService;

  constructor(props: ISpOnlineLibrarySearchProps) {
    super(props);
    loadTheme({
      palette: {
        themePrimary: '#5d21a6',
        themeLighterAlt: '#f7f3fb',
        themeLighter: '#e0d2f1',
        themeLight: '#c6ade4',
        themeTertiary: '#9469c9',
        themeSecondary: '#6c34b0',
        themeDarkAlt: '#541e95',
        themeDark: '#47197e',
        themeDarker: '#34135d',
        neutralLighterAlt: '#faf9f8',
        neutralLighter: '#f3f2f1',
        neutralLight: '#edebe9',
        neutralQuaternaryAlt: '#e1dfdd',
        neutralQuaternary: '#d0d0d0',
        neutralTertiaryAlt: '#c8c6c4',
        neutralTertiary: '#a19f9d',
        neutralSecondary: '#605e5c',
        neutralPrimaryAlt: '#3b3a39',
        neutralPrimary: '#323130',
        neutralDark: '#201f1e',
        black: '#000000',
        white: '#ffffff',
      }});
    this._spService = new SPService(this.props.context);
    this.state = {
    }
  }

  componentDidMount(): void {
  }


  

  public render(): React.ReactElement<ISpOnlineLibrarySearchProps> {
    const {
      description,
    } = this.props;

    return (
      <section className={`${styles.spOnlineLibrarySearch} spOnlineDocumentSearchContainer`}>
        <LibrarySearch 
          Context={this.props.context} 
          SiteUrl={this.props.siteUrl}
           />
      </section>
    );
  }
}
