import * as React from "react";
import { ISearchResultsProps } from "./ISearchResultsProps";
import { ISearchResultsState } from "./ISearchResultsState";
import { DetailsList, DetailsListLayoutMode, IButtonStyles, IColumn, IIconProps, IconButton, Link, Panel, PanelType, SearchBox, SelectionMode, Spinner, SpinnerSize, getTheme } from "office-ui-fabric-react";
import { IProposalItem } from "../../models/IProposalItem";
import { ISearchResponse } from "../../services/ISearchService";
import { ISearchResult } from "../../models/ISearchResult";
import { SPService } from "../../services/SPService";
import { FILETYPE, FieldType } from "../../common/constants";
import * as $ from 'jquery';


const viewIcon: IIconProps = { iconName: 'View' };
const theme = getTheme();
const iconButtonStyles: Partial<IButtonStyles> = {
    root: {
        color: theme.palette.neutralPrimary,
        marginLeft: 'auto',
        marginTop: '4px',
        marginRight: '2px',
    },
    rootHovered: {
        color: theme.palette.neutralDark,
    },
};

export default class SearchResults extends React.Component<ISearchResultsProps, ISearchResultsState> {
    private _spService: SPService;
    constructor(props: ISearchResultsProps) {
        super(props);
        this._spService = new SPService(this.props.Context)
        this.state = {
            columns: [],
            orderByColumn: 'ID',
            orderDirection: 'desc',
            isPanelOpen: false,
            selectedDocumentName: "",
            selectedDocumentUrl: "",
            fileContent: "",
            isLoading: false,
            searchQuery: "",
            searchResults: []
        }
    }

    componentDidMount(): void {
        this.generateColumns();
    }

    private async _onColumnClick(ev: React.MouseEvent<HTMLElement>, column: IColumn): Promise<void> {
        const { columns, orderByColumn } = this.state;
        let orderDirection: string;

        const newColumns: IColumn[] = columns.slice();
        const currColumn: IColumn = newColumns.filter(currCol => column.key === currCol.key)[0];
        newColumns.forEach((newCol: IColumn) => {
            if (newCol === currColumn) {
                currColumn.isSortedDescending = !currColumn.isSortedDescending;
                orderDirection = currColumn.isSortedDescending ? 'desc' : 'asc';
                currColumn.isSorted = true;
            } else {
                newCol.isSorted = false;
                newCol.isSortedDescending = true;
            }
        });

        this.setState({
            columns: newColumns,
            orderByColumn: column.fieldName,
            orderDirection,
        }); //currentTabFilter       
    }

    private async viewDocument(item: IProposalItem) {
        this.setState({
            isLoading: true,
            isPanelOpen: true
        });
        const documentUri: URL = new URL(item.ServerRedirectedEmbedUrl);

        if (documentUri.searchParams.has('action')) {
            documentUri.searchParams.delete('action');
        }

        const documentUrl = documentUri.href;
        const documentName = item.FileLeafRef;
        const fileRelativeUrl = item.FileRef;
        const type = item.DocIcon;
        const data = await this._spService.getFileContent(fileRelativeUrl, type);
        this.setState({
            selectedDocumentName: documentName,
            selectedDocumentUrl: documentUrl,
            fileContent: data.value,
            isLoading: false
        }, () => {
            // load search
            const searchText = this.props.SearchPanelText;
            this.handleSearch(searchText);
        });
    }

    private onFrameLoaded(data: any) {
        console.log('frame loaded');
        console.log(data);
    }

    private onFrameLoad(e: any) {
        console.log('e');
        console.log(e);
        const iFrame: any = document.getElementById('#searchDocumentFrame');
        const iFrameWindow = iFrame.contentWindow;
        const iFrameDocument = iFrame.contentDocument;

        console.log("frame window");
        console.log(iFrameWindow);

        console.log("frame document");
        console.log(iFrameWindow);
    }

    private generateColumns() {
        const columns: IColumn[] = [
            {
                key: 'name',
                name: 'Name',
                fieldName: 'FileLeafRef',
                minWidth: 250,
                maxWidth: 350,
                isRowHeader: true,
                isResizable: true,
                isSorted: true,
                //isSortedDescending: this.state.orderDirection === 'desc' ? true : false,
                sortAscendingAriaLabel: 'Sorted A to Z',
                sortDescendingAriaLabel: 'Sorted Z to A',
                onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true,
                onRender: (item: IProposalItem) => {
                    return (
                        <span className="linkContainer">
                            <Link target="_blank" href={item.ServerRedirectedEmbedUrl} >{item.FileLeafRef}</Link>
                        </span>
                    )
                }
            },
            {
                key: 'view',
                name: '',
                minWidth: 50,
                maxWidth: 50,
                isRowHeader: true,
                isResizable: true,
                onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true,
                onRender: (item: IProposalItem) => {
                    return (
                        <span className="linkContainer">
                            <IconButton
                                styles={iconButtonStyles}
                                iconProps={viewIcon}
                                ariaLabel="View document"
                                onClick={() => this.viewDocument(item)}
                            />
                        </span>
                    )
                }
            },
            {
                key: 'agency',
                name: 'Agency', //Type
                fieldName: 'Agency',
                minWidth: 50,
                maxWidth: 100, //185,
                isResizable: true,
                onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true,
                onRender: (item: IProposalItem) => {
                    const path = item.FileDirRef.split('/').pop();
                    return path
                }
            },
            {
                key: 'category',
                name: 'Category', //Type
                fieldName: 'Category',
                minWidth: 50,
                maxWidth: 100, //185,
                isResizable: true,
                onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true,
                onRender: (item: IProposalItem) => {
                    return item.Category_x0020_Type ? item.Category_x0020_Type.Label : "";
                }
            },
            {
                key: 'subjectdomain',
                name: 'Subject Domain', //Type
                fieldName: 'SubjectDomain',
                minWidth: 100,
                maxWidth: 150, //185,
                isResizable: true,
                onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true
            },
            {
                key: 'publisheddate',
                name: 'Published Date', //Type
                fieldName: 'PublishedDate',
                minWidth: 100,
                maxWidth: 150, //185,
                isResizable: true,
                onColumnClick: this._onColumnClick,
                data: 'string',
                isPadded: true
            },
            {
                key: 'highlightedsummary',
                name: 'Summary',
                fieldName: '',
                minWidth: 300,
                maxWidth: 360,
                onRender: (item: IProposalItem) => {
                    const type = typeof (item);
                    return <div className="ms-ListGhostingExample-itemName searchResultSummary"
                        dangerouslySetInnerHTML={{ __html: item.Summary }}>
                    </div>
                }
                //
            }
        ];

        this.setState({
            columns
        });
    }

    private _onDismiss() {
        this.setState({
            isPanelOpen: false
        });
    }

    private handleSearchTextChange(e: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string) {
        // this.setState({
        //     searchQuery: newText
        // });
        this.props.OnSearchTextChange(newText);
    }

    private handleSearch(searchText: string) {
        if (!searchText) {
            return;
        }
        const searchResults: any = [];
        searchText = (searchText || "").toLowerCase();
        const containerElem = document.getElementById('documentContentDiv');
        const elemsArray1 = Array.from(containerElem.getElementsByTagName('p'));
        const matchedElems1 = elemsArray1.filter((elem) => {
            return elem.innerText.toLowerCase().indexOf(searchText) != -1;
        });

        matchedElems1.forEach((elem, i) => {
            const id = `p_searchresult${i}`;
            elem.id = id;
            this.highlightText(elem, searchText);
            const innerText = elem.innerText;
            const index = innerText.indexOf(searchText);
            const summary = innerText.substring(0, 100);
            searchResults.push({ tag: "p", id: id, summary });
        });

        const elemsArray2 = Array.from(containerElem.getElementsByTagName('li'));
        const matchedElems2 = elemsArray2.filter((elem) => {
            return elem.innerText.toLowerCase().indexOf(searchText) != -1;
        });

        matchedElems2.forEach((elem, i) => {
            const id = `li_searchresult${i}`;
            elem.id = id;
            this.highlightText(elem, searchText);
            const innerText = elem.innerText;
            const index = innerText.indexOf(searchText);
            const summary = innerText.substring(0, 100);
            searchResults.push({ tag: "li", id: id, summary });
        });

        console.log('matched elements');
        console.log(searchResults);

        this.setState({
            searchResults
        });
    }

    private highlightText(elem: any, searchText: string) {
        if (searchText !== '') {
            const content = elem.innerHTML;
            const highlightedContent = content.replace(
                new RegExp(searchText, 'gi'),
                '<span class="highlight">$&</span>'
            );
            elem.innerHTML = highlightedContent;
        }
    }

    private removeHighlight() {
        const searchText = this.props.SearchPanelText;
        const searchResults = this.state.searchResults;
        searchResults.forEach((result: any) => {
            const elem = document.getElementById(result.id);
            const content = elem.innerHTML;
            const highlightedPhrase = `<span class="highlight">${searchText}</span>`;
            const highlightedContent = content.replace(
                new RegExp(highlightedPhrase, 'gi'),
                searchText
            );
            elem.innerHTML = highlightedContent;
        });
    }

    private clearSearch() {
        const fileContent = this.state.fileContent;
        const searchText = this.state.searchQuery;
        this.removeHighlight();
        this.setState({
            searchQuery: "",
            searchResults: [],
            //fileContent
        });
        this.props.OnSearchTextChange("");
    }

    private onRenderFooterContent(): React.ReactElement {
        const btnDisabled = false;
        return (
            <div className="panelFooterContainer" style={{height: "40px"}}>
               
            </div>
        )
    }

    public render(): React.ReactElement<ISearchResultsProps> {
        const { SearchResults: results } = this.props;
        const columns = this.state.columns;

        return (
            <div className="resultsContainer">

                <DetailsList
                    items={results}
                    compact={false}
                    columns={columns}
                    selectionMode={SelectionMode.none}
                    setKey="set"
                    layoutMode={DetailsListLayoutMode.justified}
                    isHeaderVisible={true}
                    selectionPreservedOnEmptyClick={false}
                    enterModalSelectionOnTouch={true}
                    ariaLabelForSelectionColumn="Toggle selection"
                    ariaLabelForSelectAllCheckbox="Toggle selection for all items"
                />
                {
                    <div className="documentPreviewContainer" style={{ display: this.state.isPanelOpen ? "block" : "none" }}>
                        <Panel
                            type={PanelType.large}
                            isOpen={this.state.isPanelOpen}
                            onDismiss={this._onDismiss.bind(this)}
                            headerText={this.state.selectedDocumentName}
                            closeButtonAriaLabel="Close"
                            onRenderFooterContent={this.onRenderFooterContent}
                            isFooterAtBottom={true} >
                            <div className="panelContent">
                                <div className="documentPreview">
                                    {/* <iframe 
                                        id="searchDocumentFrame" 
                                        src={`${this.state.selectedDocumentUrl}`}
                                        onLoad={(e) => this.onFrameLoad(e)}
                                        onLoadedData={this.onFrameLoaded} /> */}
                                    {
                                        this.state.isLoading ?

                                            <div className="spiinerContainer">
                                                <Spinner size={SpinnerSize.medium} />
                                            </div>

                                            :
                                            <div className="previewContainer">
                                                <div className="searchPanel">
                                                    <div className="searchPanelContainer">
                                                        <div className="searchBoxContainer">
                                                            <SearchBox placeholder="Search"
                                                                value={this.props.SearchPanelText}
                                                                onChange={(e, value) => this.handleSearchTextChange(e, value)}
                                                                onSearch={(value) => this.handleSearch(value)}
                                                                onClear={() => this.clearSearch()} />
                                                        </div>
                                                        <div className="searchResultElements">
                                                            <ul>
                                                                {
                                                                    this.state.searchResults.map((result: any) => {
                                                                        return (
                                                                            <li>
                                                                                <Link href={`#${result.id}`} >
                                                                                    {result.summary}
                                                                                </Link>
                                                                            </li>
                                                                        )
                                                                    })
                                                                }
                                                            </ul>
                                                        </div>
                                                    </div>
                                                </div>
                                                <div className="documentContainer">
                                                    <div
                                                        className="documentContent"
                                                        id="documentContentDiv"
                                                        dangerouslySetInnerHTML={{ __html: this.state.fileContent }}>
                                                    </div>
                                                </div>
                                            </div>
                                    }
                                </div>
                            </div>
                        </Panel>
                    </div>
                }
            </div>
        )
    }
}