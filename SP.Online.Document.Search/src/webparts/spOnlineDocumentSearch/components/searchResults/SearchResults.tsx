import * as React from "react";
import { ISearchResultsProps } from "./ISearchResultsProps";
import { ISearchResultsState } from "./ISearchResultsState";
import { FocusZone, FocusZoneDirection, Icon, Link, List, Spinner, SpinnerSize, Stack, StackItem } from "office-ui-fabric-react";
import { ISearchResult } from "../../../../models/ISearchResult";
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { stackStyles } from "../../../../common/fabricStyles";
import { IconPicker } from "@pnp/spfx-controls-react";


export default class SearchResults extends React.Component<ISearchResultsProps, ISearchResultsState> {
    
    constructor(props: ISearchResultsProps) {
        super(props);
    }

    private _onRenderCell(item: ISearchResult, index: number, isScrolling: boolean): JSX.Element {
        console.log('doc type', item.FileExtension);
        const extn = item.FileExtension;
        const iconUrl = item.FileExtension == 'url' ? 
                        `https://modernb.akamai.odsp.cdn.office.net/files/fabric-cdn-prod_20210703.001/assets/item-types/16/link.svg` :
                        `https://modernb.akamai.odsp.cdn.office.net/files/fabric-cdn-prod_20210703.001/assets/item-types/16/${extn}.svg`;

        return (
            <div className="ms-ListGhostingExample-itemCell searchResultItem" data-is-focusable={true}>
                <div className="ms-ListGhostingExample-itemContent">
                    <div>
                        <Stack styles={stackStyles} horizontal>
                            <StackItem className="">
                                <FileTypeIcon 
                                    type={IconType.image}
                                    path={item.Path} />
                            </StackItem>
                            <StackItem>
                                <div className="ms-ListGhostingExample-itemName searchResultLink">
                                    <Link target="_blank" data-interception="off" href={item.ServerRedirectedURL} >{item.Title}</Link>
                                </div>
                                <div className="ms-ListGhostingExample-itemName searchResultSummary"
                                    dangerouslySetInnerHTML={{ __html: item.HitHighlightedSummary }}>
                                </div>
                            </StackItem>
                        </Stack>
                    </div>
                    <p></p>
                    {/* <div className="ms-ListGhostingExample-itemIndex">{`Item ${index}`}</div> */}
                </div>
            </div>
        );
    }


    public render(): React.ReactElement<ISearchResultsProps> {
        const searchResults = this.props.SearchResults;
        return (
            <div className="searchResultsContainer">
                <FocusZone direction={FocusZoneDirection.vertical}>
                    <div className="ms-ListGhostingExample-container" data-is-scrollable={true}>
                        {
                            this.props.InProgress ?

                                <Spinner size={SpinnerSize.medium} />

                                :

                                <List items={searchResults} onRenderCell={this._onRenderCell} />

                        }
                    </div>
                </FocusZone>
            </div>
        )
    }
}