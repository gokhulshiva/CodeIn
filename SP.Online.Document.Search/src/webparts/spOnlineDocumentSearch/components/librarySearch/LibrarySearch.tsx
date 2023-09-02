import * as React from "react";
import { ILibrarySearchProps } from "./ILibrarySearchProps";
import { ILibrarySearchState } from "./ILibrarySearchState";
import {
    ChoiceGroup,
    DatePicker,
    Dropdown,
    IChoiceGroupOption,
    IDropdownOption,
    Label,
    PrimaryButton,
    Stack,
    StackItem,
    TextField,
    mergeStyleSets,
    IDatePickerStrings
} from "office-ui-fabric-react";
import { stackStyles } from "../../../../common/fabricStyles";
import { DateFilterType, DATE_FILTER_TYPES, DATE_FILTER_PERIODS } from "../../../../common/WPConstants";
import { SPFx, spfi } from "@pnp/sp";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';

const styles = mergeStyleSets({
    root: { selectors: { '> *': { marginBottom: 15 } } },
    control: { maxWidth: 300, marginBottom: 15 },
});

export default class LibrarySearch extends React.Component<ILibrarySearchProps, ILibrarySearchState> {
    private datePickerRefFrom: any;
    private datePickerRefTo: any;
    private _sp;
    constructor(props: ILibrarySearchProps) {
        super(props);
        this._sp = spfi().using(SPFx(this.props.Context));
        this.state = {
            dateFilterType: DateFilterType.Period,
            selectedDatePeriod: "Last 1 month",
            fromDate: undefined,
            toDate: undefined
        }
    }

    private handleDateFilterTypeChange(ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption) {
        const dateFilterType = option.text;
        this.setState({
            dateFilterType
        });
    }

    handleDateFilterPeriodChange(ev: React.SyntheticEvent<HTMLElement>, option: IDropdownOption) {
        const selectedDatePeriod = option.text as string;
        this.setState({
            selectedDatePeriod
        });
    }

    private handleDocumentSearch() {
        const libraryName = 'Proposal Folder';
        const { fromDate, toDate } = this.state;
        const fromDateString = fromDate.toISOString();
        const _fromDate = `${fromDateString.substring(0, fromDateString.indexOf('T'))}T23:59:59Z`;
        const toDateString = fromDate.toISOString();
        const _toDate = `${toDateString.substring(0, toDateString.indexOf('T'))}T23:59:59Z`;
        const filter = `PublishedDate ge datetime'${_fromDate}' and PublishedDate le datetime'${_toDate}'`;
        const requestUri = `${this.props.SiteUrl}/_api/web/lists/getbytitle('${libraryName}')/items?$filter=${filter}`;
        console.log('uri', requestUri);
        this.props.Context.spHttpClient.get(requestUri, SPHttpClient.configurations.v1).then((res: SPHttpClientResponse) => {
            return res.json();
        }).then((res: any) => {
            console.log('data', res);
        }).catch(error => {
            return Promise.reject(JSON.stringify(error));
        });

        // const searchResults = this._sp.web.lists.getByTitle(libraryName).items
        //     //.select(select)
        //     //.expand(expand)
        //     //.orderBy('DisplayOrder', true)()
        //     .filter(`datetime'${_fromDate}' ge PublishedDate and datetime'${_toDate}' le PublishedDate`)
        //     .then((data) => {
        //         console.log('data', data);
        //         return data;
        //     })
        //     .catch((error) => {
        //         console.log(error);
        //         return [];
        //     });
    }

    private handleFromDateChange(date: Date) {
        this.setState({
            fromDate: date
        });
    }

    private handleToDateChange(date: Date) {
        this.setState({
            toDate: date
        });
    }

    private onFormatDate = (date?: Date): string => {
        return !date ? '' : date.getDate() + '/' + (date.getMonth() + 1) + '/' + (date.getFullYear() % 100);
    };

    private getFormattedDate(oldValue: Date, newValue: string): Date {
        const previousValue = oldValue || new Date();
        const newValueParts = (newValue || '').trim().split('/');
        const day =
            newValueParts.length > 0 ? Math.max(1, Math.min(31, parseInt(newValueParts[0], 10))) : previousValue.getDate();
        const month =
            newValueParts.length > 1
                ? Math.max(1, Math.min(12, parseInt(newValueParts[1], 10))) - 1
                : previousValue.getMonth();
        let year = newValueParts.length > 2 ? parseInt(newValueParts[2], 10) : previousValue.getFullYear();
        if (year < 100) {
            year += previousValue.getFullYear() - (previousValue.getFullYear() % 100);
        }
        return new Date(year, month, day);
    }

    public render(): React.ReactElement<ILibrarySearchProps> {
        const dateTypeOptions = DATE_FILTER_TYPES.map((t) => ({ key: t, text: t }));
        const periodOptions = DATE_FILTER_PERIODS.map((p) => ({ key: p, text: p }));
        const { dateFilterType, selectedDatePeriod, fromDate, toDate } = this.state;
        return (
            <div className="librarySearchContainer">
                <Stack styles={stackStyles}>

                </Stack>
                <Stack styles={stackStyles} horizontal className="width100">
                    <StackItem className="col10">

                    </StackItem>
                    <StackItem className="col50">
                        <Stack styles={stackStyles} horizontal style={{ justifyContent: 'space-between' }}>
                            <StackItem>
                                <Label style={{ marginTop: "10px" }}>
                                    Date Created
                                </Label>
                            </StackItem>
                            <StackItem>
                                <Stack>
                                    <StackItem>
                                        <ChoiceGroup
                                            options={dateTypeOptions}
                                            selectedKey={dateFilterType}
                                            className="searchType"
                                            onChange={(e, o) => this.handleDateFilterTypeChange(e, o)} />
                                    </StackItem>
                                </Stack>
                            </StackItem>
                        </Stack>
                        <Stack styles={stackStyles}>
                            {
                                dateFilterType == DateFilterType.Period ?
                                    <StackItem>
                                        <Dropdown
                                            options={periodOptions}
                                            selectedKey={selectedDatePeriod}
                                            onChange={(e, o) => this.handleDateFilterPeriodChange(e, o)} />
                                    </StackItem>

                                    :

                                    <StackItem>
                                        <Stack styles={stackStyles} horizontal>
                                            <Label>
                                                From
                                            </Label>
                                            <DatePicker
                                                componentRef={this.datePickerRefFrom}
                                                //label="Start date"
                                                allowTextInput
                                                ariaLabel="Select a date. Input format is day slash month slash year."
                                                value={fromDate}
                                                onSelectDate={(date?: Date) => this.handleFromDateChange(date)}
                                                formatDate={this.onFormatDate}
                                                //parseDateFromString={onParseDateFromString}
                                                className={styles.control}
                                            // DatePicker uses English strings by default. For localized apps, you must override this prop.
                                            />
                                            <Label>
                                                To
                                            </Label>
                                            <DatePicker
                                                componentRef={this.datePickerRefFrom}
                                                //label="Start date"
                                                allowTextInput
                                                ariaLabel="Select a date. Input format is day slash month slash year."
                                                value={toDate}
                                                onSelectDate={(date?: Date) => this.handleToDateChange(date)}
                                                formatDate={this.onFormatDate}
                                                //parseDateFromString={onParseDateFromString}
                                                className={styles.control}
                                            // DatePicker uses English strings by default. For localized apps, you must override this prop.
                                            />
                                        </Stack>
                                    </StackItem>
                            }

                        </Stack>
                    </StackItem>
                </Stack>
                <Stack>
                    <StackItem className="width100 btnRow" >
                        <PrimaryButton text="Search" onClick={() => this.handleDocumentSearch()} />
                    </StackItem>
                </Stack>
            </div>
        )
    }
}