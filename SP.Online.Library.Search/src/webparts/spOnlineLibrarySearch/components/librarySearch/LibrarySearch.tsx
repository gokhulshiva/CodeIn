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
    IDatePickerStrings,
    Link,
    DefaultButton,
    FocusZone,
    FocusZoneDirection,
    Spinner,
    SpinnerSize,
    List,
    SearchBox,
    Panel,
    PanelType,
} from "office-ui-fabric-react";
import { FileTypeIcon, ApplicationType, IconType, ImageSize } from "@pnp/spfx-controls-react/lib/FileTypeIcon";
import { SPFx, spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IRenderListDataParameters } from "@pnp/sp/lists";
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { Condition, DATE_FILTER_PERIODS, DATE_FILTER_TYPES, DateFilterType, FieldType, ItemType, LIBRARY_PROPOSAL, Operator, SUBJECT_DOMAIN } from "../../common/constants";
import { columnStackStyles, dateColumnStackStyles, headerRowStyles, stackStyleBtnRow, stackStyles } from "../../common/fabricStyles";
import { IFileItem } from "../../models/IFileItem";
import { LibsOrderBy } from "@pnp/spfx-controls-react/lib/services/ISPService";
import { IProposalItem } from "../../models/IProposalItem";
import SearchResults from "../searchResults/SearchResults";
import { SPService } from "../../services/SPService";
import { MergeCAMLConditions, MergeType } from "../../helpers/caml";
import { IMetaFieldInfo } from "../../models/IMetaFieldInfo";
import { IPickerTerms, TaxonomyPicker } from "@pnp/spfx-controls-react";
import {
    TermStore,
    ITermStore,
    ITermSet,
    ITerms,
    ITerm,
} from "@pnp/sp/taxonomy";
import { ITaxonomyGroup } from "../../models/ITaxonomyGroup";
import { IFolder, IFolderInfo } from "@pnp/sp/folders";
import * as mammoth from 'mammoth';

const styles = mergeStyleSets({
    root: { selectors: { '> *': { marginBottom: 15 } } },
    control: { maxWidth: 300, marginBottom: 15 },
});

export default class LibrarySearch extends React.Component<ILibrarySearchProps, ILibrarySearchState> {
    private datePickerRefFrom: any;
    private datePickerRefTo: any;
    private _sp;
    private _spService: SPService;
    constructor(props: ILibrarySearchProps) {
        super(props);
        this._sp = spfi().using(SPFx(this.props.Context));
        this._spService = new SPService(this.props.Context);
        this.state = {
            dateFilterType: DateFilterType.Period,
            selectedDatePeriod: 0,
            selectedAgencies: ["All"],
            selectedCategory: "All",
            selectedSubjectDomains: [],
            fromDate: undefined,
            toDate: undefined,
            allLibraryResults: [],
            dateFilterItems: [],
            documents: [],
            libraryFilterResults: [],
            inProgress: false,
            termSetIdAgency: "",
            termSetIdCategory: "",
            agencies: [],
            categories: [],
            searchQuery: "",
            searchPanelText: "",
            querySearchResults: [],
            searchResults: [],
            fileContent: ""
        }

        this.handleSearchPanelTextChange = this.handleSearchPanelTextChange.bind(this);
    }

    componentDidMount(): void {
       // this.getFileContent();
        this.loadData();
    }

    private async loadData(): Promise<void> {

        const folders: IFolderInfo[] = await this._spService.getAgencyFolders();
        const agencies = folders.filter((f) => f.Name != 'Forms');
        const categories = await this._spService.getCategories();
        const dateFilterItems = await this._spService.getDateFilterConfig();
        await this.getAllDocuments();
        this.setState({
            agencies,
            categories,
            dateFilterItems,
        });

        // const metaInfo1: IMetaFieldInfo = await this._spService.getTermSetIdByFieldName('Agency');
        // const metaInfo2: IMetaFieldInfo = await this._spService.getTermSetIdByFieldName('Category Type');

        // const agencyTerms = await this._spService.getTermSetById(metaInfo1.TermSetId);
        // const categoryTerms = await this._spService.getTermSetById(metaInfo1.TermSetId);
        // // const groups: ITaxonomyGroup[] = await this._spService.getTermGroups();
        // // const termStores: any = await this._spService.getTermStores();

        // // console.log('info1', metaInfo1.TermSetId);
        // // console.log('info2', metaInfo2.TermSetId);
        // console.log('terms1', agencyTerms);
        // console.log('terms2', categoryTerms);

        // this.setState({
        //     termSetIdAgency: metaInfo1.TermSetId || "",
        //     termSetIdCategory: metaInfo2.TermSetId || ""
        // });


        // const dateFilterItems = await this._spService.getDateFilterConfig();
        // this.setState({
        //     dateFilterItems
        // })
    }

    private handleDateFilterTypeChange(ev: React.SyntheticEvent<HTMLElement>, option: IChoiceGroupOption) {
        const dateFilterType = option.text;
        this.setState({
            dateFilterType,
            selectedDatePeriod: 0,
            fromDate: undefined,
            toDate: undefined
        });
    }

    private handleDateFilterPeriodChange(ev: React.SyntheticEvent<HTMLElement>, option: IDropdownOption) {
        const selectedDatePeriod = option.key as number;
        this.setState({
            selectedDatePeriod
        });
    }

    private handleAgencyChange(ev: React.SyntheticEvent<HTMLElement>, option: IDropdownOption) {
        const selectedKeys = this.state.selectedAgencies;
        let selectedAgencies: string[] = [];
        if (option.selected) {
            selectedAgencies = option.key == "All" ? [option.key] : [...selectedKeys.filter(key => key !== "All"), option.key as string];
        } else {
            selectedKeys.filter(key => key !== option.key)
        }
        this.setState({
            selectedAgencies
        });
    }

    private handleCategoryChange(ev: React.SyntheticEvent<HTMLElement>, option: IDropdownOption) {
        this.setState({
            selectedCategory: option.key as string
        });
    }

    private handleDomainChange(ev: React.SyntheticEvent<HTMLElement>, option: IDropdownOption) {
        const selectedKeys = this.state.selectedSubjectDomains;
        const selectedSubjectDomains = option.selected ? [...selectedKeys, option.key as string] : selectedKeys.filter(key => key !== option.key)
        this.setState({
            selectedSubjectDomains
        });
    }

    private handleDocumentSearch1() {
        const libraryName = 'Proposal Folder';
        const searchQuery = this.getSearchQuery();
        const requestUri = `${this.props.SiteUrl}/_api/web/lists/getbytitle('${libraryName}')/GetItems`;
        //const requestUri = `${this.props.SiteUrl}/_api/web/lists/getbytitle('${libraryName}')/items?$select=*&$expand=File/ListItemAllFields,File/ListItemAllFields/FieldValuesAsText&$filter=${filter}`;
        console.log('uri', requestUri);
        let headers: any = {
            'Accept': 'application/json;odata.metadata=full',
            'Content-type': 'application/json;odata.metadata=full',
        };

        headers = {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': '',
            'X-HTTP-Method': 'POST'
        };
        const body = {
            'query': {
                '__metadata': { 'type': 'SP.CamlQuery' },
                'ViewXml': searchQuery
            }
        }
        console.log('body');
        console.log(body);

        this.props.Context.spHttpClient.post(
            requestUri,
            SPHttpClient.configurations.v1,
            {
                headers: headers,
                body: body
            }
        ).then((res: SPHttpClientResponse) => {
            return res.json();
        }).then((res: IFileItem[]) => {
            console.log('data', res);
        }).catch((error: any) => {
            return Promise.reject(JSON.stringify(error));
        });

        const proposalLibrary = this._sp.web.lists.getByTitle(libraryName);
        const getDocuments = proposalLibrary.renderListDataAsStream({
            ViewXml: searchQuery
        });

        getDocuments.then((data) => {
            const searchResults: IProposalItem[] = data.Row;
            this.setState({
                searchResults
            });
            console.log('search results');
            console.log(data);
        }).catch((error) => {
            console.log('error');
            console.log(error);
        })

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

    private handleSearch() {
        const searchQuery = this.state.searchQuery;
        if(searchQuery) {
            this.searchQuery(searchQuery);
        }
        else {
            this.handleDocumentSearch();
        }

        this.setState({
            searchPanelText: searchQuery
        });
    }

    private handleDocumentSearch() {
        this.setState({
            inProgress: true
        });
        const libraryName = 'Proposal Folder';
        const searchQuery = this.getFilterQuery();
        const proposalLibrary = this._sp.web.lists.getByTitle(libraryName);
        const getDocuments = proposalLibrary.renderListDataAsStream({
            ViewXml: searchQuery
        });

        console.log('searchquery', searchQuery);
        getDocuments.then((data) => {
            let searchResults: IProposalItem[] = data.Row;
            searchResults = searchResults.filter((r) => r.FSObjType == ItemType.File);
            this.setState({
                searchResults,
                inProgress: false
            });
            console.log('search results');
            console.log(data);
        }).catch((error) => {
            console.log('error');
            console.log(error);
        });

    }

    private async getAllDocuments() {
        const viewXml = `<View Scope="RecursiveAll">
                            <Query>
                            </Query>
                         </View>`;

        const proposalLibrary = this._sp.web.lists.getByTitle(LIBRARY_PROPOSAL);
        const getDocuments = proposalLibrary.renderListDataAsStream({
            ViewXml: viewXml
        });

        getDocuments.then((data) => {
            let allDocuments: IProposalItem[] = data.Row;
            allDocuments = allDocuments.filter((r) => r.FSObjType == ItemType.File);
            this.setState({
                allLibraryResults: allDocuments,
                inProgress: false
            });
            console.log('search results');
            console.log(allDocuments);
        }).catch((error) => {
            console.log('error');
            console.log(error);
        })
    }

    private async getFileContent() {
         //
         const fileRelativeUrl = "/sites/PerformanceManagementDev/Proposal Folder/Sample/DIEZ-Finance-1.docx";
         let fileContent = await this._spService.getFileContent(fileRelativeUrl, undefined).then((data) => {
            return data;
         }).catch((error) => {
            console.log('parse error');
            console.log(error);
            return error;
         });
         console.log('filecontent');
         console.log(fileContent);
         this.setState({
            fileContent: fileContent.value
         });
 
         return;
    }

    private getFilterQuery(): string {

        let filterQuery = "";
        const filterConditions = [];
        const {
            fromDate,
            toDate,
            dateFilterType,
            selectedDatePeriod,
            selectedAgencies,
            selectedCategory,
            selectedSubjectDomains } = this.state;

        const filterObjects = [];
        let fromDateString, toDateString;
        if (dateFilterType == DateFilterType.Period) {
            const monthCount: number = selectedDatePeriod;
            const daysOffset = monthCount * 30 * 12;
            var d = new Date();
            d.setDate(d.getDate() - daysOffset);
            fromDateString = d.toISOString();
            toDateString = new Date().toISOString();
        }
        else if (dateFilterType == DateFilterType.Range) {
            fromDateString = fromDate.toISOString();
            toDateString = toDate.toISOString();
        }

        const _fromDate = `${fromDateString.substring(0, fromDateString.indexOf('T'))}T23:59:59Z`;
        const _toDate = `${toDateString.substring(0, toDateString.indexOf('T'))}T00:00:00Z`;

        if ((dateFilterType == DateFilterType.Period && selectedDatePeriod && selectedDatePeriod != 0) || dateFilterType == DateFilterType.Range) {
            filterConditions.push(`<Geq><FieldRef Name='PublishedDate'/><Value IncludeTimeValue='TRUE' Type='DateTime'>${_fromDate}</Value></Geq>`);
            filterConditions.push(`<Leq><FieldRef Name='PublishedDate'/><Value IncludeTimeValue='TRUE' Type='DateTime'>${_toDate}</Value></Leq>`);
        }


        if (selectedAgencies.length > 0 && !selectedAgencies.includes("All")) {
            //filterConditions.push(`<Eq><FieldRef Name='FileDirRef'/><Value Type='${FieldType.Text}'>${selectedAgency}</Value></Eq>`);
            //{ fieldName: "Agency", filedType: FieldType.Taxonomy, filterValue: selectedAgency, operator: Operator.Equals, condition: Condition.And }
            const _agencyConditions: any[] = [];
            selectedAgencies.forEach((agency) => {
                _agencyConditions.push(`<Eq><FieldRef Name='FileDirRef'/><Value Type='${FieldType.Text}'>${agency}</Value></Eq>`);
            });
            const agencyConditions = MergeCAMLConditions(_agencyConditions, MergeType.Or);
            filterConditions.push(agencyConditions);
        }

        if (selectedCategory && selectedCategory != "All") {
            filterConditions.push(`<Eq><FieldRef Name='Category_x0020_Type'/><Value Type='${FieldType.Taxonomy}'>${selectedCategory}</Value></Eq>`);
            //filterObjects.push({ fieldName: "Category_x0020_Type", filedType: FieldType.Taxonomy, filterValue: selectedCategory, operator: Operator.Equals, condition: Condition.And });
        }

        // if (selectedSubjectDomains.length > 0) {
        //     const _domainConditions: any[] = [];
        //     selectedSubjectDomains.forEach((domain) => {
        //         _domainConditions.push(`<Eq><FieldRef Name='SubjectDomain'/><Value Type='${FieldType.Text}'>${domain}</Value></Eq>`);
        //         //filterObjects.push({ fieldName: "SubjectDomain", filedType: FieldType.Text, filterValue: domain, operator: Operator.Equals, condition: Condition.Or });
        //     });
        //     const domainConditions = MergeCAMLConditions(_domainConditions, MergeType.Or);
        //     filterConditions.push(domainConditions);
        // }

        const query = MergeCAMLConditions(filterConditions, MergeType.And);
        filterQuery = `<View Scope="RecursiveAll">
                            <Query>
                                <Where>
                                    ${query}
                                </Where>
                            </Query>
                        </View>`;

        console.log('filterQuery');
        console.log(filterQuery);

        return filterQuery;
    }

    private getSearchQuery(): any {
        const { fromDate,
            toDate,
            dateFilterType,
            selectedDatePeriod,
            selectedAgencies,
            selectedCategory,
            selectedSubjectDomains } = this.state;
        const filterObjects = [];
        let fromDateString, toDateString;
        if (dateFilterType == DateFilterType.Period) {
            const monthCount: number = selectedDatePeriod;
            const daysOffset = monthCount * 30 * 12;
            var d = new Date();
            d.setDate(d.getDate() - daysOffset);
            fromDateString = d.toISOString();
            toDateString = new Date().toISOString();
        }
        else if (dateFilterType == DateFilterType.Range) {
            fromDateString = fromDate.toISOString();
            toDateString = toDate.toISOString();
        }

        const _fromDate = `${fromDateString.substring(0, fromDateString.indexOf('T'))}T23:59:59Z`;
        const _toDate = `${toDateString.substring(0, toDateString.indexOf('T'))}T00:00:00Z`;
        console.log('fromDate', _fromDate);
        console.log('todate', _toDate);

        const filterConditions = [];

        let filterQuery = `<And>
                                <Geq>
                                    <FieldRef Name='PublishedDate'/>
                                    <Value IncludeTimeValue='TRUE' Type='DateTime'>${_fromDate}</Value>
                                </Geq>
                                <Leq>
                                    <FieldRef Name='PublishedDate'/>
                                    <Value IncludeTimeValue='TRUE' Type='DateTime'>${_toDate}</Value>
                                </Leq>`;

        //</And>`;

        filterConditions.push(`<And><Geq><FieldRef Name='PublishedDate'/><Value IncludeTimeValue='TRUE' Type='DateTime'>${_fromDate}</Value></Geq></And>`);
        filterConditions.push(`<And><Leq><FieldRef Name='PublishedDate'/><Value IncludeTimeValue='TRUE' Type='DateTime'>${_toDate}</Value></Leq></And>`);

        filterObjects.push({ fieldName: "PublishedDate", filedType: FieldType.Date, filterValue: _fromDate, operator: Operator.GreaterThanEquals, condition: Condition.And });
        filterObjects.push({ fieldName: "PublishedDate", filedType: FieldType.Date, filterValue: _toDate, operator: Operator.LessThanEquals, condition: Condition.And });

        // if (selectedAgency) {
        //     filterObjects.push({ fieldName: "Agency", filedType: FieldType.Taxonomy, filterValue: selectedAgency, operator: Operator.Equals, condition: Condition.And });
        // }
        if (selectedCategory) {
            filterObjects.push({ fieldName: "Category_x0020_Type", filedType: FieldType.Taxonomy, filterValue: selectedCategory, operator: Operator.Equals, condition: Condition.And });
        }
        if (selectedSubjectDomains.length > 0) {
            selectedSubjectDomains.forEach((domain) => {
                filterObjects.push({ fieldName: "SubjectDomain", filedType: FieldType.Text, filterValue: domain, operator: Operator.Equals, condition: Condition.Or });
            })

        }

        // if(filterObjects.length == 0) {
        //     filterQuery += `</And>`;
        // }
        filterQuery = "";
        const filterObjectsLength = filterObjects.length;

        if (filterObjectsLength == 5) {
            filterQuery += `<And>
                                <${filterObjects[0].operator}>
                                    <FieldRef Name='${filterObjects[0].fieldName}'/>
                                    <Value Type='${filterObjects[0].filedType}'>${filterObjects[0].filterValue}</Value>
                                </${filterObjects[0].operator}>
                                <And>
                                    <${filterObjects[1].operator}>
                                        <FieldRef Name='${filterObjects[1].fieldName}'/>
                                        <Value Type='${filterObjects[1].filedType}'>${filterObjects[1].filterValue}</Value>
                                    </${filterObjects[1].operator}>
                                <And>
                                <${filterObjects[2].operator}>
                                    <FieldRef Name='${filterObjects[2].fieldName}'/>
                                    <Value Type='${filterObjects[2].filedType}'>${filterObjects[2].filterValue}</Value>
                                </${filterObjects[2].operator}>
                                <And>
                                    <${filterObjects[3].operator}>
                                        <FieldRef Name='${filterObjects[3].fieldName}'/>
                                        <Value Type='${filterObjects[3].filedType}'>${filterObjects[3].filterValue}</Value>
                                    </${filterObjects[3].operator}>
                                    <${filterObjects[4].operator}>
                                        <FieldRef Name='${filterObjects[4].fieldName}'/>
                                        <Value Type='${filterObjects[4].filedType}'>${filterObjects[4].filterValue}</Value>
                                    </${filterObjects[4].operator}>
                            </And></And></And></And>`;
        }
        else if (filterObjectsLength == 4) {
            filterQuery += `<And>
                                <${filterObjects[0].operator}>
                                    <FieldRef Name='${filterObjects[0].fieldName}'/>
                                    <Value Type='${filterObjects[0].filedType}'>${filterObjects[0].filterValue}</Value>
                                </${filterObjects[0].operator}>
                                <And>
                                    <${filterObjects[1].operator}>
                                        <FieldRef Name='${filterObjects[1].fieldName}'/>
                                        <Value Type='${filterObjects[1].filedType}'>${filterObjects[1].filterValue}</Value>
                                    </${filterObjects[1].operator}>
                                <And>
                                    <${filterObjects[2].operator}>
                                        <FieldRef Name='${filterObjects[2].fieldName}'/>
                                        <Value Type='${filterObjects[2].filedType}'>${filterObjects[2].filterValue}</Value>
                                    </${filterObjects[2].operator}>
                                    <${filterObjects[3].operator}>
                                        <FieldRef Name='${filterObjects[3].fieldName}'/>
                                        <Value Type='${filterObjects[3].filedType}'>${filterObjects[3].filterValue}</Value>
                                    </${filterObjects[3].operator}>
                            </And></And></And>`;
        }
        else if (filterObjectsLength == 3) {
            filterQuery += `<And>
                                <${filterObjects[0].operator}>
                                    <FieldRef Name='${filterObjects[0].fieldName}'/>
                                    <Value Type='${filterObjects[0].filedType}'>${filterObjects[0].filterValue}</Value>
                                </${filterObjects[0].operator}>
                                <And>
                                    <${filterObjects[1].operator}>
                                        <FieldRef Name='${filterObjects[1].fieldName}'/>
                                        <Value Type='${filterObjects[1].filedType}'>${filterObjects[1].filterValue}</Value>
                                    </${filterObjects[1].operator}>
                                    <${filterObjects[2].operator}>
                                        <FieldRef Name='${filterObjects[2].fieldName}'/>
                                        <Value Type='${filterObjects[2].filedType}'>${filterObjects[2].filterValue}</Value>
                                    </${filterObjects[2].operator}>
                            </And></And>`;
        }
        else if (filterObjectsLength == 2) {
            filterQuery += `<And>
                                <${filterObjects[0].operator}>
                                    <FieldRef Name='${filterObjects[0].fieldName}'/>
                                    <Value Type='${filterObjects[0].filedType}'>${filterObjects[0].filterValue}</Value>
                                </${filterObjects[0].operator}>
                                <${filterObjects[1].operator}>
                                    <FieldRef Name='${filterObjects[1].fieldName}'/>
                                    <Value Type='${filterObjects[1].filedType}'>${filterObjects[1].filterValue}</Value>
                                </${filterObjects[1].operator}>
                            </And>`;
        }
        else if (filterObjectsLength == 1) {
            filterQuery += `<${filterObjects[0].operator}>
                                    <FieldRef Name='${filterObjects[0].fieldName}'/>
                                    <Value Type='${filterObjects[0].filedType}'>${filterObjects[0].filterValue}</Value>
                            </${filterObjects[0].operator}>`;
        }


        const filter = `(PublishedDate ge datetime'${_fromDate}') and (PublishedDate le datetime'${_toDate}')`;
        let searchQuery = `<View Scope="RecursiveAll">
                                <Query>
                                    <Where>
                                        ${filterQuery}
                                    </Where>
                                </Query>
                            </View>`;


        //const query = JSON.stringify(searchQuery).trim();

        return searchQuery;
    }

    private handleReset() {
        this.setState({
            selectedAgencies: ["All"],
            selectedCategory: "All",
            selectedSubjectDomains: [],
            selectedDatePeriod: 1,
            fromDate: undefined,
            toDate: undefined,
            searchQuery: "",
            searchPanelText: "",
            searchResults: []
        });
    }

    private async searchDocuments(): Promise<any> {
        const libraryName = 'Proposal Folder';
        const searchQuery = {
            ViewXml: `<View>
                                            <Query>
                                                <Where>
                                                </Where>
                                            </Query>
                                           </View>`
        }
        const requestUri = `${this.props.SiteUrl}/_api/web/lists/getbytitle('${libraryName}')/GetItems(query=@v1)?@v1=${JSON.stringify(searchQuery)}`;
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

    private handleSearchQueryChange(e: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newText: string) {
        this.setState({
            searchQuery: newText
        });
    }

    private clearSearch() {
        this.setState({
            searchQuery: "",
            searchResults: []
        });
    }

    private async searchQuery(query: string): Promise<void> {
        this.setState({
            inProgress: true
        });
        const results = await this._spService.getSearchResults(query);
        const allResults = this.state.allLibraryResults;
        allResults.map((r) => {
            const decodedUrl = decodeURI(r.EncodedAbsUrl);
            console.log(decodedUrl);
        });
        let searchResults = allResults.filter((result) => {
            const decodedUrl = decodeURI(result.EncodedAbsUrl);
            let summary = "";
            const _results = results.filter((r) => { 
                if(r.Path == decodedUrl) {
                    summary = r.HitHighlightedSummary;
                    return true;
                }
            });
            if(_results && _results.length > 0) {
                result.Summary = summary;
                return true;
            }
        });

        // apply filters 
        const { fromDate, toDate, selectedCategory, selectedAgencies } = this.state;
        if(fromDate && toDate) {
            searchResults = searchResults.filter((result) => {
                return result.PublishedDate >= fromDate.toLocaleDateString() &&
                       result.PublishedDate <= toDate.toLocaleDateString()
            });
        }

        if(selectedAgencies.length > 0 && !selectedAgencies.includes("All")) {
            const libraryPath = `${this.props.SiteUrl}/${LIBRARY_PROPOSAL}/`;
            searchResults = searchResults.filter((result) => {
                const agency = result.FileDirRef.replace(libraryPath, "");
                return selectedAgencies.includes(agency);
            });
        }

        
        if(selectedCategory && selectedCategory != "All") {
            searchResults = searchResults.filter((result) => result.Category == this.state.selectedCategory);
        }

        console.log("results");
        console.log(results);
        console.log("search results");
        console.log(searchResults);

        this.setState({
            querySearchResults: results,
            searchResults,
            inProgress: false
        });
    }

    private onTaxPickerChangeAgency(terms: IPickerTerms) {
        console.log("Terms", terms);
    }

    private onTaxPickerChangeCategory(terms: IPickerTerms) {
        console.log("Terms", terms);
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

    private _onRenderCell(item: IProposalItem, index: number, isScrolling: boolean): JSX.Element {
        console.log('doc type', item.DocIcon);
        const extn = item.DocIcon;
        const iconUrl = extn == 'url' ?
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
                                    path={item.EncodedAbsUrl} />
                            </StackItem>
                            <StackItem>
                                <div className="ms-ListGhostingExample-itemName searchResultLink">
                                    <Link target="_blank" data-interception="off" href={item.ServerRedirectedEmbedUrl} >{item.FileLeafRef}</Link>
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

    private _onDismiss() {

    }

    private handleSearchPanelTextChange(searchText: string) {
        this.setState({
            searchPanelText: searchText
        })
    }

    public render(): React.ReactElement<ILibrarySearchProps> {
        const {
            dateFilterItems,
            dateFilterType,
            selectedDatePeriod,
            selectedAgencies,
            selectedCategory,
            selectedSubjectDomains,
            fromDate,
            toDate,
            searchResults,
            termSetIdAgency,
            termSetIdCategory,
            agencies,
            categories } = this.state;
        const dateTypeOptions = DATE_FILTER_TYPES.map((t) => ({ key: t, text: t }));
        const periodOptions = dateFilterItems.map((p) => ({ key: p.Value, text: p.Title }));
        const domainOptions = SUBJECT_DOMAIN.map((d) => ({ key: d, text: d }));
        const agencyOptions = agencies.map((a) => ({ key: a.ServerRelativeUrl, text: a.Name }));
        const categoryOptions = categories.map((c) => ({ key: c, text: c }));
        agencyOptions.unshift({ key: "All", text: "All" });
        categoryOptions.unshift({ key: "All", text: "All" });
        periodOptions.unshift({ key: 0, text: "All" });
        return (
            <div className="librarySearchContainer">
                <Stack style={{ display: "none" }}>
                    <StackItem className="pad20">
                        <Label>
                            Filtered Search
                        </Label>
                    </StackItem>
                </Stack>
                <Stack horizontal className="width100 filterRow">
                    <StackItem>
                        <Stack styles={dateColumnStackStyles}>
                            <StackItem className="dateSection" >
                                <Stack horizontal style={{ justifyContent: 'space-between' }}>
                                    <StackItem styles={headerRowStyles}>
                                        <Label>
                                            Date Created
                                        </Label>
                                    </StackItem>
                                    <StackItem>
                                        <ChoiceGroup
                                            options={dateTypeOptions}
                                            selectedKey={dateFilterType}
                                            className="searchType"
                                            onChange={(e, o) => this.handleDateFilterTypeChange(e, o)} />
                                    </StackItem>
                                </Stack>
                            </StackItem>
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
                                        <Stack horizontal className="dateRangeContainer">
                                            <StackItem className="col50" style={{ display: "flex" }} >
                                                <Label style={{ marginRight: "5px" }}>
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
                                            </StackItem>
                                            <StackItem className="col50" style={{ display: "flex" }} >
                                                <Label style={{ marginRight: "5px", marginLeft: "5px" }}>
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
                                            </StackItem>
                                        </Stack>
                                    </StackItem>
                            }
                        </Stack>
                    </StackItem>
                    <StackItem>
                        <Stack styles={columnStackStyles}>
                            <StackItem styles={headerRowStyles}>
                                <Label>
                                    Agency
                                </Label>
                            </StackItem>
                            <StackItem>
                                <Dropdown
                                    selectedKeys={selectedAgencies}
                                    options={agencyOptions}
                                    multiSelect
                                    onChange={(e, option) => this.handleAgencyChange(e, option)} />
                            </StackItem>
                        </Stack>
                    </StackItem>
                    <StackItem>
                        <Stack styles={columnStackStyles}>
                            <StackItem styles={headerRowStyles}>
                                <Label>
                                    Category
                                </Label>
                            </StackItem>
                            <StackItem>
                                <Dropdown
                                    selectedKey={selectedCategory}
                                    options={categoryOptions}
                                    onChange={(e, option) => this.handleCategoryChange(e, option)} />
                            </StackItem>
                        </Stack>
                    </StackItem>
                    <StackItem>
                        <Stack styles={columnStackStyles} style={{ display: 'none' }}>
                            <StackItem styles={headerRowStyles}>
                                <Label>
                                    Subject Domain
                                </Label>
                            </StackItem>
                            <StackItem>
                                <Dropdown
                                    selectedKeys={selectedSubjectDomains}
                                    placeholder="Select"
                                    options={domainOptions}
                                    multiSelect
                                    onChange={(e, option) => this.handleDomainChange(e, option)} />
                            </StackItem>
                        </Stack>
                    </StackItem>
                    <StackItem>
                        <Stack styles={columnStackStyles}>
                            <StackItem styles={headerRowStyles}>
                                <Label>
                                    Text Search
                                </Label>
                            </StackItem>
                            <StackItem>
                                {/* <SearchBox placeholder="Search"
                                    value={this.state.searchQuery}
                                    onChange={(e, value) => this.handleSearch(value)}
                                    onSearch={(value) => this.searchQuery(value)}
                                    onClear={() => this.clearSearch()} /> */}
                                <TextField
                                    placeholder="Search query"
                                    value={this.state.searchQuery}
                                    onChange={(e, value) => this.handleSearchQueryChange(e, value)}
                                />
                            </StackItem>
                        </Stack>
                    </StackItem>
                </Stack>
                <Stack styles={stackStyleBtnRow} style={{ alignItems: "center", justifyContent: "center", width: "100%" }} horizontal >
                    <StackItem>
                        <DefaultButton text="Reset" onClick={() => this.handleReset()} style={{ marginRight: '20px' }} />
                        <PrimaryButton text="Search" onClick={() => this.handleSearch()} />
                    </StackItem>
                </Stack>
                <div className="searchResultsContainer">
                    <FocusZone direction={FocusZoneDirection.vertical}>
                        <div className="ms-ListGhostingExample-container" data-is-scrollable={true}>
                            {
                                this.state.inProgress ?

                                    <Spinner size={SpinnerSize.medium} />

                                    :

                                    // <List
                                    //     items={searchResults}
                                    //     onRenderCell={this._onRenderCell}
                                    // />
                                    <SearchResults
                                        Context={this.props.Context}
                                        IsLoading={this.state.inProgress}
                                        SearchPanelText={this.state.searchPanelText}
                                        SearchResults={this.state.searchResults}
                                        OnSearchTextChange={this.handleSearchPanelTextChange} />

                            }
                        </div>
                    </FocusZone>
                </div>
                <div dangerouslySetInnerHTML={{__html: this.state.fileContent}}>

                </div>
            </div>
        )
    }
}