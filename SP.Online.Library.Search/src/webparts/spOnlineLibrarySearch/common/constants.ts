export const LIST_DATE_FILTER_CONFIG = 'DateFilterConfig';
export const LIBRARY_PROPOSAL = 'Proposal Folder';

export enum DateFilterType {
    Period = 'Period',
    Range = 'Range'
}

export enum FieldType {
    Date = 'DateTime',
    Text = 'Text',
    Choice = 'Choice',
    Taxonomy = 'TaxonomyFieldType'
}

export enum FILETYPE {
    DOC = 'doc',
    DOCX = 'docx',
    PDF = 'pdf',
    PNG = 'png',
    XLSX = 'xlsx',
    XLS = 'xls',
    PPT = 'ppt',
    PPTX = 'pptx',

}

export enum Condition {
    And = 'And',
    Or = 'Or'
}

export enum ItemType {
    Folder = "1",
    File = "0"
}

export enum Operator {
    LessThan = "Le",
    GreaterThan = "Ge",
    Equals = "Eq",
    GreaterThanEquals = "Geq",
    LessThanEquals = "Leq"
}

export enum TaxonomyGroupType {
    RegularGroup = "RegularGroup",
    SiteCollectionGroup = "SiteCollectionGroup",
    SystemGroup = "SystemGroup"
}

export const DATE_FILTER_TYPES = ["Period", "Range"];

export const DATE_FILTER_PERIODS = [
                                        {key: 1, text:"Last 1 month"}, 
                                        {key: 2, text:"Last 2 months"}, 
                                        {key: 3, text:"Last 3 months"},
                                        {key: 6, text:"Last 6 months"}, 
                                        {key: 12, text:"Last 1 year"}
                                   ];

export const SUBJECT_DOMAIN = ["Artificial Intelligence", "Business Intelligence", "Content Management"];

export const Properties = [
    'Title',
    'Path',
    'DefaultEncodingURL',
    'FileType',
    'HitHighlightedSummary',
    'HitHighlightedProperties',
    'AuthorOWSUSER',
    'owstaxidmetadataalltagsinfo',
    'Created',
    'UniqueID',
    'NormSiteID',
    'NormWebID',
    'NormListID',
    'NormUniqueID',
    'ContentTypeId',
    'contentclass',
    'UserName',
    'JobTitle',
    'WorkPhone',
    'SPSiteURL',
    'SPWebUrl',
    'SiteTitle',
    'CreatedBy',
    'HtmlFileType',
    'SiteLogo',
    'PictureThumbnailURL',
    'ServerRedirectedURL',
    'ServerRedirectedEmbedURL',
    'ServerRedirectedPreviewURL',
    'LinkingUrl',
    'FileExtension',
    'IsDocument'
];