export class Constants {

    public static readonly COMPONENT_NAME = 'SP Search';

}

export const SearchTypes = ['Basic', 'Advanced'];

export enum SearchType {
    Basic = 'Basic',
    Advanced = 'Advanced'
}

export enum DateFilterType {
    Period = 'Period',
    Range = 'Range'
}

export const DATE_FILTER_TYPES = ["Period", "Range"];

export const DATE_FILTER_PERIODS = ["Last 1 month", "Last 2 months", "Last 3 months", "Last 6 months", "Last 1 year"];

export enum BuiltinSourceIds {
    Documents = 'e7ec8cee-ded8-43c9-beb5-436b54b31e84',
    ItemsMatchingContentType = '5dc9f503-801e-4ced-8a2c-5d1237132419',
    ItemsMatchingTag = 'e1327b9c-2b8c-4b23-99c9-3730cb29c3f7',
    ItemsRelatedToCurrentUser = '48fec42e-4a92-48ce-8363-c2703a40e67d',
    ItemsWithSameKeywordAsThisItem = '5c069288-1d17-454a-8ac6-9c642a065f48',
    LocalPeopleResults = 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
    LocalReportsAndDataResults = '203fba36-2763-4060-9931-911ac8c0583b',
    LocalSharePointResults = '8413cd39-2156-4e00-b54d-11efd9abdb89',
    LocalVideoResults = '78b793ce-7956-4669-aa3b-451fc5defebf',
    Pages = '5e34578e-4d08-4edc-8bf3-002acf3cdbcc',
    Pictures = '38403c8c-3975-41a8-826e-717f2d41568a',
    Popular = '97c71db1-58ce-4891-8b64-585bc2326c12',
    RecentlyChangedItems = 'ba63bbae-fa9c-42c0-b027-9a878f16557c',
    RecommendedItems = 'ec675252-14fa-4fbe-84dd-8d098ed74181',
    Wiki = '9479bf85-e257-4318-b5a8-81a180f5faa1',
}

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


