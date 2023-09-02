import { IPeopleItem } from "./IPeopleItem";
import { ITaxonomyItem } from "./ITaxonomyItem";

export interface IProposalItem {
    BaseName: string;
    Category_x0020_Type: ITaxonomyItem;
    Category: string;
    SubjectDomain: string;
    Author: IPeopleItem;
    Editor: IPeopleItem;
    FileDirRef: string;
    FileLeafRef: string;
    FileRef: string;
    DocIcon: string;
    EncodedAbsUrl: string;
    FSObjType: string;
    ServerRedirectedEmbedUrl: string;
    ServerUrl: string;
    FileSizeDisplay: string;
    PublishedDate: string;
    File_x0020_Type: string;
    "File_x0020_Type.mapapp": string;
    "HTML_x0020_File_x0020_Type.File_x0020_Type.mapico": string;
    Summary?: string;
}