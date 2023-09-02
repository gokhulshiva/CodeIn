import { IFolderInfo } from "@pnp/sp/folders";
import { MessageBarType } from "office-ui-fabric-react";

export interface ISpOnlineFileUploadState {
    inProgress: boolean;
    showNotification: boolean;
    notificationType: MessageBarType;
    notificationText: string;
    fileObject: any;
    selectedAgency: string;
    agencies: IFolderInfo[];
    categories: string[];
    subjectDomains: string[];
    fileID: number;
    showMetaDataForm: boolean;
    metaDataAgency: string;
    metaDataCategory: string;
    publishedDate: Date;
}