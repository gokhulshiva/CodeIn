import { IDropdownOption, MessageBarType } from "office-ui-fabric-react";
import { Result } from "../../common/constants";
import { SPService } from "../../services/SPService";

export interface IMetaDataFormProps {
    Context: any;
    SiteRelativeUrl?: string;
    SPService: SPService;
    IsModalOpen: boolean;
    Domain: string;
    Category: string;
    PublishedDate: Date;
    CategoryOptions: IDropdownOption[];
    DomainOptions: IDropdownOption[];
    FileID: number;
    CloseModal: () => void;
    OnAgencyChange: (newValue: string) => void;
    OnCategoryChange: (newValue: string) => void;
    OnPublishedDateChange: (newDate: Date) => void;
    UpdateContext: (result: Result, notificationText: string, notificationType: MessageBarType) => void;
}