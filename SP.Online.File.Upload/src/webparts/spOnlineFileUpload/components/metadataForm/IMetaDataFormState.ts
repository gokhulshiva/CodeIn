import { MessageBarType } from "office-ui-fabric-react";

export interface IMetaDataFormState {
    showNotification: boolean;
    notificationText: string;
    notificationType: MessageBarType;
    inProgress: boolean;
}