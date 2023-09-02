export interface IFileItem {
    UniqueId: string;
    Name: string;
    ServerRelativeUrl: string;
    CheckOutType: number;
    ListItemAllFields: {
        ID: number;
        AuthorId: number;
        CheckoutUserId: number;
        SharedWithUsersId: number[];
        EffectiveBasePermissions: { High: number, Low: number };
        FieldValuesAsText: {
            Created: string;
            Modified: string;
            Editor: string;
            Author: string;
            Status: string;
            ApprovedDate: string;
            CheckedOutTitle: string;
            CheckedOutUserId: string;
            SharedWithUsers: string;
            File_x005f_x0020_x005f_Type: string;
        }
    }
}