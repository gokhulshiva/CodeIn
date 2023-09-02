import { FontWeights, IButtonStyles, IIconProps, IStackItemStyles, IStackStyles, getTheme, mergeStyleSets } from "office-ui-fabric-react";

const theme = getTheme();
export const contentStyles = mergeStyleSets({
    container: {
        display: 'flex',
        flexFlow: 'column nowrap',
        alignItems: 'stretch',
        height: '450px !important',
        borderRadius: '5px',
        textAlign: 'center',
        padding: '30px'
    },
    containerSmall: {
        display: 'flex',
        flexFlow: 'column nowrap',
        alignItems: 'stretch',
        height: '320px !important',
        borderRadius: '5px',
        textAlign: 'center',
        padding: '30px'
    },
    header: [
        // eslint-disable-next-line deprecation/deprecation
        theme.fonts.xLargePlus,
        {
            flex: '1 1 auto',
            borderTop: `4px solid ${theme.palette.themePrimary}`,
            color: theme.palette.neutralPrimary,
            display: 'flex',
            alignItems: 'center',
            fontWeight: FontWeights.semibold,
            padding: '12px 12px 14px 24px',
            borderBottom: '1px solid #dee2e6'
        },
    ],
    heading: {
        color: theme.palette.neutralPrimary,
        fontWeight: FontWeights.semibold,
        fontSize: 'inherit',
        margin: '0',
    },
    body: {
        flex: '4 4 auto',
        padding: '0 24px 24px 24px',
        overflowY: 'hidden',
        selectors: {
            p: { margin: '14px 0' },
            'p:first-child': { marginTop: 0 },
            'p:last-child': { marginBottom: 0 },
        },
    },
    bodySm: {
        flex: '4 4 auto',
        padding: '14px 24px 24px 24px',
        overflowY: 'hidden',
        selectors: {
            p: { margin: '14px 0' },
            'p:first-child': { marginTop: 0 },
            'p:last-child': { marginBottom: 0 },
        },
    }
});


export const stackStyles: Partial<IStackStyles> = { root: { marginBottom: 20 } };

export const stackItemStyles: Partial<IStackItemStyles> = { root: { display: "flex", justifyContent: "flex-start", padding: "10px 0" } };

export const stackStyleBtnRow: Partial<IStackStyles> = { root: { marginTop: 20, marginBottom: 20, justifyContent: "center" } };

export const stackStyleHeaderRow: Partial<IStackStyles> = { root: { marginBottom: 20, justifyContent: "space-between" } };

export const iconButtonStyles: Partial<IButtonStyles> = {
    root: {
        color: theme.palette.neutralPrimary,
        marginLeft: 'auto',
        marginTop: '4px',
        marginRight: '2px',
    },
    rootHovered: {
        color: theme.palette.neutralDark,
    },
};

export const cancelIcon: IIconProps = { iconName: 'Cancel' };
export const accessIcon: IIconProps = { iconName: 'AddFriend' };
export const editIcon: IIconProps = { iconName: 'Edit' };
export const deleteIcon: IIconProps = { iconName: 'Delete' };