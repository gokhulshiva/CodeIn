import { FontWeights, IButtonStyles, IIconProps, IStackItemStyles, IStackStyles, getTheme, loadTheme, mergeStyleSets } from "office-ui-fabric-react";



const theme = getTheme();
const myTheme = loadTheme({
    palette: {
      themePrimary: '#9d7bab',
      themeLighterAlt: '#fbf9fc',
      themeLighter: '#eee7f2',
      themeLight: '#e0d2e6',
      themeTertiary: '#c3aacd',
      themeSecondary: '#a888b5',
      themeDarkAlt: '#8e6f9a',
      themeDark: '#785d82',
      themeDarker: '#584560',
      neutralLighterAlt: '#faf9f8',
      neutralLighter: '#f3f2f1',
      neutralLight: '#edebe9',
      neutralQuaternaryAlt: '#e1dfdd',
      neutralQuaternary: '#d0d0d0',
      neutralTertiaryAlt: '#c8c6c4',
      neutralTertiary: '#a19f9d',
      neutralSecondary: '#605e5c',
      neutralPrimaryAlt: '#3b3a39',
      neutralPrimary: '#323130',
      neutralDark: '#201f1e',
      black: '#000000',
      white: '#ffffff',
    }});
export const contentStyles = mergeStyleSets({
    container: {
        display: 'flex',
        flexFlow: 'column nowrap',
        alignItems: 'stretch',
        height: '380px !important',
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

export const columnStackStyles: Partial<IStackStyles> = { root: { minWidth: 300,  marginRight: "10px" } };

export const dateColumnStackStyles: Partial<IStackStyles> = { root: { minWidth: 450,  marginRight: "10px" } };

export const headerRowStyles: Partial<IStackItemStyles> = { root: { padding: "10px 0 20px 0" } };

export const stackStyleBtnRow: Partial<IStackStyles> = { root: { padding: "40px 0", justifyContent: "flex-end" } };

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