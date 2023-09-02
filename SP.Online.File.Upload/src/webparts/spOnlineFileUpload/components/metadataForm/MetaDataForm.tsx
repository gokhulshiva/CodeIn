import * as React from "react";
import { IMetaDataFormProps } from "./IMetaDataFormProps";
import { IMetaDataFormState } from "./IMetaDataFormState";
import { DatePicker, Dropdown, IDropdownOption, IIconProps, IconButton, MessageBar, MessageBarType, Modal, PrimaryButton, Spinner, SpinnerSize, Stack, StackItem, mergeStyleSets } from "office-ui-fabric-react";
import { contentStyles, iconButtonStyles, stackItemStyles, stackStyleBtnRow, stackStyles } from "../../common/fabricStyles";
import { BTN_TEXT_UPDATE, LIBRARY_PROPOSAL, MESSAGE_FILE_UPLOAD_ERROR, MESSAGE_FILE_UPLOAD_SUCCESS, Result } from "../../common/constants";

const cancelIcon: IIconProps = { iconName: 'Cancel' };

const styles = mergeStyleSets({
    root: { selectors: { '> *': { marginBottom: 15 } } },
    control: { maxWidth: 300, marginBottom: 15 },
});

export default class MetaDataForm extends React.Component<IMetaDataFormProps, IMetaDataFormState> {
    private datePickerRef: any;
    constructor(props: IMetaDataFormProps) {
        super(props);
        this.state = {
            inProgress: false,
            notificationText: "",
            notificationType: MessageBarType.success,
            showNotification: false
        }
    }

    private closeModal() {
        this.props.CloseModal();
    }

    private async handleUpdate() {
        this.setState({
            inProgress: true
        });
        const body = {
            PublishedDate: this.props.PublishedDate,
            Category: this.props.Category
        }
        const libraryName = LIBRARY_PROPOSAL;
        const updateFileResult = await this.props.SPService.updateFile(libraryName, this.props.FileID, body);
        const result = updateFileResult == Result.SUCCESS ? Result.SUCCESS : Result.ERROR;
        const notificationText = updateFileResult == Result.SUCCESS ? MESSAGE_FILE_UPLOAD_SUCCESS : MESSAGE_FILE_UPLOAD_ERROR;
        const notificationType = updateFileResult == Result.SUCCESS ? MessageBarType.success : MessageBarType.error;
        this.setState({
            inProgress: false
        });
        this.props.UpdateContext(result, notificationText, notificationType);
    }

    private handleAgencyChange(ev: React.SyntheticEvent<HTMLElement>, option: IDropdownOption) {
        const value = option.key as string;
        this.props.OnAgencyChange(value);
    }

    private handleCategoryChange(ev: React.SyntheticEvent<HTMLElement>, option: IDropdownOption) {
        const value = option.key as string;
        this.props.OnCategoryChange(value);
    }

    private handlePublishedDateChange(date: Date) {
        this.props.OnPublishedDateChange(date);
    }

    private onFormatDate = (date?: Date): string => {
        return !date ? '' : date.getDate() + '/' + (date.getMonth() + 1) + '/' + (date.getFullYear() % 100);
    };

    public render(): React.ReactElement<IMetaDataFormProps> {
        const titleId = 'Update Metadata';
        const { Domain: domain, 
                Category: category, 
                DomainOptions: domainOptions, 
                CategoryOptions: categoryOptions,
                PublishedDate: publishedDate } = this.props;
        const btnDisabled = !publishedDate || !category;
        return (
            <div className="metaDataFormContainer">
                <Modal
                    titleAriaId={"Metadata Form Modal"}
                    isOpen={this.props.IsModalOpen}
                    onDismiss={() => this.closeModal()}
                    isBlocking={false}
                    containerClassName={contentStyles.container}
                >
                    <div className={contentStyles.header}>
                        <h2 className={contentStyles.heading} id={titleId}>
                            {`Update Metadata`}
                        </h2>
                        <IconButton
                            styles={iconButtonStyles}
                            iconProps={cancelIcon}
                            ariaLabel="Close popup modal"
                            onClick={() => this.closeModal()}
                        />
                    </div>

                    <div className={contentStyles.body} >

                        <Stack styles={stackStyles}>
                            <StackItem styles={stackItemStyles}>
                                <label className="form-check-label" >
                                    {`Category`}
                                </label>
                            </StackItem>
                            <StackItem>
                                <Dropdown
                                    selectedKey={category}
                                    options={categoryOptions}
                                    onChange={(e, option) => this.handleCategoryChange(e, option)} />
                            </StackItem>
                        </Stack>

                        <Stack styles={stackStyles}>
                            <StackItem styles={stackItemStyles}>
                                <label className="form-check-label" >
                                    {`Published Date`}
                                </label>
                            </StackItem>
                            <StackItem>
                                <DatePicker
                                    componentRef={this.datePickerRef}
                                    //label="Start date"
                                    allowTextInput
                                    ariaLabel="Select a date. Input format is day slash month slash year."
                                    value={publishedDate}
                                    onSelectDate={(date?: Date) => this.handlePublishedDateChange(date)}
                                    formatDate={this.onFormatDate}
                                    //parseDateFromString={onParseDateFromString}
                                    className={styles.control}
                                // DatePicker uses English strings by default. For localized apps, you must override this prop.
                                />
                            </StackItem>
                        </Stack>

                        <Stack styles={stackStyles}>
                            <div style={{ display: this.state.showNotification ? 'block' : 'none' }}>
                                <MessageBar messageBarType={this.state.notificationType}>
                                    {this.state.notificationText}
                                </MessageBar>
                            </div>
                        </Stack>

                        <Stack styles={stackStyleBtnRow} horizontal>
                            <StackItem>
                                <Spinner size={SpinnerSize.medium} style={{
                                    display: this.state.inProgress ? 'block' : 'none',
                                    marginTop: '8px',
                                    marginRight: '10px'
                                }} />
                            </StackItem>
                            <StackItem>
                                <PrimaryButton
                                    text={BTN_TEXT_UPDATE}
                                    disabled={btnDisabled}
                                    onClick={() => this.handleUpdate()} />
                            </StackItem>
                        </Stack>
                    </div>
                </Modal >
            </div >
        )
    }
}