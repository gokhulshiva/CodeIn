import * as React from 'react';
import styles from './SpOnlineFileUpload.module.scss';
import { ISpOnlineFileUploadProps } from './ISpOnlineFileUploadProps';
import { ISpOnlineFileUploadState } from './ISpOnlineFileUploadState';
import { escape } from '@microsoft/sp-lodash-subset';
import { DefaultButton, Dropdown, IDropdownOption, Label, MessageBar, MessageBarType, PrimaryButton, Spinner, SpinnerSize, Stack, StackItem, loadTheme } from 'office-ui-fabric-react';
import { getMetadata } from "docx-templates";
import * as fs from 'fs';
import { BTN_TEXT_RESET, BTN_TEXT_UPLOAD, LABEL_TEXT_FILE, LABEL_TEXT_FILE_UPLOAD, LABEL_TEXT_FOLDER, MESSAGE_FILE_UPLOAD_ERROR, MESSAGE_FILE_UPLOAD_SUCCESS, Result, WARNING_TEXT_FILE_EXISTS } from '../common/constants';
import { stackStyleBtnRow } from '../common/fabricStyles';
import { SPService } from '../services/SPService';
import { IFolderInfo } from '@pnp/sp/folders';
import MetaDataForm from './metadataForm/MetaDataForm';
import { IFileItem } from '../models/IFileItem';
require('../css/style.css');

export default class SpOnlineFileUpload extends React.Component<ISpOnlineFileUploadProps, ISpOnlineFileUploadState> {
  private _spService: SPService;
  constructor(props: ISpOnlineFileUploadProps) {
    super(props);
    this._spService = new SPService(this.props.context);
    loadTheme({
      palette: {
        themePrimary: '#5d21a6',
        themeLighterAlt: '#f7f3fb',
        themeLighter: '#e0d2f1',
        themeLight: '#c6ade4',
        themeTertiary: '#9469c9',
        themeSecondary: '#6c34b0',
        themeDarkAlt: '#541e95',
        themeDark: '#47197e',
        themeDarker: '#34135d',
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
      }
    });
    this.state = {
      inProgress: false,
      notificationText: "",
      notificationType: MessageBarType.info,
      showNotification: false,
      fileObject: null,
      selectedAgency: "",
      agencies: [],
      categories: [],
      subjectDomains: [],
      fileID: 0,
      metaDataAgency: "",
      metaDataCategory: "",
      showMetaDataForm: false,
      publishedDate: undefined
    }

    this.handleMetaDataAgencyChange = this.handleMetaDataAgencyChange.bind(this);
    this.handleMetaDataCategoryChange = this.handleMetaDataCategoryChange.bind(this);
    this.updateContextMetaDataForm = this.updateContextMetaDataForm.bind(this);
    this.handlePublishedDateChange = this.handlePublishedDateChange.bind(this);
    this.closeMetaDataForm = this.closeMetaDataForm.bind(this);
  }

  componentDidMount(): void {
    this.loadData();
  }

  private async loadData(): Promise<void> {
    const folders: IFolderInfo[] = await this._spService.getAgencyFolders();
    const agencies = folders.filter((f) => f.Name != 'Forms');
    const categories = await this._spService.getCategories();
    const subjectDomains = await this._spService.getDomains();
    this.setState({
      agencies,
      categories,
      subjectDomains
    });
  }

  private async handleFileChange(e: any) {
    if (e.target && e.target.files && e.target.files.length > 0) {
      const fileObject = e.target.files[0];
      const fileName = fileObject.name;
      this.setState({
        fileObject
      });
      // const arrayBuffer = await fileObject.arrayBuffer();
      // console.log('file object');
      // console.log(arrayBuffer);

      // const metaData = await getMetadata(arrayBuffer);
      // console.log('metadata');
      // console.log(metaData);


    }
    else {

    }
  }

  private closeNotification() {
    this.setState({
      showNotification: false
    });
  }

  private handleAgencyChange(ev: React.SyntheticEvent<HTMLElement>, option: IDropdownOption) {
    console.log('agency path');
    console.log(option.key);
    this.setState({
      selectedAgency: option.key as string
    });
  }

  private handleMetaDataAgencyChange(newAgency: string) {
    this.setState({
      metaDataAgency: newAgency
    });
  }

  private handleMetaDataCategoryChange(newCategory: string) {
    this.setState({
      metaDataCategory: newCategory
    });
  }

  private handlePublishedDateChange(date: Date) {
    this.setState({
      publishedDate: date
    });
}

  private async handleUpload() {
    // this.setState({
    //     inProgress: true
    // });
    const libraryPath = 'Proposal Folder';
    const folderPath = this.state.selectedAgency;  //`${this.props.siteRelativeUrl}/${libraryPath}`;
    const fileName = this.state.fileObject.name;
    const filePath = `${folderPath}/${fileName}`;
    console.log('filepath', filePath);
    const { isFileExists, authorId }: any = await this._spService.checkIfFileExists(filePath);
    if (isFileExists) {
      this.setState({
        notificationText: WARNING_TEXT_FILE_EXISTS,
        notificationType: MessageBarType.warning,
        showNotification: true
      });
      return;
    }

    this.setState({
      inProgress: true
    });

    const uploadResult = await new Promise<Result>((resolve) => {
      this._spService.uploadFile(this.state.fileObject, folderPath)
        .then((fileUploadResult) => {
          //this._spServices.breakFileInheritance(filePath);
          console.log('fileUploadResult', fileUploadResult);
          const fileID = fileUploadResult.fileID;
          const fileUrl = fileUploadResult.fileUrl;
          console.log('fileID', fileID);

          // check metadata
          this._spService.getFileById(fileID).then((data: IFileItem) => {
            if (!data) {
              resolve(Result.ERROR);
            }

            if (!data.PublishedDate || !data.Category) {
              this.setState({
                metaDataCategory: data.Category || "",
                publishedDate: data.PublishedDate,
                fileID: data.ID,
                showMetaDataForm: true
              });
            }
            else {
              resolve(Result.SUCCESS);
            }

          })

        }).catch((error) => {
          console.log('uploadFile error', error);
          resolve(Result.ERROR);
        });
    });

    setTimeout(() => {
      const notificationText = uploadResult == Result.SUCCESS ? MESSAGE_FILE_UPLOAD_SUCCESS : MESSAGE_FILE_UPLOAD_ERROR;
      const notificationType = uploadResult == Result.SUCCESS ? MessageBarType.success : MessageBarType.error;
      this.setState({
        notificationText,
        notificationType,
        showNotification: true,
        inProgress: false
      });
    }, 3000);

    //return;
    //const fileUploadResult = await this._spServices.uploadFile(this.state.fileObject, libraryPath, metadata);
  }

  private resetControls() {
    const fileUploadCtrl: any = document.getElementById('fileUploadCtrl');
    fileUploadCtrl.value = null;
    this.setState({
      fileObject: null,
      selectedAgency: ""
    });

  }

  private updateContextMetaDataForm(result: Result, notificationText: string, notificationType: MessageBarType) {
    this.setState({
      notificationText,
      notificationType,
      showNotification: true,
      showMetaDataForm: false,
      inProgress: false,
    });
    this.resetControls();
  }

  private closeMetaDataForm() {
    this.setState({
      showMetaDataForm: false
    });
  }

  public render(): React.ReactElement<ISpOnlineFileUploadProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    const {
      showNotification,
      notificationType,
      notificationText,
      inProgress,
      fileObject,
      agencies,
      categories,
      subjectDomains,
      selectedAgency,
      metaDataAgency,
      metaDataCategory,
      showMetaDataForm } = this.state;
    const btnDisabled = fileObject == null || !selectedAgency;
    const agencyOptions = agencies.map((a) => ({ key: a.ServerRelativeUrl, text: a.Name }));
    const domainOptions = subjectDomains.map((a) => ({ key: a, text: a }));
    const categoryOptions = categories.map((c) => ({ key: c, text: c }));

    console.log('domains', domainOptions);

    return (
      <section className={`${styles.spOnlineFileUpload} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className="fileUploadContainer main-content">
          <div className="content-wrapper">
            <div className="notificationContainer" style={{ display: showNotification ? 'block' : 'none' }}>
              <MessageBar
                messageBarType={notificationType}
                isMultiline={false}
                onDismiss={() => this.closeNotification()}
                dismissButtonAriaLabel="Close"
              >
                {notificationText}
              </MessageBar>
            </div>
            <div className='headerContainer'>
              <Label>
                {LABEL_TEXT_FILE_UPLOAD}
              </Label>
            </div>
            <div className='formContainer'>
              <div className="">
                <Stack style={{width: "50%"}}>
                  <StackItem>
                    <Stack>
                      <StackItem>
                        <Label>
                          {LABEL_TEXT_FILE}
                        </Label>
                      </StackItem>
                      <StackItem>
                        <input type="file"
                          id="fileUploadCtrl"
                          className="form-control"
                          onChange={(e) => this.handleFileChange(e)} />
                      </StackItem>
                    </Stack>
                    <Stack>
                      <StackItem>
                        {LABEL_TEXT_FOLDER}
                      </StackItem>
                      <StackItem>
                        <Dropdown
                          selectedKey={selectedAgency}
                          options={agencyOptions}
                          onChange={(e, option) => this.handleAgencyChange(e, option)} />
                      </StackItem>
                    </Stack>
                  </StackItem>
                </Stack>
                <Stack styles={stackStyleBtnRow} horizontal>
                  <StackItem>
                    <div style={{ position: 'relative', top: '8px', right: '10px', display: inProgress ? 'block' : 'none' }}>
                      <Spinner size={SpinnerSize.small} />
                    </div>
                  </StackItem>
                  <StackItem>
                    <DefaultButton
                      text={BTN_TEXT_RESET}
                      onClick={() => this.resetControls()}
                      style={{ marginRight: '20px' }} />
                    <PrimaryButton
                      text={BTN_TEXT_UPLOAD}
                      disabled={btnDisabled || inProgress}
                      onClick={() => this.handleUpload()}
                      style={{ marginRight: '20px' }} />
                  </StackItem>
                </Stack>
              </div>
            </div>
          </div>
        </div>
        {
          showMetaDataForm &&
          <MetaDataForm
            Domain={metaDataAgency}
            Category={metaDataCategory}
            DomainOptions={domainOptions}
            CategoryOptions={categoryOptions}
            PublishedDate={this.state.publishedDate}
            FileID={this.state.fileID}
            Context={this.props.context}
            SiteRelativeUrl={this.props.siteRelativeUrl}
            SPService={this._spService}
            IsModalOpen={showMetaDataForm}
            OnAgencyChange={this.handleMetaDataAgencyChange}
            OnCategoryChange={this.handleMetaDataCategoryChange}
            OnPublishedDateChange={this.handlePublishedDateChange}
            CloseModal={this.closeMetaDataForm}
            UpdateContext={this.updateContextMetaDataForm}
          />
        }
      </section>
    );
  }
}
