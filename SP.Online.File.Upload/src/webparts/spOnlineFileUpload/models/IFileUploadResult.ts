import { Result } from "../common/constants";

export interface IFileUploadResult {
    result: Result;
    fileID: any;
    fileUrl: string;
    error?: string;
}

export interface IFileUploadResponse {
    isFileExists: boolean;
    fileUpload: string;
    fileName: string;
    fileID: number;
}