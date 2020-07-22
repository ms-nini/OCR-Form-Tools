// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import Guard from "../common/guard";
import { IProject } from "../models/applicationState";
import { IStorageProvider, StorageProviderFactory } from "../providers/storage/storageProviderFactory";
import { constants } from "../common/constants";
import ServiceHelper from "./serviceHelper";

// export enum InvoiceResultStatus {
//     loadingFromAzureBlob,
//     done,
// }

/**
 * @name - Invoice Prediction Result Service
 * @description - Functions for dealing with Invoice Prediction Result
 */
export class InvoiceResultService {

    constructor(private storageProvider: IStorageProvider) {
        Guard.null(storageProvider);
    }

    /**
     * get recognized text from Invoice Result service
     * @param filePath - filepath sent to Invoice Result
     * @param fileName - name of Invoice Result file
     */
    public async getRecognizedText(
        filePath: string,
        fileName: string,
        // onStatusChanged?: (predictionResultStatus: InvoiceResultStatus) => void,
        // rewrite?: boolean
    ): Promise<any> {
        Guard.empty(filePath);

        // const notifyStatusChanged = (invoiceResultStatus: InvoiceResultStatus) => onStatusChanged && onStatusChanged(invoiceResultStatus);
        const invoiceResultFileName = decodeURIComponent(`${fileName}${constants.invoiceFileExtension}`);

        let invoiceResultJson;
        // notifyStatusChanged(InvoiceResultStatus.loadingFromAzureBlob);
        invoiceResultJson = await this.readInvoiceResultFile(invoiceResultFileName);
        return invoiceResultJson;
    }

    private readInvoiceResultFile = async (invoiceResultFileName: string) => {
        const json = await this.storageProvider.readText(invoiceResultFileName, true);
        if (json !== null) {
            return new Promise((resolve, reject) => {
                resolve(JSON.parse(json));
            });
        }
    }

    /**
     * Poll function to repeatly check if request succeeded
     * @param func - function that will be called repeatly
     * @param timeout - timeout
     * @param interval - interval
     */
    private poll = (func, timeout, interval): Promise<any> => {
        const endTime = Number(new Date()) + (timeout || 10000);
        interval = interval || 100;

        const checkSucceeded = (resolve, reject) => {
            const ajax = func();
            ajax.then((response) => {
                if (response.data.status.toLowerCase() === constants.statusCodeSucceeded) {
                    resolve(response.data);
                } else if (Number(new Date()) < endTime) {
                    // If the request isn't succeeded and the timeout hasn't elapsed, go again
                    setTimeout(checkSucceeded, interval, resolve, reject);
                } else {
                    // Didn't succeeded after too much time, reject
                    reject(new Error("Timed out for getting Invoice results"));
                }
            });
        };

        return new Promise(checkSucceeded);
    }
}
