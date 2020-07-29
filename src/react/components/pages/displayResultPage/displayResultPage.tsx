// Copyright (c) Microsoft Corporation.
// Licensed under the MIT license.

import React, { RefObject } from "react";
import SplitPane from "react-split-pane";
import EditorSideBar from "../editorPage/editorSideBar";
import { AzureBlobStorage } from "../../../../providers/storage/azureBlobStorage";
import Form, { Widget, IChangeEvent, FormValidation } from "react-jsonschema-form";
import LocalFolderPicker from "../../common/localFolderPicker/localFolderPicker";
import CustomFieldTemplate from "../../common/customField/customFieldTemplate";
import ConnectionProviderPicker from "../../common/connectionProviderPicker/connectionProviderPicker";
import Checkbox from "rc-checkbox";
import { ProtectedInput } from "../../common/protectedInput/protectedInput";
import { CustomWidget } from "../../common/customField/customField";
import { connect } from "react-redux";
import { RouteComponentProps } from "react-router-dom";
import { bindActionCreators } from "redux";
import { FontIcon, PrimaryButton, Spinner, SpinnerSize, IconButton, TextField, IDropdownOption, Dropdown} from "@fluentui/react";
import IProjectActions, * as projectActions from "../../../../redux/actions/projectActions";
import IApplicationActions, * as applicationActions from "../../../../redux/actions/applicationActions";
import IAppTitleActions, * as appTitleActions from "../../../../redux/actions/appTitleActions";
import "./displayResultPage.scss";
import {
    IApplicationState, IConnection, IProject, IAppSettings, AppError, ErrorCode,
    AssetState, AssetType, EditorMode,
    IAsset, IAssetMetadata, IRegion,
    ISize, ITag,
    ILabel,
    FieldType,
    FieldFormat,
} from "../../../../models/applicationState";
import { ImageMap } from "../../common/imageMap/imageMap";
import Style from "ol/style/Style";
import Stroke from "ol/style/Stroke";
import Fill from "ol/style/Fill";
import PredictResult from "./predictResult";
import _ from "lodash";
import pdfjsLib from "pdfjs-dist";
import Alert from "../../common/alert/alert";
import url from "url";
import HtmlFileReader from "../../../../common/htmlFileReader";
import { Feature } from "ol";
import Polygon from "ol/geom/Polygon";
import { strings, interpolate, addLocValues } from "../../../../common/strings";
import PreventLeaving from "../../common/preventLeaving/preventLeaving";
import ServiceHelper from "../../../../services/serviceHelper";
import { parseTiffData, renderTiffToCanvas, loadImageToCanvas } from "../../../../common/utils";
import { constants } from "../../../../common/constants";
import { getPrimaryGreenTheme, getPrimaryWhiteTheme,
         getGreenWithWhiteBackgroundTheme } from "../../../../common/themes";
import axios from "axios";
import { IAnalyzeModelInfo } from './predictResult';
import {AssetPreview, ContentSource} from "../../common/assetPreview/assetPreview";
import Canvas from "../editorPage/canvas";
import { InvoiceResultService } from "../../../../services/invoiceResultService";
// tslint:disable-next-line:no-var-requires
const formSchema = addLocValues(require("./connectionForm.json"));
// tslint:disable-next-line:no-var-requires
const uiSchema = addLocValues(require("./connectionForm.ui.json"));

pdfjsLib.GlobalWorkerOptions.workerSrc = constants.pdfjsWorkerSrc(pdfjsLib.version);
const cMapUrl = constants.pdfjsCMapUrl(pdfjsLib.version);

export interface IDisplayResultPageProps extends RouteComponentProps, React.Props<DisplayResultPage> {
    // recentProjects: IProject[];
    connections: IConnection[];
    appSettings: IAppSettings;
    project: IProject;
    actions: IProjectActions;
    applicationActions: IApplicationActions;
    appTitleActions: IAppTitleActions;
    connection: IConnection;
}

export interface IDisplayResultPageState {
    isFetching: boolean;
    fileLabel: string;
    currentPage: number;
    imageUri: string;
    imageWidth: number;
    imageHeight: number;
    shouldShowAlert: boolean;
    alertTitle: string;
    alertMessage: string;
    highlightedField: string;
    formSchema: any;
    uiSchema: any;
    classNames: string[];
    assets: IAsset[];
    selectedAsset?: IAsset;
    thumbnailSize: ISize;
    numPages: number;
    tiffImages: any[];
    pdfFile: any;
    invoiceAnalyzeResult: {};
    invoiceResultForCurrentPage: any;
    predictRun: boolean;
}

export interface IModel {
    modelId: string;
    createdDateTime: string;
    lastUpdatedDateTime: string;
    status: string;
}

function mapStateToProps(state: IApplicationState) {
    return {
        recentProjects: state.recentProjects,
        connections: state.connections,
        appSettings: state.appSettings,
    };
}

function mapDispatchToProps(dispatch) {
    return {
        actions: bindActionCreators(projectActions, dispatch),
        applicationActions: bindActionCreators(applicationActions, dispatch),
        appTitleActions: bindActionCreators(appTitleActions, dispatch),
    };
}

@connect(mapStateToProps, mapDispatchToProps)
export default class DisplayResultPage extends React.Component<IDisplayResultPageProps, IDisplayResultPageState> {
    private widgets = {
        localFolderPicker: (LocalFolderPicker as any) as Widget,
        connectionProviderPicker: (ConnectionProviderPicker as any) as Widget,
        protectedInput: (ProtectedInput as any) as Widget,
        checkbox: CustomWidget(Checkbox, (props) => ({
            checked: props.value,
            onChange: (value) => props.onChange(value.target.checked),
            disabled: props.disabled,
        })),
    };

    public state: IDisplayResultPageState = {
        isFetching: false,
        fileLabel: "",
        currentPage: undefined,
        imageUri: null,
        imageWidth: 0,
        imageHeight: 0,
        shouldShowAlert: false,
        alertTitle: "",
        alertMessage: "",
        highlightedField: "",
        classNames: ["needs-validation"],
        formSchema: { ...formSchema },
        uiSchema: { ...uiSchema },
        assets: [],
        thumbnailSize: { width: 175, height: 155 },
        numPages: 1,
        tiffImages: [],
        pdfFile: null,
        invoiceAnalyzeResult: null,
        invoiceResultForCurrentPage: {},
        predictRun: false,
    };

    private imageMap: ImageMap;
    private azureBlobStorage: AzureBlobStorage;
    private folderPath: string;
    private invoiceResultService: InvoiceResultService;
    private tags: ITag[] = [
        {
            name: "InvoiceDate",
            color: "#CC543A",
            type: FieldType.String,
            format: FieldFormat.NotSpecified
        },
        {
            name: "InvoiceNumber",
            color: "#7BA23F",
            type: FieldType.String,
            format: FieldFormat.NotSpecified
        },
        {
            name: "VendorAddress",
            color: "#58B2DC",
            type: FieldType.String,
            format: FieldFormat.NotSpecified
        },
        {
            name: "VendorName",
            color: "#FFB11B",
            type: FieldType.String,
            format: FieldFormat.NotSpecified
        },
        {
            name: "CustomerAddress",
            color: "#2E5C6E",
            type: FieldType.String,
            format: FieldFormat.NotSpecified
        },
        {
            name: "TotalInvoiceAmount",
            color: "#A96360",
            type: FieldType.String,
            format: FieldFormat.NotSpecified
        },
        {
            name: "CustomerId",
            color: "#D7C4BB",
            type: FieldType.String,
            format: FieldFormat.NotSpecified
        },
        {
            name: "SubTotal",
            color: "#8F77B5",
            type: FieldType.String,
            format: FieldFormat.NotSpecified
        },
        {
            name: "TotalTax",
            color: "#EEA9A9",
            type: FieldType.String,
            format: FieldFormat.NotSpecified
        },
        {
            name: "CustomerName",
            color: "#24936E",
            type: FieldType.String,
            format: FieldFormat.NotSpecified
        },
        {
            name: "BillingAddress",
            color: "#994639",
            type: FieldType.String,
            format: FieldFormat.NotSpecified
        },
        {
            name: "DueDate",
            color: "#BEC23F",
            type: FieldType.String,
            format: FieldFormat.NotSpecified
        },
        {
            name: "ShippingAddress",
            color: "#26453D",
            type: FieldType.String,
            format: FieldFormat.NotSpecified
        }
    ]

    public async componentDidMount() {
        window.addEventListener("focus", this.onFocused);

        document.title = strings.displayResult.title + " - " + strings.appName;
        console.log("document.title: " + document.title);
    }

    public componentWillUnmount() {
        window.removeEventListener("focus", this.onFocused);
    }

    private onFocused = () => {
        // This approach will remove the entered sas everytime swich window, need to figure a better solution
        // this.loadInvoiceAssets();
    }

    public componentDidUpdate(prevProps: Readonly<IDisplayResultPageProps>, prevState: Readonly<IDisplayResultPageState>) {
        // Handles asset changing
        if (!this.state.selectedAsset && !prevState.selectedAsset)
        {
            return;
        }
        if ((this.state.selectedAsset && !prevState.selectedAsset) ||
            (this.state.selectedAsset.name !== prevState.selectedAsset.name)) {
            this.imageMap.removeAllFeatures();
            this.setState({
                numPages: 1,
                currentPage: 1,
                pdfFile: null,
                imageUri: null,
                tiffImages: [],
            }, async () => {
                await this.loadImage();
                // await this.loadOcr();
                // this.loadLabelData(asset);
                await this.drawInvoicePredictionResult(this.state.selectedAsset);
            });
        } else if (prevState.currentPage !== this.state.currentPage) {
            this.imageMap.removeAllFeatures();
            if (this.state.pdfFile !== null) {
                this.loadPdfPage(this.state.selectedAsset.id, this.state.pdfFile, this.state.currentPage);
                this.drawInvoicePredictionResult(this.state.selectedAsset);
            } else if (this.state.tiffImages.length !== 0) {
                this.loadTiffPage(this.state.tiffImages, this.state.currentPage);
                this.drawInvoicePredictionResult(this.state.selectedAsset);
            }
        }

        if (prevState.highlightedField !== this.state.highlightedField) {
            this.setPredictedFieldHighlightStatus(this.state.highlightedField);
        }
    }

    public render() {
        console.log("displayResultPage render() gets called");
        const predictions = this.getPredictionsFromAnalyzeResult(this.state.invoiceAnalyzeResult);
        const modelInfo: IAnalyzeModelInfo = this.getAnalyzeModelInfo(this.state.invoiceAnalyzeResult);
        const selectedAsset = this.state.selectedAsset;
        const assets = this.state.assets;

        return (
            <div className="editor-page skipToMainContent" id="pageEditor">
                <SplitPane split="vertical"
                    defaultSize={this.state.thumbnailSize.width}
                    minSize={150}
                    maxSize={325}
                    paneStyle={{ display: "flex" }}
                    // onChange={this.onSideBarResize}
                    // onDragFinished={this.onSideBarResizeComplete}
                >
                    <div className="editor-page-sidebar bg-lighter-1">
                        <EditorSideBar
                            assets={assets}
                            selectedAsset={selectedAsset ? selectedAsset : null}
                            onBeforeAssetSelected={this.onBeforeAssetSelected}
                            onAssetSelected={this.onAssetSelected}
                            onAssetLoaded={this.onAssetLoaded}
                            thumbnailSize={this.state.thumbnailSize}
                        />
                    </div>
                    <div className="editor-page-content">
                        <SplitPane split = "vertical"
                            primary = "second"
                            maxSize = {625}
                            minSize = {290}
                            pane1Style = {{height: "100%"}}
                            pane2Style = {{height: "auto"}}
                            resizerStyle = {{width: "5px", margin: "0px", border: "2px", background: "transparent"}}
                            // onChange = {() => this.resizeCanvas()}
                        >
                            <div className="editor-page-content-main" >
                                <div className="editor-page-content-main-body">
                                    {this.state.selectedAsset && this.renderImageMap()}
                                    {this.renderPrevPageButton()}
                                    {this.renderNextPageButton()}
                                    { this.shouldShowMultiPageIndicator() &&
                                        <p className="page-number">
                                            Page {this.state.currentPage} of {this.state.numPages}
                                        </p>
                                    }
                                </div>
                            </div>
                            <div className="editor-page-right-sidebar">
                                <div className="condensed-list">
                                    <h6 className="condensed-list-header bg-darker-2 p-2 flex-center">
                                        <FontIcon className="mr-1" iconName="Insights" />
                                        <span>Display Results</span>
                                    </h6>
                                    <div className="p-3">
                                        <div className="container-space-between">
                                            <Form
                                                className={this.state.classNames.join(" ")}
                                                showErrorList={false}
                                                liveValidate={true}
                                                noHtml5Validate={true}
                                                FieldTemplate={CustomFieldTemplate}
                                                widgets={this.widgets}
                                                schema={this.state.formSchema}
                                                uiSchema={this.state.uiSchema}
                                                onSubmit={form => this.onFormSubmit(form.formData)}>
                                                <div>
                                                    <PrimaryButton
                                                        theme={getPrimaryGreenTheme()}
                                                        className="mr-2"
                                                        type="submit"
                                                        text = "Render"
                                                    />
                                                </div>
                                            </Form>
                                        </div>
                                        <div className="alight-vertical-center mt-2">
                                            <div className="seperator"/>
                                        </div>
                                        {this.state.isFetching &&
                                            <div className="loading-container">
                                                <Spinner
                                                    label="Fetching..."
                                                    ariaLive="assertive"
                                                    labelPosition="right"
                                                    size={SpinnerSize.large}
                                                />
                                            </div>
                                        }
                                        {Object.keys(predictions).length > 0 &&
                                            <PredictResult
                                                predictions={predictions}
                                                analyzeResult={this.state.invoiceAnalyzeResult}
                                                analyzeModelInfo={modelInfo}
                                                page={this.state.currentPage}
                                                tags={this.tags}
                                                downloadResultLabel={this.state.fileLabel}
                                                onPredictionClick={this.onPredictionClick}
                                                onPredictionMouseEnter={this.onPredictionMouseEnter}
                                                onPredictionMouseLeave={this.onPredictionMouseLeave}
                                            />
                                        }
                                        {
                                            (Object.keys(predictions).length === 0 && this.state.predictRun) &&
                                            <div>
                                                No field can be extracted.

                                            </div>
                                        }
                                    </div>
                                </div>
                            </div>
                        </SplitPane>
                    </div>
                </SplitPane>
                <Alert
                    show={this.state.shouldShowAlert}
                    title={this.state.alertTitle}
                    message={this.state.alertMessage}
                    onClose={() => this.setState({
                        shouldShowAlert: false,
                        alertTitle: "",
                        alertMessage: "",
                    })}
                />
            </div>
        );
    }

    private shouldShowMultiPageIndicator = () => {
        return (this.state.pdfFile !== null || this.state.tiffImages.length !== 0) && this.state.numPages > 1;
    }

    private onFormSubmit = async (formData:any) => {
        console.log("displayResultPage onFormSubmit get called");
        var sas = formData["sas"].trim();
        var folderPath = formData["folderPath"];
        console.log("init assets.length: " + this.state.assets.length);
        this.getPrediction(sas, folderPath)
            .then((result) => {
                console.log("after getPrediction assets.length: " + this.state.assets.length);
            }).catch((error) => {
                let alertMessage = "";
                if (error.response) {
                    alertMessage = error.response.data;
                } else if (error.errorCode === ErrorCode.PredictWithoutTrainForbidden) {
                    alertMessage = strings.errors.predictWithoutTrainForbidden.message;
                } else if (error.errorCode === ErrorCode.ModelNotFound) {
                    alertMessage = error.message;
                } else {
                    alertMessage = interpolate(strings.errors.endpointConnectionError.message, { endpoint: "form recognizer backend URL" });
                }
                this.setState({
                    shouldShowAlert: true,
                    alertTitle: "Prediction Failed",
                    alertMessage
                });
            });
    }

    private onBeforeAssetSelected = (): boolean => {
        console.log("displayResultPage onBeforeAssetSelected gets called");

        return true;
    }

    private onAssetSelected = async (asset: IAsset): Promise<void> => {
        console.log("displayResultPage onAssetSelected gets called name: " + asset.name);
        // Nothing to do if we are already on the same asset.
        if (this.state.selectedAsset && this.state.selectedAsset.id === asset.id) {
            return;
        }

        this.setState({ selectedAsset: asset });
    }

    private onAssetLoaded = (asset: IAsset, contentSource: ContentSource) => {
        console.log("displayResultPage onAssetLoaded gets called");
        const assets = [...this.state.assets];
        const assetIndex = assets.findIndex((item) => item.id === asset.id);
        if (assetIndex > -1) {
            const assets = [...this.state.assets];
            const item = {...assets[assetIndex]};
            item.cachedImage = (contentSource as HTMLImageElement).src;
            assets[assetIndex] = item;
            this.setState({
                assets: assets,
                currentPage: 1,
            });
        }
    }


    private renderPrevPageButton = () => {
        if (!this.state.selectedAsset) {
            return <div></div>;
        }
        const prevPage = () => {
            this.setState((prevState) => ({
                currentPage: Math.max(1, prevState.currentPage - 1),
            }));
        };

        if (this.state.currentPage > 1) {
            return (
                <IconButton
                    className="toolbar-btn prev"
                    title="Previous"
                    iconProps={{iconName: "ChevronLeft"}}
                    onClick={prevPage}
                />
            );
        } else {
            return <div></div>;
        }
    }

    private renderNextPageButton = () => {
        if (!this.state.selectedAsset) {
            return <div></div>;
        }

        const nextPage = () => {
            this.setState((prevState) => ({
                currentPage: Math.min(prevState.currentPage + 1, this.state.numPages),
            }));
        };

        if (this.state.currentPage < this.state.numPages) {
            return (
                <IconButton
                    className="toolbar-btn next"
                    title="Next"
                    onClick={nextPage}
                    iconProps={{iconName: "ChevronRight"}}
                />
            );
        } else {
            return <div></div>;
        }
    }

    private renderImageMap = () => {
        console.log("displayResultPage renderImageMap gets called");

        return (
            <ImageMap
                ref={(ref) => this.imageMap = ref}
                imageUri={this.state.imageUri || ""}
                imageWidth={this.state.imageWidth}
                imageHeight={this.state.imageHeight}

                featureStyler={this.featureStyler}
                onMapReady={this.noOp}
            />
        );
    }

    private async getPrediction(sas: string, folderPath: string): Promise<any> {
        console.log("getPrediction get called, past in sas: " + sas + " folderPath: " + folderPath);
        
        let storageOption = { sas: sas };
        this.folderPath = folderPath;
        this.azureBlobStorage = new AzureBlobStorage(storageOption);
        this.invoiceResultService = new InvoiceResultService(this.azureBlobStorage);
        await this.loadInvoiceAssets();
        console.log("getPrediction this.state.assets.length: " + this.state.assets.length);
        console.log("displayResultPage getPrediction first asset name: " + this.state.assets[0].name);
        console.log("displayResultPage getPrediction first asset path: " + this.state.assets[0].path);
    }

    private async loadInvoiceAssets(): Promise<any> {
        console.log("loadInvoiceAssets get called, this.azureBlobStorage: " + this.azureBlobStorage + " this.folderPath: " + this.folderPath);
        if (!this.azureBlobStorage) {
            return;
        }
        const assets = await this.azureBlobStorage.getInvoicePredictionAssets(this.folderPath);
        console.log("displayResultPage getPrediction assets.length: " + assets.length);
        const verifiedAssets = assets.map((asset) => {
            asset.name = decodeURIComponent(asset.name);
            return asset;
        }).filter((asset) => this.isInExactFolderPath(asset.name, this.folderPath));
        console.log("displayResultPage loadInvoiceAssets verifiedAssets.length: " + verifiedAssets.length);
        this.setState({assets: verifiedAssets});
    }

    private isInExactFolderPath = (assetName: string, normalizedPath: string): boolean => {
        if (normalizedPath === "") {
            return assetName.lastIndexOf("/") === -1;
        }

        const startsWithFolderPath = assetName.indexOf(`${normalizedPath}/`) === 0;
        return startsWithFolderPath && assetName.lastIndexOf("/") === normalizedPath.length;
    }

    private loadImage = async () => {
        const asset = this.state.selectedAsset;
        if (asset.type === AssetType.Image) {
            const canvas = await loadImageToCanvas(asset.path);
            this.setState({
                imageUri: canvas.toDataURL(constants.convertedImageFormat, constants.convertedImageQuality),
                imageWidth: canvas.width,
                imageHeight: canvas.height,
            });
        } else if (asset.type === AssetType.TIFF) {
            await this.loadTiffFile(asset);
        } else if (asset.type === AssetType.PDF) {
            await this.loadPdfFile(asset.id, asset.path);
        }
    }

    private loadTiffFile = async (asset: IAsset) => {
        const assetArrayBuffer = await HtmlFileReader.getAssetArray(asset);
        const tiffImages = parseTiffData(assetArrayBuffer);
        this.loadTiffPage(tiffImages, this.state.currentPage);
    }

    private loadTiffPage = (tiffImages: any[], pageNumber: number) => {
        const tiffImage = tiffImages[pageNumber - 1];
        const canvas = renderTiffToCanvas(tiffImage);
        this.setState({
            imageUri: canvas.toDataURL(constants.convertedImageFormat, constants.convertedImageQuality),
            imageWidth: tiffImage.width,
            imageHeight: tiffImage.height,
            numPages: tiffImages.length,
            currentPage: pageNumber,
            tiffImages,
        });
    }

    private loadPdfFile = async (assetId, url) => {
        try {
            const pdf = await pdfjsLib.getDocument({url, cMapUrl, cMapPacked: true}).promise;
            // Fetch current page
            if (assetId === this.state.selectedAsset.id) {
                await this.loadPdfPage(assetId, pdf, this.state.currentPage);
            }
        } catch (reason) {
            // PDF loading error
            console.error(reason);
        }
    }

    private loadPdfPage = async (assetId, pdf, pageNumber) => {
        const page = await pdf.getPage(pageNumber);
        const defaultScale = 2;
        const viewport = page.getViewport({ scale: defaultScale });

        // Prepare canvas using PDF page dimensions
        const canvas = document.createElement("canvas");
        const context = canvas.getContext("2d");
        canvas.height = viewport.height;
        canvas.width = viewport.width;

        // Render PDF page into canvas context
        const renderContext = {
            canvasContext: context,
            viewport,
        };

        await page.render(renderContext).promise;
        if (assetId === this.state.selectedAsset.id) {
            this.setState({
                imageUri: canvas.toDataURL(constants.convertedImageFormat, constants.convertedImageQuality),
                imageWidth: canvas.width,
                imageHeight: canvas.height,
                numPages: pdf.numPages,
                currentPage: pageNumber,
                pdfFile: pdf,
            });
        }
    }


    private createBoundingBoxVectorFeature = (text, boundingBox, imageExtent, ocrExtent) => {
        const coordinates: number[][] = [];

        // extent is int[4] to represent image dimentions: [left, bottom, right, top]
        const imageWidth = imageExtent[2] - imageExtent[0];
        const imageHeight = imageExtent[3] - imageExtent[1];
        const ocrWidth = ocrExtent[2] - ocrExtent[0];
        const ocrHeight = ocrExtent[3] - ocrExtent[1];

        for (let i = 0; i < boundingBox.length; i += 2) {
            coordinates.push([
                Math.round((boundingBox[i] / ocrWidth) * imageWidth),
                Math.round((1 - (boundingBox[i + 1] / ocrHeight)) * imageHeight),
            ]);
        }

        const feature = new Feature({
            geometry: new Polygon([coordinates]),
        });
        const tag = this.tags.find((tag) => tag.name.toLocaleLowerCase() === text.toLocaleLowerCase());
        const isHighlighted = (text.toLocaleLowerCase() === this.state.highlightedField.toLocaleLowerCase());
        feature.setProperties({
            color: _.get(tag, "color", "#ff0000"),
            fieldName: text,
            isHighlighted,
        });

        return feature;
    }

    private featureStyler = (feature) => {
        return new Style({
            stroke: new Stroke({
                color: feature.get("color"),
                width: feature.get("isHighlighted") ? 4 : 2,
            }),
            fill: new Fill({
                color: "rgba(255, 255, 255, 0)",
            }),
        });
    }

    private drawInvoicePredictionResult = async (asset: IAsset) => {
        try {
            const invoiceAnalyzeResult = await this.invoiceResultService.getRecognizedText(asset.path, asset.name);
            this.setState({
                invoiceAnalyzeResult,
                predictRun: true,
            }, () => {
                const features = [];
                const imageExtent = [0, 0, this.state.imageWidth, this.state.imageHeight];
                const ocrForCurrentPage: any = this.getOcrFromAnalyzeResult(this.state.invoiceAnalyzeResult)[this.state.currentPage - 1];
                const ocrExtent = [0, 0, ocrForCurrentPage.width, ocrForCurrentPage.height];
                const predictions = this.getPredictionsFromAnalyzeResult(this.state.invoiceAnalyzeResult);

                for (const fieldName of Object.keys(predictions)) {
                    const field = predictions[fieldName];
                    if (_.get(field, "page", null) === this.state.currentPage) {
                        const text = fieldName;
                        const boundingbox = _.get(field, "boundingBox", []);
                        const feature = this.createBoundingBoxVectorFeature(text, boundingbox, imageExtent, ocrExtent);
                        features.push(feature);
                    }
                }
                this.imageMap.addFeatures(features);
            });
        } catch (error) {
            console.log(error);
            this.setState({
                // isError: true,
                // errorTitle: error.title,
                // errorMessage: error.message,
            });
        }
    }

    private getPredictionsFromAnalyzeResult(analyzeResult: any) {
        return _.get(analyzeResult, "analyzeResult.documentResults[0].fields", {});
    }

    private getAnalyzeModelInfo(analyzeResult) {
        const { modelId, docType, docTypeConfidence } = _.get(analyzeResult, "analyzeResult.documentResults[0]", {})
        return { modelId, docType, docTypeConfidence };
    }

    private getOcrFromAnalyzeResult(analyzeResult: any) {
        return _.get(analyzeResult, "analyzeResult.readResults", []);
    }

    private createObjectURL = (object: File) => {
        // generate a URL for the object
        return (window.URL) ? window.URL.createObjectURL(object) : "";
    }

    private noOp = () => {
        // no operation
    }

    private onPredictionClick = (predictedItem: any) => {
        const targetPage = predictedItem.page;
        if (Number.isInteger(targetPage) && targetPage !== this.state.currentPage) {
            this.setState({
                currentPage: targetPage,
                highlightedField: predictedItem.fieldName ?? "",
            });
        }
    }

    private onPredictionMouseEnter = (predictedItem: any) => {
        this.setState({
            highlightedField: predictedItem.fieldName ?? "",
        });
    }

    private onPredictionMouseLeave = (predictedItem: any) => {
        this.setState({
            highlightedField: "",
        });
    }

    private setPredictedFieldHighlightStatus = (highlightedField: string) => {
        const features = this.imageMap.getAllFeatures();
        for (const feature of features) {
            if (feature.get("fieldName").toLocaleLowerCase() === highlightedField.toLocaleLowerCase()) {
                feature.set("isHighlighted", true);
            } else {
                feature.set("isHighlighted", false);
            }
        }
    }
}
