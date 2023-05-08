import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as XLSX from 'xlsx';

export class ImportExcelV2 implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private _excelUploadinput: HTMLInputElement;
    private _paragraphinput: HTMLLabelElement;
    private _backloginput: HTMLLabelElement;
    private _notifyOutputChanged: () => void;
    private _container: HTMLDivElement;

    /**
     * Empty constructor.
     */
    constructor() {

    }

    /**
     * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
     * Data-set values are not initialized here, use updateView.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
     * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
     * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
     * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
     */
    public init(context: ComponentFramework.Context<IInputs>, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary, container: HTMLDivElement) {
        // Add control initialization code
        this._excelUploadinput = document.createElement("input");
        this._excelUploadinput.id = "fileUploader";
        this._excelUploadinput.type = "file";
        this._excelUploadinput.name = "fileUploader";
        this._excelUploadinput.accept = ".xls, .xlsx";
        this._excelUploadinput.style.opacity = "1";
        this._excelUploadinput.style.width = "auto";
        this._excelUploadinput.style.height = "auto";
        this._excelUploadinput.style.pointerEvents = "all";

        this._notifyOutputChanged = notifyOutputChanged;
        //this.button.addEventListener("click", (event) => { this._value = this._value + 1; this._notifyOutputChanged();});
        this._excelUploadinput.addEventListener("change", this.excelupdated.bind(this));
        this._excelUploadinput.addEventListener("click", this.excelupdated.bind(this));
        this._container = document.createElement("div");
        this._container.appendChild(this._excelUploadinput);

        this._paragraphinput = document.createElement("label");
        this._paragraphinput.id = "jsonOrder";
        this._paragraphinput.style.display = "none";

        this._backloginput = document.createElement("label");
        this._backloginput.id = "jsonBacklog";
        this._backloginput.style.display = "none";

        this._container.appendChild(this._paragraphinput);
        this._container.appendChild(this._backloginput);
        container.appendChild(this._container);
    }


    /**
     * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
     * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
     */
    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Add code to update control view
    }


    private excelupdated(event: Event): void {
        this._excelUploadinput.addEventListener('change', function(evt) {
            let selectedFile = (<HTMLInputElement>document.getElementById('fileUploader')).files[0];
            // 如果没有选择任何文件，清空jsonOrder并退出
            if (!selectedFile) {
                document.getElementById("jsonOrder").innerHTML = "";
                document.getElementById("jsonBacklog").innerHTML = "";
                return;
            }

            let reader = new FileReader();

            reader.onload = function(event) {
                let data = new Uint8Array(event.target.result as ArrayBuffer);
                let workbook = XLSX.read(data, {
                    type: 'array'
                });
                let firstSheetName=workbook.SheetNames[0];
                let XL_row_object: any = XLSX.utils.sheet_to_json(workbook.Sheets[firstSheetName],{ header: undefined });
    
                
                //提取Backlog
                interface RowData {
                    ProductId: string;
                    Backlog: number;
                    MaterialId:string;
                    CustomerId:string;
                    [key: string]: any;
                  }

                const transformedBacklog = XL_row_object.map((row:RowData) => {
                    return {
                      ProductId: row.ProductId,
                      Backlog: row.Backlog,
                      MaterialId:row.MaterialId,
                      CustomerId:row.CustomerId 
                    };
                  });
                console.log(transformedBacklog);
                let json_Backlog = JSON.stringify(transformedBacklog);
                document.getElementById("jsonBacklog").innerHTML = json_Backlog;


                //转换数据格式
                const transformedData = [];
                // 循环遍历源数据的每一行 
                for (const row of XL_row_object) {
                    // 循环遍历每个日期列
                    for (const dateColumn in row) {
                    // 排除ProductId列
                    if (dateColumn !== "ProductId" && dateColumn !== "CustomerId" && dateColumn !== "MaterialId" && dateColumn !== "Backlog") {
                        // 将每个日期列的数据添加到转换后的数组中
                        transformedData.push({
                        ProductId: row.ProductId,
                        MaterialId: row.MaterialId,
                        CustomerId: row.CustomerId,
                        Date: dateColumn,
                        QTY: row[dateColumn]
                        });
                    }
                    }
                }
                //console.log(transformedData);
                let json_object = JSON.stringify(transformedData);
                document.getElementById("jsonOrder").innerHTML = json_object;
            };
    
            reader.onerror = function(event) {
                console.error("File could not be read! Code " + event.target.error.code);
            };
    
            reader.readAsArrayBuffer(selectedFile);
            //reader.readAsBinaryString(selectedFile);
        });
        this._notifyOutputChanged();
    }
    
    /** 
     * It is called by the framework prior to a control receiving new data. 
     * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
     */
    public getOutputs(): IOutputs {
        return {
            Output: document.getElementById("jsonOrder").innerHTML,
            Backlog:document.getElementById("jsonBacklog").innerHTML
        };
    }

    /** 
     * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
     * i.e. cancelling any pending remote calls, removing listeners, etc.
     */
    public destroy(): void {
        // Add code to cleanup control if necessary
    }
}