import { IInputs, IOutputs } from "./generated/ManifestTypes";

export class ExcelUpload implements ComponentFramework.StandardControl<IInputs, IOutputs> {
    private container: HTMLDivElement;
    private fileInput: HTMLInputElement;
    private selectedFileDiv: HTMLDivElement;
    private fileNameSpan: HTMLSpanElement;
    private uploadBtn: HTMLButtonElement;
    private notifyOutputChanged: () => void;
    private selectedFile: File | null = null;
    private encodedFileData: string = "";
    private uploadStatus: string = "none";

    constructor() { }

    public init(
        context: ComponentFramework.Context<IInputs>,
        notifyOutputChanged: () => void,
        state: ComponentFramework.Dictionary,
        container: HTMLDivElement
    ): void {
        this.container = container;
        this.notifyOutputChanged = notifyOutputChanged;

        // Create the entire UI
        this.container.innerHTML = `
            <style>
                .excel-upload-container {
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                    display: flex;
                    align-items: center;
                    justify-content: center;
                    background: linear-gradient(135deg, #742774 0%, #5A1E5C 50%, #4B4B9D 100%);
                    padding: 20px;
                    min-height: 400px;
                }

                .excel-card {
                    background-color: white;
                    border-radius: 8px;
                    box-shadow: 0 10px 25px rgba(0,0,0,0.2);
                    padding: 32px;
                    width: 100%;
                    max-width: 448px;
                }

                .excel-header {
                    text-align: center;
                    margin-bottom: 32px;
                }

                .excel-icon-circle {
                    display: inline-flex;
                    align-items: center;
                    justify-content: center;
                    width: 64px;
                    height: 64px;
                    background-color: #F3E5F5;
                    border-radius: 50%;
                    margin-bottom: 16px;
                }

                .excel-icon-circle svg {
                    width: 32px;
                    height: 32px;
                    stroke: #742774;
                }

                .excel-title {
                    font-size: 24px;
                    font-weight: bold;
                    color: #333;
                    margin: 0 0 8px 0;
                }

                .excel-subtitle {
                    color: #666;
                    font-size: 14px;
                    margin: 0;
                }

                .excel-drop-zone {
                    display: flex;
                    flex-direction: column;
                    align-items: center;
                    justify-content: center;
                    width: 90%;
                    height: 128px;
                    border: 2px dashed #D8BFD8;
                    border-radius: 8px;
                    cursor: pointer;
                    background-color: #F9F5F9;
                    transition: background-color 0.2s;
                    margin-bottom: 16px;
                    padding: 20px;
                }

                .excel-drop-zone:hover {
                    background-color: #F3E5F5;
                }

                .excel-drop-zone svg {
                    width: 40px;
                    height: 40px;
                    stroke: #742774;
                    margin-bottom: 8px;
                }

                .excel-drop-zone-text {
                    font-size: 14px;
                    color: #666;
                    margin-bottom: 4px;
                }

                .excel-drop-zone-subtext {
                    font-size: 12px;
                    color: #999;
                }

                .excel-upload-btn {
                    width: 100%;
                    background-color: #ccc;
                    color: white;
                    font-weight: 600;
                    padding: 12px 16px;
                    border-radius: 8px;
                    border: none;
                    cursor: not-allowed;
                    font-size: 14px;
                    transition: background-color 0.2s;
                }

                .excel-upload-btn.active {
                    background-color: #742774;
                    cursor: pointer;
                }

                .excel-upload-btn.active:hover {
                    background-color: #5A1E5C;
                }

                .excel-file-input {
                    display: none;
                }

                .excel-selected-file {
                    background-color: #F9F5F9;
                    border: 1px solid #E1BEE7;
                    border-radius: 8px;
                    padding: 12px;
                    margin-bottom: 16px;
                    display: none;
                    font-size: 14px;
                    color: #333;
                }

                .excel-selected-file.show {
                    display: block;
                }

                .excel-error {
                    background-color: #FFF3F3;
                    border: 1px solid #FFB3B3;
                    border-radius: 8px;
                    padding: 12px;
                    margin-bottom: 16px;
                    display: none;
                    font-size: 14px;
                    color: #D32F2F;
                }

                .excel-error.show {
                    display: block;
                }
            </style>

            <div class="excel-upload-container">
                <div class="excel-card">
                    <div class="excel-header">
                        <div class="excel-icon-circle">
                            <svg viewBox="0 0 24 24" fill="none" stroke-width="2">
                                <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                                <polyline points="17 8 12 3 7 8"/>
                                <line x1="12" y1="3" x2="12" y2="15"/>
                            </svg>
                        </div>
                        <h1 class="excel-title">Upload DOI BiWeekly Employee Roster File</h1>
                        <p class="excel-subtitle">Select an Excel file to upload</p>
                    </div>

                    <label for="excel-file-upload" class="excel-drop-zone">
                        <svg viewBox="0 0 24 24" fill="none" stroke-width="2">
                            <path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/>
                            <polyline points="17 8 12 3 7 8"/>
                            <line x1="12" y1="3" x2="12" y2="15"/>
                        </svg>
                        <p class="excel-drop-zone-text">
                            <span style="font-weight: 600;">Click to upload</span> or drag and drop
                        </p>
                        <p class="excel-drop-zone-subtext">Excel files (.xlsx, .xls)</p>
                        <input id="excel-file-upload" type="file" class="excel-file-input" accept=".xlsx,.xls">
                    </label>

                    <div class="excel-error" id="excel-error">
                        <span style="font-weight: 600;">Error:</span> <span id="excel-error-text"></span>
                    </div>

                    <div class="excel-selected-file" id="excel-selected-file">
                        <span style="font-weight: 600;">Selected file:</span> <span id="excel-file-name"></span>
                    </div>

                    <button class="excel-upload-btn" id="excel-upload-btn" disabled>Upload File</button>
                </div>
            </div>
        `;

        // Get references to elements
        this.fileInput = this.container.querySelector('#excel-file-upload') as HTMLInputElement;
        this.selectedFileDiv = this.container.querySelector('#excel-selected-file') as HTMLDivElement;
        this.fileNameSpan = this.container.querySelector('#excel-file-name') as HTMLSpanElement;
        this.uploadBtn = this.container.querySelector('#excel-upload-btn') as HTMLButtonElement;
        const errorDiv = this.container.querySelector('#excel-error') as HTMLDivElement;
        const errorText = this.container.querySelector('#excel-error-text') as HTMLSpanElement;

        // File input change handler
        this.fileInput.addEventListener('change', (e) => {
            const file = (e.target as HTMLInputElement).files?.[0];
            if (file) {
                // Validate file
                const validExtensions = ['.xlsx', '.xls'];
                const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase();
                
                if (!validExtensions.includes(fileExtension)) {
                    errorText.textContent = 'Invalid file type. Please upload .xlsx or .xls files only.';
                    errorDiv.classList.add('show');
                    this.selectedFileDiv.classList.remove('show');
                    this.uploadBtn.classList.remove('active');
                    this.uploadBtn.disabled = true;
                    this.selectedFile = null;
                    return;
                }

                // Hide error, show selected file
                errorDiv.classList.remove('show');
                this.fileNameSpan.textContent = file.name;
                this.selectedFileDiv.classList.add('show');
                this.uploadBtn.classList.add('active');
                this.uploadBtn.disabled = false;
                this.selectedFile = file;
            }
        });

        // Upload button click handler
        this.uploadBtn.addEventListener('click', () => {
            if (this.selectedFile) {
                // Process the file here
                console.log('Processing file:', this.selectedFile.name);
                
                // You can read the file and process it
                const reader = new FileReader();
                reader.onload = (e) => {
                    const data = e.target?.result;
                    console.log('File data loaded', data);
                    // Process Excel data here
                    // You might want to use a library like xlsx or sheetjs
                };
                reader.readAsArrayBuffer(this.selectedFile);

                // Notify Power Apps that output has changed
                this.notifyOutputChanged();
            }
        });
    }

    public updateView(context: ComponentFramework.Context<IInputs>): void {
        // Update view if needed based on context changes
    }

public getOutputs(): IOutputs {
    return {
        fileName: this.selectedFile?.name || "",
        fileData: this.encodedFileData || "",
        uploadStatus: this.uploadStatus || "none"
    };
}

    public destroy(): void {
        // Cleanup
    }
}