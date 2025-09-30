import * as React from "react";
import * as XLSX from "xlsx";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHelpers } from "../../services/SPHelper";
import "./AdvancedBulkUpload.css";
interface IUploadAndCreateProps {
    context: WebPartContext;
    listName: string;
    listFields: string[]; // Internal names of the SharePoint list fields
}

interface IUploadStatus {
    message: string;
    success: boolean;
}

interface IColumnMapping {
    fileColumn: string;
    listField: string;
}

const AdvancedBulkUpload: React.FC<IUploadAndCreateProps> = ({ context, listName, listFields }) => {
    const [uploading, setUploading] = React.useState(false);
    const [progress, setProgress] = React.useState(0);
    const [totalItems, setTotalItems] = React.useState(0);
    const [statusLog, setStatusLog] = React.useState<IUploadStatus[]>([]);
    const [dataRows, setDataRows] = React.useState<any[]>([]);
    const [fileColumns, setFileColumns] = React.useState<string[]>([]);
    const [columnMapping, setColumnMapping] = React.useState<IColumnMapping[]>([]);
    const [sheetNames, setSheetNames] = React.useState<string[]>([]);
    const [selectedSheet, setSelectedSheet] = React.useState<string>("");
    const [dragOver, setDragOver] = React.useState(false);
    const spHelper = new SPHelpers(context.spHttpClient);
    const addLog = (message: string, success: boolean): void => {
        setStatusLog(prev => [...prev, { message, success }]);
    };

    const resetState = (): void => {
        setUploading(false);
        setProgress(0);
        setTotalItems(0);
        setStatusLog([]);
        setDataRows([]);
        setFileColumns([]);
        setColumnMapping([]);
        setSheetNames([]);
        setSelectedSheet("");
    };

    // Handle file drop or selection
    const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement> | React.DragEvent<HTMLDivElement>): Promise<void> => {
        e.preventDefault();
        resetState();

        let file: File | null = null;
        if ("dataTransfer" in e && e.dataTransfer.files.length > 0) {
            file = e.dataTransfer.files[0];
        } else if ("target" in e && (e as React.ChangeEvent<HTMLInputElement>).target.files?.length > 0) {
            file = (e as React.ChangeEvent<HTMLInputElement>).target.files[0];
        }
        if (!file) return;

        setUploading(true);

        try {
            if (file.name.indexOf(".json") > -1) {
                const text = await file.text();
                const rows = JSON.parse(text);
                setDataRows(rows);
                setFileColumns(Object.keys(rows[0] || {}));
                setColumnMapping(Object.keys(rows[0] || {}).map(col => ({ fileColumn: col, listField: col })));
                setTotalItems(rows.length);
            } else if (file.name.indexOf(".xlsx") > -1 || file.name.indexOf(".xls") > -1) {
                const reader = new FileReader();
                reader.onload = (ev: any) => {
                    const workbook = XLSX.read(ev.target.result, { type: "binary" });
                    setSheetNames(workbook.SheetNames);
                    setSelectedSheet(workbook.SheetNames[0]);
                    loadSheet(workbook, workbook.SheetNames[0]);
                };
                reader.readAsBinaryString(file);
            } else {
                addLog("❌ Unsupported file type", false);
            }
        } catch (err) {
            addLog("❌ Error reading file: " + err, false);
        } finally {
            setUploading(false);
        }
    };

    // Load selected sheet
    const loadSheet = (workbook: XLSX.WorkBook, sheetName: string): void => {
        const sheet = workbook.Sheets[sheetName];
        const rows = XLSX.utils.sheet_to_json(sheet);
        setDataRows(rows);
        setFileColumns(Object.keys(rows[0] as {}));
        setColumnMapping(Object.keys(rows[0] as {}).map(col => ({ fileColumn: col, listField: col })));
        setTotalItems(rows.length);
    };

    const handleSheetChange = (sheetName: string): void => {
        setSelectedSheet(sheetName);
        const reader = new FileReader();
        console.log(reader);
        // Reload the sheet
        // For simplicity, require user to re-upload file if multiple sheets
        addLog("⚠ Please re-upload file to change sheet.", false);
    };

    const handleMappingChange = (index: number, listField: string): void => {
        const updated = [...columnMapping];
        updated[index].listField = listField;
        setColumnMapping(updated);
    };

    // Bulk upload with batch processing
    const handleBulkUpload = async (batchSize: number = 50): Promise<void> => {
        if (dataRows.length === 0) {
            addLog("❌ No data to upload", false);
            return;
        }
        console.log(fileColumns);
        setUploading(true);
        const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
        let completed = 0;

        for (let i = 0; i < dataRows.length; i += batchSize) {
            const batch = dataRows.slice(i, i + batchSize);

            await Promise.all(batch.map(async row => {
                const item: any = {};
                columnMapping.forEach(map => {
                    item[map.listField] = row[map.fileColumn];
                });

                let success = false;
                let attempts = 0;
                while (!success && attempts < 3) {
                    attempts++;
                    try {
                        const response = await spHelper.setListData(url, JSON.stringify(item));
                        if (response.ok) {
                            addLog(`✅ Created: ${item.Title || JSON.stringify(item)}`, true);
                            success = true;
                        } else {
                            addLog(`⚠ Attempt ${attempts} failed for ${item.Title || "Unnamed"}: ${response.statusText}`, false);
                        }
                    } catch (err) {
                        addLog(`⚠ Attempt ${attempts} error for ${item.Title || "Unnamed"}: ${err}`, false);
                    }
                }

                if (!success) addLog(`❌ Failed after 3 attempts: ${item.Title || "Unnamed"}`, false);
                completed++;
                setProgress(Math.round((completed / dataRows.length) * 100));
            }));

            await new Promise(res => setTimeout(res, 50)); // smooth progress
        }

        addLog("✅ Bulk upload complete!", true);
        setUploading(false);
    };

    return (
        <div className="upload-card">
            <h3 className="title">📂 Advanced Bulk Upload → SharePoint</h3>

            {/* Drag & Drop */}
            <div
                className={`drop-zone ${dragOver ? "drag-over" : ""}`}
                onDrop={handleFileUpload}
                onDragOver={e => { e.preventDefault(); setDragOver(true); }}
                onDragLeave={() => setDragOver(false)}
            >
                Drag & Drop JSON/Excel here or click below
                <input type="file" accept=".json,.xlsx,.xls" onChange={handleFileUpload} disabled={uploading} />
            </div>

            {/* Sheet selection */}
            {sheetNames.length > 1 && (
                <div className="sheet-select">
                    <label>Select Sheet: </label>
                    <select disabled={uploading} value={selectedSheet} onChange={e => handleSheetChange(e.target.value)}>
                        {sheetNames.map(s => <option key={s} value={s}>{s}</option>)}
                    </select>
                </div>
            )}

            {/* Column mapping */}
            {columnMapping.length > 0 && (
                <div className="mapping-section">
                    <h4>Column Mapping</h4>
                    {columnMapping.map((map, idx) => (
                        <div key={idx} className="mapping-row">
                            <label>{map.fileColumn} → </label>
                            <select disabled={uploading} value={map.listField} onChange={e => handleMappingChange(idx, e.target.value)}>
                                <option value="">-- Select Field --</option>
                                {listFields.map(f => <option key={f} value={f}>{f}</option>)}
                            </select>
                        </div>
                    ))}
                </div>
            )}

            {/* Upload button */}
            {dataRows.length > 0 && (
                <button className="upload-btn" onClick={() => handleBulkUpload()} disabled={uploading}>
                    {uploading ? "⏳ Uploading..." : "🚀 Start Bulk Upload"}
                </button>
            )}

            {/* Progress bar */}
            {totalItems > 0 && (
                <div className="progress-container">
                    <div className="progress-bar" style={{ width: `${progress}%` }}>
                        {progress}%
                    </div>
                </div>
            )}

            {/* Logs */}
            <div className="log-box">
                {statusLog.map((log, idx) => (
                    <div key={idx} className={`log-message ${log.success ? "success" : "error"}`}>
                        {log.message}
                    </div>
                ))}
            </div>
        </div>
    );
};

export default AdvancedBulkUpload;
