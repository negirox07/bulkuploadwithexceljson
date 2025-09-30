import * as React from "react";
import * as XLSX from "xlsx";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHelpers } from "../services/SPHelper";

interface IUploadAndCreateProps {
  context: WebPartContext;
  listName: string;
}
interface IUploadStatus {
  message: string;
  success: boolean;
}
const UploadAndCreate: React.FC<IUploadAndCreateProps> = ({ context, listName }) => {
  const [progress, setProgress] = React.useState<number>(0);
  const [totalItems, setTotalItems] = React.useState<number>(0);
  const [statusLog, setStatusLog] = React.useState<IUploadStatus[]>([]);
  const [uploading, setUploading] = React.useState<boolean>(false);
  const spHelper = new SPHelpers(context.spHttpClient);
  // Add log entry
  const addLog = (message: string, success: boolean): void => {
    setStatusLog(prev => [...prev, { message, success }]);
  };
  const resetState = (): void => {
    setProgress(0);
    setTotalItems(0);
    setStatusLog([]);
  };
    // Create SharePoint list items with retry and progress
  const createItems = async (records: any[]): Promise<void> => {
    const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
    let completed = 0;

    for (const record of records) {
      let success = false;
      let attempts = 0;

      while (!success && attempts < 3) { // retry up to 3 times
        attempts++;
        try {
          const response = await spHelper.setListData(
            url,
            JSON.stringify(record)
          );

          if (response.ok) {
            addLog(`✅ Created: ${record.Title || JSON.stringify(record)}`, true);
            success = true;
          } else {
            addLog(`⚠ Attempt ${attempts} failed for ${record.Title || "Unnamed"}: ${response.statusText}`, false);
          }
        } catch (err) {
          addLog(`⚠ Attempt ${attempts} error for ${record.Title || "Unnamed"}: ${err}`, false);
        }
      }

      if (!success) {
        addLog(`❌ Failed after 3 attempts: ${record.Title || "Unnamed"}`, false);
      }

      completed++;
      setProgress(Math.round((completed / records.length) * 100));
      await new Promise(res => setTimeout(res, 50)); // smooth progress
    }

    addLog("✅ Upload complete!", true);
  };

  // Read JSON
  const readJson = async (file: File): Promise<void> => {
    const reader = new FileReader();
    reader.onload = async (e: any) => {
      try {
        const data = JSON.parse(e.target.result);
        setTotalItems(data.length);
        await createItems(data);
      } catch (err) {
        addLog("❌ Invalid JSON format: " + err, false);
      } finally {
        setUploading(false);
      }
    };
    reader.readAsText(file);
  };

  // Read Excel
  const readExcel = async (file: File): Promise<void> => {
    const reader = new FileReader();
    reader.onload = async (e: any) => {
      try {
        const workbook = XLSX.read(e.target.result, { type: "binary" });
        const allData: any[] = [];

        workbook.SheetNames.forEach(sheetName => {
          const worksheet = workbook.Sheets[sheetName];
          const jsonData = XLSX.utils.sheet_to_json(worksheet);
          allData.push(...jsonData);
        });

        setTotalItems(allData.length);
        await createItems(allData);
      } catch (err) {
        addLog("❌ Error reading Excel: " + err, false);
      } finally {
        setUploading(false);
      }
    };
    reader.readAsBinaryString(file);
  };
  // Handle file selection
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>): Promise<void> => {
    const file = e.target.files?.[0];
    if (!file) return;

    resetState();
    setUploading(true);

    if (file.name.indexOf(".json") > -1) {
      await readJson(file);
    } else if (file.name.indexOf(".xlsx") > -1 || file.name.indexOf(".xls") > -1) {
      await readExcel(file);
    } else {
      addLog("❌ Unsupported file type. Please use JSON or Excel.", false);
      setUploading(false);
    }
  };

  return (
    <div style={{ padding: 15, maxWidth: 600 }}>
      <h3>Upload JSON/Excel → Create SharePoint List Items</h3>
      <input type="file" accept=".json,.xlsx,.xls" onChange={handleFileUpload} disabled={uploading} />

      {/* Progress Bar */}
      {totalItems > 0 && (
        <div style={{ marginTop: 15 }}>
          <div style={{ width: "100%", backgroundColor: "#eee", borderRadius: 4 }}>
            <div
              style={{
                width: `${progress}%`,
                backgroundColor: "#4caf50",
                height: 25,
                borderRadius: 4,
                textAlign: "center",
                color: "white",
                fontWeight: "bold"
              }}
            >
              {progress}%
            </div>
          </div>
        </div>
      )}

      {/* Status Log */}
      <div style={{ marginTop: 15, maxHeight: 300, overflowY: "auto", border: "1px solid #ddd", padding: 10, borderRadius: 4, backgroundColor: "#f9f9f9" }}>
        {statusLog.map((log, idx) => (
          <div key={idx} style={{ color: log.success ? "green" : "red", marginBottom: 2 }}>
            {log.message}
          </div>
        ))}
      </div>
    </div>
  );
};

export default UploadAndCreate;
