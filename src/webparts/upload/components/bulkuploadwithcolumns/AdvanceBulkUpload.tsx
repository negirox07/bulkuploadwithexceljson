import * as React from "react";
import * as XLSX from "xlsx";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { SPHttpClient } from "@microsoft/sp-http";
import { SPHelpers } from "../../services/SPHelper";

interface IUploadAndCreateProps {
  context: WebPartContext;
  listName: string;
}

interface IFieldMeta {
  Title: string;
  InternalName: string;
  TypeAsString: string;
  LookupList?: string;
  LookupField?: string;
  AllowMultipleValues?: boolean;
  Choices?: string[];
}

interface IColumnMapping {
  fileColumn: string;
  listField: string;
}

interface ILog {
  message: string;
  success: boolean;
}

interface IMismatchSummary {
  field: string;
  totalMismatches: number;
  values: string[];
}

export const AdvanceBulkUpload: React.FC<IUploadAndCreateProps> = ({ context, listName }) => {
  const [fileColumns, setFileColumns] = React.useState<string[]>([]);
  const [rows, setRows] = React.useState<any[]>([]);
  const [listFields, setListFields] = React.useState<IFieldMeta[]>([]);
  const [lookupLists, setLookupLists] = React.useState<Record<string, string>>({});
  const [lookupValues, setLookupValues] = React.useState<Record<string, { Id: number; Title: string }[]>>({});
  const [columnMapping, setColumnMapping] = React.useState<IColumnMapping[]>([]);
  const [logs, setLogs] = React.useState<ILog[]>([]);
  const [mismatchSummary, setMismatchSummary] = React.useState<IMismatchSummary[]>([]);
  const [isUploading, setIsUploading] = React.useState(false);
  const [progress, setProgress] = React.useState(0);
  const spHelper = new SPHelpers(context.spHttpClient);
  // --- Helpers ---
  const addLog = (message: string, success: boolean) =>
    setLogs(prev => [...prev, { message, success }]);

  const normalizeName = (str: string) => str.replace(/\s+/g, "").toLowerCase();

  // --- Fetch list fields ---
  const fetchListFields = async (): Promise<void> => {
    try {
      const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/fields?$select=Title,InternalName,TypeAsString,LookupList,LookupField,Choices,AllowMultipleValues&$filter=Hidden eq false and ReadOnlyField eq false`;
      const res = await context.spHttpClient.get(url, SPHttpClient.configurations.v1);
      const json = await res.json();
      setListFields(json.value);

      // Collect lookup GUIDs and resolve list titles + values
      const lookups = json.value.filter((f: IFieldMeta) => (f.TypeAsString === "Lookup" || f.TypeAsString === "LookupMulti") && f.LookupList);
      for (const lu of lookups) {
        if (lu.LookupList && !lookupLists[lu.LookupList]) {
          const cleanGuid = lu.LookupList.replace(/[{}]/g, "")
          try {
            const listUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${cleanGuid}')?$select=Title`;
            const lookupRes = await context.spHttpClient.get(listUrl, SPHttpClient.configurations.v1);
            if (lookupRes.ok) {
              const lookupJson = await lookupRes.json();
              setLookupLists(prev => ({ ...prev, [cleanGuid!]: lookupJson.Title }));

              const itemsUrl = `${context.pageContext.web.absoluteUrl}/_api/web/lists(guid'${cleanGuid}')/items?$select=Id,Title&$top=5000`;
              const itemsRes = await context.spHttpClient.get(itemsUrl, SPHttpClient.configurations.v1);
              if (itemsRes.ok) {
                const itemsJson = await itemsRes.json();
                setLookupValues(prev => ({ ...prev, [lu.InternalName]: itemsJson.value }));
              }
            }
          } catch {
            setLookupLists(prev => ({ ...prev, [cleanGuid!]: "Unknown List" }));
          }
        }
      }
    } catch (err) {
      addLog("❌ Error fetching list fields: " + err, false);
    }
  };

  // --- Handle file upload ---
  // --- Handle file upload ---
  const handleFileUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const fileName = file.name.toLowerCase();

    if (fileName.endsWith(".xlsx") || fileName.endsWith(".xls")) {
      // Excel file
      const data = await file.arrayBuffer();
      const workbook = XLSX.read(data);
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = XLSX.utils.sheet_to_json(firstSheet);

      setRows(json);
      setFileColumns(Object.keys(json[0] as {}));
    } else if (fileName.endsWith(".json")) {
      // JSON file
      const text = await file.text();
      try {
        const json = JSON.parse(text);
        if (Array.isArray(json) && json.length > 0) {
          setRows(json);
          setFileColumns(Object.keys(json[0]));
        } else {
          addLog("⚠ JSON file must contain an array of objects", false);
        }
      } catch (err) {
        addLog("❌ Invalid JSON file: " + err, false);
      }
    } else {
      addLog("⚠ Unsupported file type", false);
    }

    fetchListFields();
  };


  // --- Auto map columns ---
  React.useEffect(() => {
    if (!fileColumns.length || !listFields.length) return;

    const autoMapped = fileColumns.map(fc => {
      const normalizedFC = normalizeName(fc);
      const match = listFields.find(
        lf =>
          normalizeName(lf.Title) === normalizedFC ||
          normalizeName(lf.InternalName) === normalizedFC
      );
      return {
        fileColumn: fc,
        listField: match ? match.InternalName : "",
      };
    });

    setColumnMapping(autoMapped);
  }, [fileColumns, listFields]);

  // --- Lookup validation ---
  React.useEffect(() => {
    if (!rows.length || !columnMapping.length || !listFields.length) return;

    const summaries: IMismatchSummary[] = [];

    columnMapping.forEach(map => {
      if (!map.listField) return;
      const fieldMeta = listFields.find(f => f.InternalName === map.listField);
      if (!fieldMeta || fieldMeta.TypeAsString !== "Lookup" && fieldMeta.TypeAsString !== "LookupMulti") return;

      const lookupItems = lookupValues[map.listField] || [];
      const validTitles = lookupItems.map(lv => lv.Title.toLowerCase());

      const excelValues = rows.map(r => r[map.fileColumn]).filter(Boolean);
      const mismatches: string[] = [];

      excelValues.forEach(val => {
        if (fieldMeta.AllowMultipleValues && typeof val === "string") {
          val.split(",").map(v => v.trim()).forEach(v => {
            if (!validTitles.includes(v.toLowerCase())) mismatches.push(v);
          });
        } else {
          if (!validTitles.includes(String(val).toLowerCase())) mismatches.push(String(val));
        }
      });

      if (mismatches.length > 0) {
        summaries.push({
          field: fieldMeta.Title,
          totalMismatches: mismatches.length,
          values: Array.from(new Set(mismatches)).slice(0, 10),
        });
      }
    });

    setMismatchSummary(summaries);
    console.log("Lookup mismatch summary:", mismatchSummary);
  }, [rows, columnMapping, listFields, lookupValues]);

  const handleMappingChange = (idx: number, value: string) => {
    setColumnMapping(prev => {
      const copy = [...prev];
      copy[idx].listField = value;
      return copy;
    });
  };

  /*   // --- Upload rows to SharePoint ---
    const handleUpload = async () => {
      if (!rows.length) {
        addLog("⚠ No data to upload", false);
        return;
      }
      if (columnMapping.some(m => !m.listField)) {
        addLog("⚠ Some columns are not mapped", false);
        return;
      }
  
      setIsUploading(true);
      setLogs([]);
  
      for (const [i, row] of rows.entries()) {
        const item: any = {};
  
        for (const map of columnMapping) {
          if (!map.listField) continue;
          const fieldMeta = listFields.find(f => f.InternalName === map.listField);
          if (!fieldMeta) continue;
  
          let excelValue = row[map.fileColumn];
  
          // Normalize blanks/nulls
          if (excelValue === undefined || excelValue === null) {
            excelValue = "";
          }
  
          // --- Handle Lookup fields ---
          if (fieldMeta.TypeAsString === "LookupMulti") {
            const lookupItems = lookupValues[fieldMeta.InternalName] || [];
  
            if (fieldMeta.AllowMultipleValues) {
              // LookupMulti
              let ids: number[] = [];
  
              if (typeof excelValue === "string" && excelValue.trim() !== "") {
                ids = excelValue
                  .split(",")
                  .map(v => v.trim())
                  .map(v => {
                    const match = lookupItems.find(li => li.Title.toLowerCase() === v.toLowerCase());
                    return match ? match.Id : null;
                  })
                  .filter((id): id is number => id !== null);
              }
  
              if (ids.length > 0) {
                item[`${fieldMeta.InternalName}Id`] = ids;
              }
            } else {
              // Single lookup
              if (String(excelValue).trim() !== "") {
                const match = lookupItems.find(
                  li => li.Title.toLowerCase() === String(excelValue).toLowerCase()
                );
                if (match) {
                  item[`${fieldMeta.InternalName}Id`] = [match.Id];
                }
              }
            }
          }
  
          // --- Handle Choice / Text ---
          else if (fieldMeta.TypeAsString === "Choice" || fieldMeta.TypeAsString === "Text") {
            item[fieldMeta.InternalName] = excelValue || "";
          }
  
          // --- Handle MultiChoice ---
          else if (fieldMeta.TypeAsString === "MultiChoice") {
            if (typeof excelValue === "string" && excelValue.trim() !== "") {
              const values = excelValue
                .split(",")
                .map(v => v.trim())
                .filter(v => v.length > 0);
              item[fieldMeta.InternalName] = { results: values };
            }
          }
  
          // --- Default case (Numbers, Dates, etc.) ---
          else {
            item[fieldMeta.InternalName] = excelValue;
          }
        }
  
        try {
          console.log(item);
          const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
          const res = await spHelper.setListData(url, JSON.stringify(item));
          if (res.ok) {
            addLog(`✅ Row ${i + 1} uploaded successfully`, true);
          } else {
            const err = await res.text();
            addLog(`❌ Row ${i + 1} failed: ${err}`, false);
          }
        } catch (err) {
          addLog(`❌ Row ${i + 1} failed: ${err}`, false);
        }
      }
  
      setIsUploading(false);
    }; */

  const handleUpload = async () => {
    if (!rows.length) {
      addLog("⚠ No data to upload", false);
      return;
    }
    if (columnMapping.some(m => !m.listField)) {
      addLog("⚠ Some columns are not mapped", false);
      return;
    }

    setIsUploading(true);
    setLogs([]);
    setProgress(0);

    const total = rows.length;
    const batchSize = 10; // upload 10 rows in parallel
    let processed = 0;

    const createItemPayload = (row: any) => {
      const item: any = {};
      for (const map of columnMapping) {
        if (!map.listField) continue;
        const fieldMeta = listFields.find(f => f.InternalName === map.listField);
        if (!fieldMeta) continue;

        let value = row[map.fileColumn];
        if (value === undefined || value === null) value = "";

        // --- Lookup / LookupMulti ---
        if (fieldMeta.TypeAsString === "Lookup" || fieldMeta.TypeAsString === "LookupMulti") {
          const lookupItems = lookupValues[fieldMeta.InternalName] || [];
          if (fieldMeta.AllowMultipleValues) {
            const ids =
              typeof value === "string"
                ? value
                  .split(",")
                  .map(v => v.trim())
                  .map(v => {
                    const match = lookupItems.find(li => li.Title.toLowerCase() === v.toLowerCase());
                    return match ? match.Id : null;
                  })
                  .filter((id): id is number => id !== null)
                : [];
            if (ids.length > 0) item[`${fieldMeta.InternalName}Id`] = ids;
          } else {
            const match = lookupItems.find(li => li.Title.toLowerCase() === String(value).toLowerCase());
            if (match) item[`${fieldMeta.InternalName}Id`] = [match.Id];
          }
        }

        // --- Choice / Text ---
        else if (fieldMeta.TypeAsString === "Choice" || fieldMeta.TypeAsString === "Text") {
          item[fieldMeta.InternalName] = value || "";
        }

        // --- MultiChoice ---
        else if (fieldMeta.TypeAsString === "MultiChoice") {
          if (typeof value === "string" && value.trim() !== "") {
            item[fieldMeta.InternalName] = { results: value.split(",").map(v => v.trim()) };
          }
        }

        // --- Default ---
        else {
          item[fieldMeta.InternalName] = value;
        }
      }
      return item;
    };

    for (let i = 0; i < total; i += batchSize) {
      const batch = rows.slice(i, i + batchSize);

      await Promise.all(
        batch.map(async (row, idx) => {
          const item = createItemPayload(row);
          const rowIndex = i + idx + 1;

          try {
            const url = `${context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${listName}')/items`;
            const res = await spHelper.setListData(url, JSON.stringify(item));
            if (res.ok) {
              addLog(`✅ Row ${rowIndex} uploaded successfully`, true);
            } else {
              const err = await res.text();
              addLog(`❌ Row ${rowIndex} failed: ${err}`, false);
            }
          } catch (err) {
            addLog(`❌ Row ${rowIndex} failed: ${err}`, false);
          } finally {
            processed++;
            setProgress(Math.round((processed / total) * 100));
          }
        })
      );
    }

    setIsUploading(false);
  };

  // --- UI ---
  return (
    <div className="p-4 border rounded bg-gray-50">
      <h3 className="font-bold mb-2">📂 Bulk Upload → {listName}</h3>

      <input type="file" accept=".xlsx,.xls,.json" onChange={handleFileUpload} className="mb-4" />

      {columnMapping.length > 0 && (
        <div className="mb-4">
          <h4 className="font-semibold">Column Mapping</h4>
          {columnMapping.map((map, idx) => {
            const fieldMeta = listFields.find(f => f.InternalName === map.listField);
            const isLookup = fieldMeta?.TypeAsString === "Lookup";
            const isMultiLookup = isLookup && fieldMeta?.AllowMultipleValues;

            return (
              <div key={idx} className="mb-4 border-b pb-2">
                <div className="flex items-center gap-2 mb-2">
                  <span className="w-40">{map.fileColumn}</span>
                  <select
                    value={map.listField}
                    onChange={(e) => handleMappingChange(idx, e.target.value)}
                    className={`border p-1 flex-1 ${!map.listField ? "border-red-500" : "border-gray-300"
                      }`}
                  >
                    <option value="">-- Not Mapped --</option>
                    {listFields.map(f => {
                      let label = `${f.Title} (${f.TypeAsString})`;
                      if (f.TypeAsString === "Lookup" && f.LookupList) {
                        const lookupTitle = lookupLists[f.LookupList] || f.LookupList;
                        label = `${f.Title} (Lookup → ${lookupTitle})`;
                      }
                      return (
                        <option key={f.InternalName} value={f.InternalName}>
                          {label}
                        </option>
                      );
                    })}
                  </select>
                </div>

                {/* Lookup Preview */}
                {isLookup && lookupValues[map.listField] && (
                  <div className="ml-40 text-sm text-gray-600">
                    <div>Available values (sample):</div>

                    {isMultiLookup ? (
                      <div className="border p-2 rounded bg-white mt-1 max-h-28 overflow-y-auto">
                        {lookupValues[map.listField].map(lv => (
                          <label key={lv.Id} className="block">
                            <input type="checkbox" value={lv.Id} className="mr-2" />
                            {lv.Title}
                          </label>
                        ))}
                      </div>
                    ) : (
                      <select className="border p-1 w-60 mt-1">
                        {lookupValues[map.listField].map(lv => (
                          <option key={lv.Id} value={lv.Id}>
                            {lv.Title}
                          </option>
                        ))}
                      </select>
                    )}
                  </div>
                )}
              </div>
            );
          })}
          {columnMapping.some(m => !m.listField) && (
            <div className="text-red-600 mt-2">
              ⚠ Some columns are not mapped. Please map them before uploading.
            </div>
          )}
        </div>
      )}

      {/* Lookup mismatch summary */}
      {/* {mismatchSummary.length > 0 && (
        <div className="bg-yellow-100 border border-yellow-400 text-yellow-800 p-3 rounded mb-4">
          <h4 className="font-semibold">⚠ Lookup Validation Warnings</h4>
          {mismatchSummary.map((ms, idx) => (
            <div key={idx} className="mt-1">
              <span className="font-medium">{ms.field}</span>: {ms.totalMismatches} mismatches.
              <div className="ml-4 text-sm">Examples: {ms.values.join(", ")}{ms.values.length >= 10 && " ..."}</div>
            </div>
          ))}
        </div>
      )} */}

      {/* Upload Button */}
      {rows.length > 0 && (
        <button
          onClick={handleUpload}
          disabled={isUploading}
          className="px-4 py-2 bg-blue-600 text-white rounded hover:bg-blue-700 disabled:opacity-50"
        >
          {isUploading ? "Uploading..." : "Upload to SharePoint"}
        </button>
      )}
      {isUploading && (
        <div className="w-full bg-gray-200 rounded h-4 mb-4">
          <div
            className="bg-blue-600 h-4 rounded text-xs text-white text-center"
            style={{ width: `${progress}%` }}
          >
            {progress}%
          </div>
        </div>
      )}

      {/* Logs */}
      <div className="mt-4 bg-white p-2 border max-h-48 overflow-y-auto">
        {logs.map((log, idx) => (
          <div key={idx} className={log.success ? "text-green-600" : "text-red-600"}>
            {log.message}
          </div>
        ))}
      </div>
    </div>
  );
};
