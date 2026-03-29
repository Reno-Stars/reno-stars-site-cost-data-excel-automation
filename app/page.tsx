"use client";

import { useState, useCallback, useMemo, memo } from "react";
import type { ProcessResult } from "@/lib/excel-processor";

interface ProcessApiResponse extends ProcessResult {
  file: string;
  fileName: string;
}

const FileDropZone = memo(function FileDropZone({
  label,
  file,
  onFile,
  description,
  onInvalidFile,
}: {
  label: string;
  file: File | null;
  onFile: (f: File) => void;
  description: string;
  onInvalidFile: () => void;
}) {
  const [dragOver, setDragOver] = useState(false);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragOver(false);
      const f = e.dataTransfer.files[0];
      if (f && f.name.endsWith(".xlsx")) {
        onFile(f);
      } else if (f) {
        onInvalidFile();
      }
    },
    [onFile, onInvalidFile]
  );

  const handleChange = useCallback(
    (e: React.ChangeEvent<HTMLInputElement>) => {
      const f = e.target.files?.[0];
      if (f) onFile(f);
    },
    [onFile]
  );

  return (
    <div
      onDragOver={(e) => {
        e.preventDefault();
        setDragOver(true);
      }}
      onDragLeave={() => setDragOver(false)}
      onDrop={handleDrop}
      role="region"
      aria-label={`Drop zone: ${label}`}
      className={`relative border-2 border-dashed rounded-xl p-8 text-center transition-colors cursor-pointer ${
        dragOver
          ? "border-blue-500 bg-blue-50"
          : file
          ? "border-green-400 bg-green-50"
          : "border-gray-300 hover:border-gray-400 bg-white"
      }`}
    >
      <input
        type="file"
        accept=".xlsx"
        onChange={handleChange}
        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
        aria-label={`Upload ${label}`}
      />
      <div className="space-y-2">
        <div className="text-4xl" aria-hidden="true">{file ? "✅" : "📄"}</div>
        <p className="font-semibold text-lg">{label}</p>
        <p className="text-sm text-gray-500">{description}</p>
        {file && (
          <p className="text-sm text-green-700 font-medium mt-2">
            {file.name} ({(file.size / 1024).toFixed(1)} KB)
          </p>
        )}
      </div>
    </div>
  );
});

export default function Home() {
  const [inputFile, setInputFile] = useState<File | null>(null);
  const [outputFile, setOutputFile] = useState<File | null>(null);
  const [dateLabel, setDateLabel] = useState("");
  const [processing, setProcessing] = useState(false);
  const [result, setResult] = useState<ProcessApiResponse | null>(null);
  const [error, setError] = useState<string | null>(null);

  const handleInvalidFile = useCallback(() => {
    setError("Only .xlsx files are accepted");
  }, []);

  const handleProcess = async () => {
    if (!inputFile || !outputFile || !dateLabel) return;

    setProcessing(true);
    setError(null);
    setResult(null);

    try {
      const formData = new FormData();
      formData.append("input", inputFile);
      formData.append("output", outputFile);
      formData.append("dateLabel", dateLabel);

      const response = await fetch("/api/process", {
        method: "POST",
        body: formData,
      });

      const contentType = response.headers.get("content-type") || "";
      if (!contentType.includes("application/json")) {
        throw new Error(`Server error (${response.status})`);
      }

      const data = await response.json();

      if (!response.ok) {
        throw new Error(data.error || "Processing failed");
      }

      setResult(data);
    } catch (err) {
      setError(err instanceof Error ? err.message : "An error occurred");
    } finally {
      setProcessing(false);
    }
  };

  const handleDownload = () => {
    if (!result) return;

    try {
      const byteString = atob(result.file);
      const byteArray = new Uint8Array(byteString.length);
      for (let i = 0; i < byteString.length; i++) {
        byteArray[i] = byteString.charCodeAt(i);
      }
      const blob = new Blob([byteArray], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = result.fileName;
      a.click();
      URL.revokeObjectURL(url);
    } catch {
      setError("Failed to download file");
    }
  };

  const handleReset = () => {
    setInputFile(null);
    setOutputFile(null);
    setDateLabel("");
    setResult(null);
    setError(null);
  };

  const workerGrandTotals = useMemo(() => {
    if (!result) return null;
    return {
      labor: result.workerTotals.reduce((s, w) => s + w.labor, 0),
      materials: result.workerTotals.reduce((s, w) => s + w.materials, 0),
      other: result.workerTotals.reduce((s, w) => s + w.other, 0),
      total: result.workerTotals.reduce((s, w) => s + w.total, 0),
    };
  }, [result]);

  const sortedSiteTotals = useMemo(() => {
    if (!result) return [];
    return [...result.siteTotals].sort((a, b) => b.total - a.total);
  }, [result]);

  const siteGrandTotals = useMemo(() => {
    if (!result) return null;
    return {
      labor: result.siteTotals.reduce((s, st) => s + st.labor, 0),
      materials: result.siteTotals.reduce((s, st) => s + st.materials, 0),
      other: result.siteTotals.reduce((s, st) => s + st.other, 0),
      total: result.siteTotals.reduce((s, st) => s + st.total, 0),
    };
  }, [result]);

  return (
    <main className="min-h-screen py-12 px-4">
      <div className="max-w-4xl mx-auto space-y-8">
        {/* Header */}
        <div className="text-center space-y-2">
          <h1 className="text-3xl font-bold">Site Cost Data Automation</h1>
          <p className="text-gray-600">
            Upload the input cost sheet and output project file, then download
            the updated output with data appended.
          </p>
        </div>

        {/* Upload Section */}
        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          <FileDropZone
            label="Input File (Cost Sheet)"
            description="Worker hours, rates, materials per address"
            file={inputFile}
            onFile={setInputFile}
            onInvalidFile={handleInvalidFile}
          />
          <FileDropZone
            label="Output File (Project Sheets)"
            description="Date-range sheets with project cost tracking"
            file={outputFile}
            onFile={setOutputFile}
            onInvalidFile={handleInvalidFile}
          />
        </div>

        {/* Date Label */}
        <div className="flex justify-center">
          <div className="bg-white border rounded-lg p-4 flex items-center gap-4">
            <label
              htmlFor="dateLabel"
              className="font-semibold whitespace-nowrap"
            >
              日期 (Date Label):
            </label>
            <input
              id="dateLabel"
              type="text"
              value={dateLabel}
              onChange={(e) => setDateLabel(e.target.value)}
              placeholder="e.g. 3月上"
              className="border rounded-lg px-4 py-2 w-40 focus:outline-none focus:ring-2 focus:ring-blue-500"
            />
          </div>
        </div>

        {/* Process Button */}
        <div className="flex justify-center gap-4">
          <button
            onClick={handleProcess}
            disabled={!inputFile || !outputFile || !dateLabel || processing}
            className="px-8 py-3 bg-blue-600 text-white font-semibold rounded-lg hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors"
          >
            {processing ? "Processing..." : "Process Files"}
          </button>
          {(inputFile || outputFile || result) && (
            <button
              onClick={handleReset}
              className="px-8 py-3 bg-gray-200 text-gray-700 font-semibold rounded-lg hover:bg-gray-300 transition-colors"
            >
              Reset
            </button>
          )}
        </div>

        {/* Error */}
        {error && (
          <div role="alert" className="bg-red-50 border border-red-200 text-red-700 rounded-lg p-4">
            <p className="font-semibold">Error</p>
            <p>{error}</p>
          </div>
        )}

        {/* Results */}
        {result && (
          <div className="space-y-6" role="status" aria-live="polite">
            {/* Summary */}
            <div className="bg-white border rounded-lg p-6 space-y-4">
              <h2 className="text-xl font-bold">Processing Summary</h2>

              <div className="grid grid-cols-2 md:grid-cols-4 gap-4">
                <div className="bg-blue-50 rounded-lg p-4 text-center">
                  <p className="text-2xl font-bold text-blue-700">
                    {result.workers.length}
                  </p>
                  <p className="text-sm text-gray-600">Workers Found</p>
                </div>
                <div className="bg-green-50 rounded-lg p-4 text-center">
                  <p className="text-2xl font-bold text-green-700">
                    {result.rowsAdded}
                  </p>
                  <p className="text-sm text-gray-600">Rows Added</p>
                </div>
                <div className="bg-purple-50 rounded-lg p-4 text-center">
                  <p className="text-2xl font-bold text-purple-700">
                    {result.matchedSheets.length}
                  </p>
                  <p className="text-sm text-gray-600">Sheets Updated</p>
                </div>
                <div
                  className={`rounded-lg p-4 text-center ${
                    result.unmatchedAddresses.length > 0
                      ? "bg-yellow-50"
                      : "bg-gray-50"
                  }`}
                >
                  <p
                    className={`text-2xl font-bold ${
                      result.unmatchedAddresses.length > 0
                        ? "text-yellow-700"
                        : "text-gray-400"
                    }`}
                  >
                    {result.unmatchedAddresses.length}
                  </p>
                  <p className="text-sm text-gray-600">Unmatched Addresses</p>
                </div>
              </div>

              {result.unmatchedAddresses.length > 0 && (
                <div role="alert" className="bg-yellow-50 border border-yellow-200 rounded-lg p-4">
                  <p className="font-semibold text-yellow-800">
                    Unmatched address codes (no matching project block found):
                  </p>
                  <p className="text-yellow-700">
                    {result.unmatchedAddresses.join(", ")}
                  </p>
                </div>
              )}

              {result.droppedMaterials > 0 && (
                <div role="alert" className="bg-orange-50 border border-orange-200 rounded-lg p-4">
                  <p className="font-semibold text-orange-800">
                    {result.droppedMaterials} material cost{result.droppedMaterials > 1 ? "s" : ""} could not be written (no empty K cell in block)
                  </p>
                </div>
              )}

              {result.matchedSheets.length > 0 && (
                <div className="bg-green-50 border border-green-200 rounded-lg p-4">
                  <p className="font-semibold text-green-800">
                    Updated sheets:
                  </p>
                  <p className="text-green-700">
                    {result.matchedSheets.join(", ")}
                  </p>
                </div>
              )}
            </div>

            {/* Worker Details */}
            <div className="bg-white border rounded-lg p-6 space-y-4">
              <h2 className="text-xl font-bold">Extracted Worker Data</h2>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead>
                    <tr className="border-b bg-gray-50">
                      <th scope="col" className="text-left p-2">Worker</th>
                      <th scope="col" className="text-left p-2">Rate ($/hr)</th>
                      <th scope="col" className="text-left p-2">Address</th>
                      <th scope="col" className="text-right p-2">Hours</th>
                      <th scope="col" className="text-right p-2">Materials</th>
                      <th scope="col" className="text-right p-2">Other</th>
                    </tr>
                  </thead>
                  <tbody>
                    {result.workers.map((worker, wi) =>
                      worker.entries.map((entry, ei) => (
                        <tr
                          key={`${wi}-${ei}`}
                          className={`border-b ${
                            ei === 0 ? "bg-white" : "bg-gray-50/50"
                          }`}
                        >
                          <td className="p-2 font-medium">
                            {ei === 0 ? worker.name : ""}
                          </td>
                          <td className="p-2">
                            {ei === 0 ? `$${worker.rate}` : ""}
                          </td>
                          <td className="p-2 font-mono">{entry.address}</td>
                          <td className="p-2 text-right">
                            {entry.hours || "-"}
                          </td>
                          <td className="p-2 text-right">
                            {entry.materials
                              ? `$${entry.materials.toFixed(2)}`
                              : "-"}
                          </td>
                          <td className="p-2 text-right">
                            {entry.gas + entry.ticket > 0
                              ? `$${(entry.gas + entry.ticket).toFixed(2)}`
                              : "-"}
                          </td>
                        </tr>
                      ))
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Worker Totals */}
            {result.workerTotals.length > 0 && workerGrandTotals && (
              <div className="bg-white border rounded-lg p-6 space-y-4">
                <h2 className="text-xl font-bold">Employee Totals</h2>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="border-b bg-gray-50">
                        <th scope="col" className="text-left p-2">Employee</th>
                        <th scope="col" className="text-left p-2">Rate</th>
                        <th scope="col" className="text-right p-2">Labor</th>
                        <th scope="col" className="text-right p-2">Materials</th>
                        <th scope="col" className="text-right p-2">Other</th>
                        <th scope="col" className="text-right p-2 font-bold">Total</th>
                      </tr>
                    </thead>
                    <tbody>
                      {result.workerTotals.map((wt) => (
                        <tr key={wt.name} className="border-b">
                          <td className="p-2 font-medium">{wt.name}</td>
                          <td className="p-2">${wt.rate}/hr</td>
                          <td className="p-2 text-right">${wt.labor.toFixed(2)}</td>
                          <td className="p-2 text-right">${wt.materials.toFixed(2)}</td>
                          <td className="p-2 text-right">${wt.other.toFixed(2)}</td>
                          <td className="p-2 text-right font-bold">${wt.total.toFixed(2)}</td>
                        </tr>
                      ))}
                      <tr className="border-t-2 border-gray-300 bg-gray-50 font-bold">
                        <td className="p-2" colSpan={2}>Grand Total</td>
                        <td className="p-2 text-right">${workerGrandTotals.labor.toFixed(2)}</td>
                        <td className="p-2 text-right">${workerGrandTotals.materials.toFixed(2)}</td>
                        <td className="p-2 text-right">${workerGrandTotals.other.toFixed(2)}</td>
                        <td className="p-2 text-right">${workerGrandTotals.total.toFixed(2)}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Site Totals */}
            {sortedSiteTotals.length > 0 && siteGrandTotals && (
              <div className="bg-white border rounded-lg p-6 space-y-4">
                <h2 className="text-xl font-bold">Site Totals</h2>
                <div className="overflow-x-auto">
                  <table className="w-full text-sm">
                    <thead>
                      <tr className="border-b bg-gray-50">
                        <th scope="col" className="text-left p-2">Site (Address Code)</th>
                        <th scope="col" className="text-right p-2">Labor</th>
                        <th scope="col" className="text-right p-2">Materials</th>
                        <th scope="col" className="text-right p-2">Other</th>
                        <th scope="col" className="text-right p-2 font-bold">Total</th>
                      </tr>
                    </thead>
                    <tbody>
                      {sortedSiteTotals.map((st) => (
                          <tr key={st.address} className="border-b">
                            <td className="p-2 font-mono">{st.address}</td>
                            <td className="p-2 text-right">${st.labor.toFixed(2)}</td>
                            <td className="p-2 text-right">${st.materials.toFixed(2)}</td>
                            <td className="p-2 text-right">${st.other.toFixed(2)}</td>
                            <td className="p-2 text-right font-bold">${st.total.toFixed(2)}</td>
                          </tr>
                        ))}
                      <tr className="border-t-2 border-gray-300 bg-gray-50 font-bold">
                        <td className="p-2">Grand Total</td>
                        <td className="p-2 text-right">${siteGrandTotals.labor.toFixed(2)}</td>
                        <td className="p-2 text-right">${siteGrandTotals.materials.toFixed(2)}</td>
                        <td className="p-2 text-right">${siteGrandTotals.other.toFixed(2)}</td>
                        <td className="p-2 text-right">${siteGrandTotals.total.toFixed(2)}</td>
                      </tr>
                    </tbody>
                  </table>
                </div>
              </div>
            )}

            {/* Download */}
            <div className="flex justify-center">
              <button
                onClick={handleDownload}
                className="px-10 py-4 bg-green-600 text-white font-bold text-lg rounded-lg hover:bg-green-700 transition-colors shadow-lg"
              >
                Download Updated Output File
              </button>
            </div>
          </div>
        )}
      </div>
    </main>
  );
}
