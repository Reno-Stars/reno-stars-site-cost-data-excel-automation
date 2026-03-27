"use client";

import { useState, useCallback } from "react";
import type { ProcessResult } from "@/lib/excel-processor";

interface ProcessApiResponse extends ProcessResult {
  file: string;
  fileName: string;
}

function FileDropZone({
  label,
  file,
  onFile,
  description,
}: {
  label: string;
  file: File | null;
  onFile: (f: File) => void;
  description: string;
}) {
  const [dragOver, setDragOver] = useState(false);

  const handleDrop = useCallback(
    (e: React.DragEvent) => {
      e.preventDefault();
      setDragOver(false);
      const f = e.dataTransfer.files[0];
      if (f && f.name.endsWith(".xlsx")) {
        onFile(f);
      }
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
        onChange={(e) => {
          const f = e.target.files?.[0];
          if (f) onFile(f);
        }}
        className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
      />
      <div className="space-y-2">
        <div className="text-4xl">{file ? "✅" : "📄"}</div>
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
}

export default function Home() {
  const [inputFile, setInputFile] = useState<File | null>(null);
  const [outputFile, setOutputFile] = useState<File | null>(null);
  const [dateLabel, setDateLabel] = useState("");
  const [processing, setProcessing] = useState(false);
  const [result, setResult] = useState<ProcessApiResponse | null>(null);
  const [error, setError] = useState<string | null>(null);

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
  };

  const handleReset = () => {
    setInputFile(null);
    setOutputFile(null);
    setDateLabel("");
    setResult(null);
    setError(null);
  };

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
          />
          <FileDropZone
            label="Output File (Project Sheets)"
            description="Per-project sheets with labor & material tracking"
            file={outputFile}
            onFile={setOutputFile}
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
          <div className="bg-red-50 border border-red-200 text-red-700 rounded-lg p-4">
            <p className="font-semibold">Error</p>
            <p>{error}</p>
          </div>
        )}

        {/* Results */}
        {result && (
          <div className="space-y-6">
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
                <div className="bg-yellow-50 border border-yellow-200 rounded-lg p-4">
                  <p className="font-semibold text-yellow-800">
                    Unmatched address codes (no sheet found):
                  </p>
                  <p className="text-yellow-700">
                    {result.unmatchedAddresses.join(", ")}
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
                      <th className="text-left p-2">Worker</th>
                      <th className="text-left p-2">Rate ($/hr)</th>
                      <th className="text-left p-2">Address</th>
                      <th className="text-right p-2">Hours</th>
                      <th className="text-right p-2">Materials</th>
                      <th className="text-right p-2">Other</th>
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
