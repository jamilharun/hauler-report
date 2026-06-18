import { useState, useCallback } from "react";
import { parseWorkbook } from "./parser";
import { generateReport, getMonthLabel } from "./generator";
import type { DriverGroup } from "./types";

type Step = "upload" | "preview" | "done";

export default function App() {
  const [step, setStep] = useState<Step>("upload");
  const [groups, setGroups] = useState<DriverGroup[]>([]);
  const [monthLabel, setMonthLabel] = useState("");
  const [error, setError] = useState("");
  const [loading, setLoading] = useState(false);
  const [dragging, setDragging] = useState(false);

  const handleFile = useCallback(async (file: File) => {
    if (!file.name.endsWith(".xlsx") && !file.name.endsWith(".xls")) {
      setError("Please upload an Excel file (.xlsx or .xls)");
      return;
    }
    setError("");
    setLoading(true);
    try {
      const result = await parseWorkbook(file);
      const label = getMonthLabel(result);
      setGroups(result);
      setMonthLabel(label);
      setStep("preview");
    } catch (err) {
      setError(err instanceof Error ? err.message : "Failed to parse file.");
    } finally {
      setLoading(false);
    }
  }, []);

  const onInputChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (file) handleFile(file);
  };

  const onDrop = (e: React.DragEvent) => {
    e.preventDefault();
    setDragging(false);
    const file = e.dataTransfer.files?.[0];
    if (file) handleFile(file);
  };

  const onDownload = async () => {
    await generateReport(groups, monthLabel);
    setStep("done");
  };

  const reset = () => {
    setStep("upload");
    setGroups([]);
    setMonthLabel("");
    setError("");
  };

  return (
    <div className="min-h-screen bg-gray-950 text-gray-100 flex flex-col items-center justify-center p-6">
      <div className="w-full max-w-3xl">
        <h1 className="text-2xl font-bold mb-1">Hauler Report Generator</h1>
        <p className="text-gray-400 mb-8 text-sm">
          Upload the summary sheet to generate per-driver reports.
        </p>

        {step === "upload" && (
          <div
            onDragOver={(e) => {
              e.preventDefault();
              setDragging(true);
            }}
            onDragLeave={() => setDragging(false)}
            onDrop={onDrop}
            className={`border-2 border-dashed rounded-xl p-16 text-center transition-colors cursor-pointer
              ${dragging ? "border-blue-400 bg-blue-950/20" : "border-gray-700 hover:border-gray-500"}`}
            onClick={() => document.getElementById("file-input")?.click()}
          >
            <input
              id="file-input"
              type="file"
              accept=".xlsx,.xls"
              className="hidden"
              onChange={onInputChange}
            />
            {loading ? (
              <p className="text-gray-400">Parsing file...</p>
            ) : (
              <>
                <p className="text-lg font-medium mb-2">
                  Drop summary Excel here
                </p>
                <p className="text-gray-500 text-sm">or click to browse</p>
              </>
            )}
          </div>
        )}

        {error && <p className="mt-4 text-red-400 text-sm">{error}</p>}

        {step === "preview" && (
          <div>
            <div className="flex items-center justify-between mb-4">
              <div>
                <p className="font-semibold">{monthLabel}</p>
                <p className="text-gray-400 text-sm">
                  {groups.length} drivers ·{" "}
                  {groups.reduce((s, g) => s + g.trips.length, 0)} trips
                </p>
              </div>
              <button
                onClick={reset}
                className="text-sm text-gray-500 hover:text-gray-300"
              >
                Upload different file
              </button>
            </div>

            <div className="rounded-xl border border-gray-800 overflow-hidden mb-6">
              <table className="w-full text-sm">
                <thead className="bg-gray-900 text-gray-400">
                  <tr>
                    <th className="text-left px-4 py-3">Driver</th>
                    <th className="text-left px-4 py-3">Plate#</th>
                    <th className="text-left px-4 py-3">Truck#</th>
                    <th className="text-right px-4 py-3">Trips</th>
                    <th className="text-right px-4 py-3">Total Variance</th>
                  </tr>
                </thead>
                <tbody>
                  {[...groups]
                    .sort((a, b) => a.totalVariance - b.totalVariance)
                    .map((g, i) => (
                      <tr
                        key={i}
                        className="border-t border-gray-800 hover:bg-gray-900/50"
                      >
                        <td className="px-4 py-3">{g.driver}</td>
                        <td className="px-4 py-3 text-gray-400">{g.plate}</td>
                        <td className="px-4 py-3 text-gray-400">{g.truck}</td>
                        <td className="px-4 py-3 text-right text-gray-400">
                          {g.trips.length}
                        </td>
                        <td
                          className={`px-4 py-3 text-right font-mono font-medium
                          ${g.totalVariance < 0 ? "text-green-400" : "text-red-400"}`}
                        >
                          {g.totalVariance > 0 ? "+" : ""}
                          {g.totalVariance.toFixed(4)}
                        </td>
                      </tr>
                    ))}
                </tbody>
              </table>
            </div>

            <button
              onClick={onDownload}
              className="w-full bg-blue-600 hover:bg-blue-500 text-white font-semibold py-3 rounded-xl transition-colors"
            >
              Download Report Excel
            </button>
          </div>
        )}

        {step === "done" && (
          <div className="text-center py-16">
            <p className="text-xl font-semibold mb-2">Report downloaded.</p>
            <p className="text-gray-400 text-sm mb-8">
              {monthLabel} — {groups.length} driver sheets + ranking
            </p>
            <button
              onClick={reset}
              className="text-blue-400 hover:text-blue-300 text-sm"
            >
              Process another file
            </button>
          </div>
        )}
      </div>
    </div>
  );
}
