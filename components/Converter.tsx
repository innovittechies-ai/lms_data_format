"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import {
  OUTPUT_HEADERS,
  convertRows,
  defaultFields,
  downloadTemplate,
  parseWorkbook,
  type ConvertConfig,
  type FieldConfig,
  type OutputKey,
  type OutputRow,
} from "@/lib/convert";

const BULK_FIELDS: OutputKey[] = [
  "College name",
  "Highest Qualification",
  "Year Of Passing",
  "Guardian Number 1",
  "Guardian Relation 1",
  "Guardian Number  2",
  "Guardian Relation 2",
];

const FILE_FIELDS: OutputKey[] = ["Student Name", "Student Email", "Student Mobile"];

const STORAGE_KEY = "lms-bulk-defaults";

const HINTS: Record<OutputKey, string> = {
  "Student Name": "Usually Full Name",
  "Student Email": "Usually Email",
  "Student Mobile": "Usually Phone number",
  "Guardian Number 1": "Copied into every row",
  "Guardian Relation 1": "Copied into every row",
  "Guardian Number  2": "Copied into every row",
  "Guardian Relation 2": "Copied into every row",
  "Highest Qualification": "Copied into every row",
  "Year Of Passing": "Copied into every row",
  "College name": "Copied into every row",
};

function emptyFields(): Record<OutputKey, FieldConfig> {
  return OUTPUT_HEADERS.reduce(
    (acc, key) => {
      acc[key] = { mode: "constant", column: "", value: "" };
      return acc;
    },
    {} as Record<OutputKey, FieldConfig>,
  );
}

function loadSavedValues() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return {} as Partial<Record<OutputKey, string>>;
    return JSON.parse(raw) as Partial<Record<OutputKey, string>>;
  } catch {
    return {} as Partial<Record<OutputKey, string>>;
  }
}

export default function Converter() {
  const inputRef = useRef<HTMLInputElement>(null);
  const [fileName, setFileName] = useState("");
  const [headers, setHeaders] = useState<string[]>([]);
  const [sourceRows, setSourceRows] = useState<Record<string, unknown>[]>([]);
  const [fields, setFields] = useState<Record<OutputKey, FieldConfig>>(emptyFields);
  const [removeDuplicates, setRemoveDuplicates] = useState(true);
  const [skipIncomplete, setSkipIncomplete] = useState(true);
  const [error, setError] = useState("");
  const [dragging, setDragging] = useState(false);
  const [busy, setBusy] = useState(false);
  const [hydrated, setHydrated] = useState(false);

  useEffect(() => {
    const saved = loadSavedValues();
    if (Object.keys(saved).length) {
      setFields((current) => {
        const next = { ...current };
        for (const key of BULK_FIELDS) {
          if (saved[key]) next[key] = { ...next[key], value: saved[key] as string };
        }
        return next;
      });
    }
    setHydrated(true);
  }, []);

  useEffect(() => {
    if (!hydrated) return;
    const values = BULK_FIELDS.reduce(
      (acc, key) => {
        acc[key] = fields[key].value;
        return acc;
      },
      {} as Partial<Record<OutputKey, string>>,
    );
    localStorage.setItem(STORAGE_KEY, JSON.stringify(values));
  }, [fields, hydrated]);

  const config: ConvertConfig = {
    fields,
    removeDuplicates,
    skipIncomplete,
  };

  const preview = useMemo(
    () => convertRows(sourceRows, config),
    [sourceRows, fields, removeDuplicates, skipIncomplete],
  );

  async function handleFile(file: File) {
    setError("");
    setBusy(true);
    try {
      const parsed = await parseWorkbook(file);
      if (!parsed.headers.length) {
        throw new Error("No columns found in the first sheet.");
      }
      const mapped = defaultFields(parsed.headers);
      const saved = loadSavedValues();
      for (const key of BULK_FIELDS) {
        mapped[key].value = saved[key] || mapped[key].value;
        if (mapped[key].column && !FILE_FIELDS.includes(key)) {
          mapped[key].mode = "constant";
        }
      }
      setFileName(file.name);
      setHeaders(parsed.headers);
      setSourceRows(parsed.rows);
      setFields(mapped);
    } catch (err) {
      setError(err instanceof Error ? err.message : "Could not read that file.");
      setFileName("");
      setHeaders([]);
      setSourceRows([]);
    } finally {
      setBusy(false);
    }
  }

  function updateField(key: OutputKey, patch: Partial<FieldConfig>) {
    setFields((current) => ({
      ...current,
      [key]: { ...current[key], ...patch },
    }));
  }

  async function onDownload() {
    if (!preview.length) return;
    setBusy(true);
    try {
      await downloadTemplate(preview, "Classroom-Students.xlsx");
    } finally {
      setBusy(false);
    }
  }

  return (
    <main className="mx-auto max-w-6xl px-4 py-8 sm:py-12">
      <header className="mb-8 max-w-3xl">
        <p className="text-sm font-medium tracking-wide text-teal uppercase">LMS formatter</p>
        <h1 className="font-display mt-2 text-4xl leading-tight text-ink sm:text-5xl">
          Classroom student sheet
        </h1>
        <p className="mt-3 max-w-2xl text-muted">
          Upload an onboarding Excel file. Name, email and phone come from the file.
          Fill the rest once and it is applied to every student row, in the same
          Classroom-Student template columns.
        </p>
      </header>

      <div className="grid gap-6 lg:grid-cols-[minmax(0,1.1fr)_minmax(0,0.9fr)]">
        <section className="rounded-2xl border border-line bg-card p-5 shadow-sm sm:p-6">
          <h2 className="text-lg font-semibold">1. Upload file</h2>
          <p className="mt-1 text-sm text-muted">Excel or CSV. First sheet is used.</p>

          <label
            onDragOver={(event) => {
              event.preventDefault();
              setDragging(true);
            }}
            onDragLeave={() => setDragging(false)}
            onDrop={(event) => {
              event.preventDefault();
              setDragging(false);
              const file = event.dataTransfer.files[0];
              if (file) void handleFile(file);
            }}
            className={`mt-4 flex cursor-pointer flex-col items-center justify-center rounded-xl border-2 border-dashed px-4 py-10 text-center transition ${
              dragging ? "border-teal bg-teal/5" : "border-line bg-paper/60"
            }`}
          >
            <input
              ref={inputRef}
              type="file"
              accept=".xlsx,.xls,.csv"
              className="hidden"
              onChange={(event) => {
                const file = event.target.files?.[0];
                if (file) void handleFile(file);
              }}
            />
            <span className="text-base font-medium text-ink">
              {busy ? "Reading file..." : "Drop file here or click to browse"}
            </span>
            <span className="mt-1 text-sm text-muted">
              {fileName || "Works with Infinity Onboarding responses and similar sheets"}
            </span>
          </label>

          {error ? <p className="mt-3 text-sm text-amber">{error}</p> : null}

          {sourceRows.length > 0 ? (
            <p className="mt-3 text-sm text-muted">
              {sourceRows.length} rows found · {headers.length} columns · {preview.length} students after cleanup
            </p>
          ) : null}

          {headers.length > 0 ? (
            <div className="mt-6 space-y-4">
              <h3 className="text-base font-semibold">Map student fields</h3>
              {FILE_FIELDS.map((key) => (
                <FieldRow
                  key={key}
                  label={key}
                  hint={HINTS[key]}
                  headers={headers}
                  field={fields[key]}
                  onChange={(patch) => updateField(key, patch)}
                  forceColumn
                />
              ))}
            </div>
          ) : null}
        </section>

        <section className="rounded-2xl border border-line bg-card p-5 shadow-sm sm:p-6">
          <h2 className="text-lg font-semibold">2. Apply to every row</h2>
          <p className="mt-1 text-sm text-muted">
            These values stay the same for the whole batch. Change them only when the next college or batch is different.
          </p>

          <div className="mt-5 grid gap-4 sm:grid-cols-2">
            {BULK_FIELDS.map((key) => (
              <FieldRow
                key={key}
                label={key.replace("  ", " ")}
                hint={HINTS[key]}
                headers={headers}
                field={fields[key]}
                onChange={(patch) => updateField(key, patch)}
              />
            ))}
          </div>

          <div className="mt-5 flex flex-col gap-2 text-sm">
            <label className="flex items-center gap-2">
              <input
                type="checkbox"
                checked={removeDuplicates}
                onChange={(event) => setRemoveDuplicates(event.target.checked)}
              />
              Remove duplicate emails (keeps the latest row)
            </label>
            <label className="flex items-center gap-2">
              <input
                type="checkbox"
                checked={skipIncomplete}
                onChange={(event) => setSkipIncomplete(event.target.checked)}
              />
              Skip rows without name or email
            </label>
          </div>
        </section>
      </div>

      <section className="mt-6 rounded-2xl border border-line bg-card p-5 shadow-sm sm:p-6">
        <div className="flex flex-wrap items-center justify-between gap-3">
          <div>
            <h2 className="text-lg font-semibold">3. Preview and download</h2>
            <p className="mt-1 text-sm text-muted">
              Output columns match the Classroom-Student template exactly.
            </p>
          </div>
          <button
            type="button"
            onClick={() => void onDownload()}
            disabled={!preview.length || busy}
            className="rounded-full bg-teal px-5 py-2.5 text-sm font-semibold text-white disabled:cursor-not-allowed disabled:opacity-50 hover:bg-teal-dark"
          >
            Download Classroom-Students.xlsx
          </button>
        </div>

        {preview.length ? (
          <div className="mt-4 overflow-auto rounded-xl border border-line">
            <table className="min-w-full text-left text-sm">
              <thead className="bg-paper">
                <tr>
                  {OUTPUT_HEADERS.map((header) => (
                    <th key={header} className="whitespace-nowrap px-3 py-2 font-semibold">
                      {header.replace("  ", " ")}
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {preview.slice(0, 40).map((row: OutputRow, index) => (
                  <tr key={`${row["Student Email"]}-${index}`} className="border-t border-line">
                    {OUTPUT_HEADERS.map((header) => (
                      <td key={header} className="whitespace-nowrap px-3 py-2">
                        {row[header] || "—"}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        ) : (
          <p className="mt-4 text-sm text-muted">Upload a file to see the formatted preview.</p>
        )}

        {preview.length > 40 ? (
          <p className="mt-2 text-sm text-muted">Showing first 40 of {preview.length} students.</p>
        ) : null}
      </section>
    </main>
  );
}

function FieldRow({
  label,
  hint,
  headers,
  field,
  onChange,
  forceColumn = false,
}: {
  label: string;
  hint: string;
  headers: string[];
  field: FieldConfig;
  onChange: (patch: Partial<FieldConfig>) => void;
  forceColumn?: boolean;
}) {
  const usingColumn = field.mode === "column" && Boolean(field.column);

  return (
    <label className="block text-sm">
      <span className="font-medium text-ink">{label}</span>
      <span className="ml-2 text-muted">{hint}</span>
      {forceColumn ? (
        <select
          className="mt-1 w-full rounded-lg border border-line bg-white px-3 py-2"
          value={field.column}
          onChange={(event) => onChange({ mode: "column", column: event.target.value })}
          disabled={!headers.length}
        >
          <option value="">{headers.length ? "Select column" : "Upload a file first"}</option>
          {headers.map((header) => (
            <option key={header} value={header}>
              {header}
            </option>
          ))}
        </select>
      ) : (
        <div className="mt-1 space-y-2">
          <input
            className="w-full rounded-lg border border-line bg-white px-3 py-2"
            value={usingColumn ? "" : field.value}
            placeholder={`Same value for all ${label.toLowerCase()}`}
            disabled={usingColumn}
            onChange={(event) => onChange({ mode: "constant", value: event.target.value })}
          />
          {headers.length ? (
            <select
              className="w-full rounded-lg border border-line bg-white px-3 py-2 text-muted"
              value={usingColumn ? field.column : ""}
              onChange={(event) => {
                if (event.target.value) {
                  onChange({ mode: "column", column: event.target.value });
                } else {
                  onChange({ mode: "constant", column: "" });
                }
              }}
            >
              <option value="">Use the value above for every row</option>
              {headers.map((header) => (
                <option key={header} value={header}>
                  Take from column: {header}
                </option>
              ))}
            </select>
          ) : null}
        </div>
      )}
    </label>
  );
}
