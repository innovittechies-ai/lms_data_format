export const OUTPUT_HEADERS = [
  "Student Name",
  "Student Email",
  "Student Mobile",
  "Guardian Number 1",
  "Guardian Relation 1",
  "Guardian Number  2",
  "Guardian Relation 2",
  "Highest Qualification",
  "Year Of Passing",
  "College name",
] as const;

export type OutputKey = (typeof OUTPUT_HEADERS)[number];

export type OutputRow = Record<OutputKey, string>;

export type FieldConfig = {
  mode: "column" | "constant";
  column: string;
  value: string;
};

export type ConvertConfig = {
  fields: Record<OutputKey, FieldConfig>;
  removeDuplicates: boolean;
  skipIncomplete: boolean;
};

export const COLUMN_ALIASES: Record<OutputKey, string[]> = {
  "Student Name": ["full name", "student name", "name"],
  "Student Email": ["student email", "email id", "email", "e mail", "mail id"],
  "Student Mobile": [
    "student mobile",
    "phone number",
    "mobile number",
    "mobile",
    "phone",
    "contact number",
    "contact",
  ],
  "Guardian Number 1": [
    "guardian number 1",
    "guardian phone",
    "guardian mobile",
    "parent phone",
    "father number",
    "father phone",
  ],
  "Guardian Relation 1": ["guardian relation 1", "relation 1", "guardian relation"],
  "Guardian Number  2": [
    "guardian number 2",
    "guardian number  2",
    "mother number",
    "mother phone",
  ],
  "Guardian Relation 2": ["guardian relation 2", "relation 2"],
  "Highest Qualification": [
    "highest qualification",
    "qualification",
    "degree",
    "course",
  ],
  "Year Of Passing": ["year of passing", "passing year", "yop", "year"],
  "College name": ["college name", "college", "institute", "institution", "university"],
};

function norm(value: string) {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

export function guessColumn(headers: string[], aliases: string[]) {
  const normalized = headers.map(norm);

  for (const alias of aliases) {
    const index = normalized.findIndex((header) => header === norm(alias));
    if (index >= 0) return headers[index];
  }

  for (const alias of aliases) {
    const needle = norm(alias);
    const index = normalized.findIndex((header) => {
      if (header.includes("registration") || header.includes("id number")) {
        return false;
      }
      return header.includes(needle);
    });
    if (index >= 0) return headers[index];
  }

  return "";
}

export function defaultFields(headers: string[]): Record<OutputKey, FieldConfig> {
  const mapped = OUTPUT_HEADERS.reduce(
    (acc, key) => {
      const column = guessColumn(headers, COLUMN_ALIASES[key]);
      const fromFile = key === "Student Name" || key === "Student Email" || key === "Student Mobile";
      acc[key] = {
        mode: column && fromFile ? "column" : "constant",
        column,
        value: "",
      };
      return acc;
    },
    {} as Record<OutputKey, FieldConfig>,
  );

  if (mapped["Student Name"].column) mapped["Student Name"].mode = "column";
  if (mapped["Student Email"].column) mapped["Student Email"].mode = "column";
  if (mapped["Student Mobile"].column) mapped["Student Mobile"].mode = "column";

  return mapped;
}

export function cleanName(value: unknown) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

export function cleanEmail(value: unknown) {
  return String(value ?? "")
    .replace(/\s+/g, "")
    .trim()
    .toLowerCase();
}

export function cleanPhone(value: unknown) {
  if (value == null || value === "") return "";

  let text = String(value).trim();
  const numeric = Number(value);

  if (Number.isFinite(numeric) && (/e\+/i.test(text) || /^\d+\.\d+$/.test(text) || typeof value === "number")) {
    text = Math.round(numeric).toString();
  }

  const digits = text.replace(/\D/g, "");
  if (digits.length === 12 && digits.startsWith("91")) return digits.slice(2);
  if (digits.length === 11 && digits.startsWith("0")) return digits.slice(1);
  return digits;
}

export function cleanText(value: unknown) {
  return String(value ?? "")
    .replace(/\s+/g, " ")
    .trim();
}

export function cleanYear(value: unknown) {
  if (value == null || value === "") return "";
  const numeric = Number(value);
  if (Number.isFinite(numeric) && numeric > 1900 && numeric < 2100) {
    return String(Math.round(numeric));
  }
  return cleanText(value);
}

function readField(row: Record<string, unknown>, field: FieldConfig, key: OutputKey) {
  const raw = field.mode === "column" && field.column ? row[field.column] : field.value;

  if (key === "Student Name") return cleanName(raw);
  if (key === "Student Email") return cleanEmail(raw);
  if (key === "Student Mobile" || key === "Guardian Number 1" || key === "Guardian Number  2") {
    return cleanPhone(raw);
  }
  if (key === "Year Of Passing") return cleanYear(raw);
  return cleanText(raw);
}

export function convertRows(source: Record<string, unknown>[], config: ConvertConfig) {
  const unique = new Map<string, OutputRow>();
  const rows: OutputRow[] = [];

  for (const sourceRow of source) {
    const row = OUTPUT_HEADERS.reduce((acc, key) => {
      acc[key] = readField(sourceRow, config.fields[key], key);
      return acc;
    }, {} as OutputRow);

    if (config.skipIncomplete && (!row["Student Name"] || !row["Student Email"])) {
      continue;
    }

    if (config.removeDuplicates && row["Student Email"]) {
      unique.set(row["Student Email"], row);
    } else {
      rows.push(row);
    }
  }

  return config.removeDuplicates ? Array.from(unique.values()) : rows;
}

export async function parseWorkbook(file: File) {
  const XLSX = await import("xlsx");
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheetName = workbook.SheetNames[0];
  const sheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json<Record<string, unknown>>(sheet, {
    defval: "",
    raw: true,
  });
  const headerRow = XLSX.utils.sheet_to_json<(string | number)[]>(sheet, {
    header: 1,
    defval: "",
  })[0] as (string | number)[] | undefined;
  const headers = headerRow
    ? headerRow.map((cell) => String(cell).trim()).filter(Boolean)
    : rows[0]
      ? Object.keys(rows[0])
      : [];

  return { headers, rows, sheetName };
}

export async function downloadTemplate(rows: OutputRow[], filename: string) {
  const XLSX = await import("xlsx");
  const sheet = XLSX.utils.json_to_sheet(rows, {
    header: [...OUTPUT_HEADERS],
  });
  sheet["!cols"] = [
    { wch: 22 },
    { wch: 28 },
    { wch: 16 },
    { wch: 18 },
    { wch: 20 },
    { wch: 18 },
    { wch: 20 },
    { wch: 22 },
    { wch: 16 },
    { wch: 28 },
  ];
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "Sheet1");
  XLSX.writeFile(workbook, filename);
}
