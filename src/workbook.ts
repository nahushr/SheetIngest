import ExcelJS from "exceljs";
import Papa from "papaparse";
import type { SheetIngestField, SheetIngestTemplate } from "./types";

export type SheetIngestMatrix = string[][];

export interface WorkbookSheet {
  name: string;
  rows: SheetIngestMatrix;
}

const stringifyCellValue = (value: unknown): string => {
  if (value == null) return "";
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? "" : value.toISOString();
  }
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  try {
    return JSON.stringify(value) ?? "";
  } catch {
    return "";
  }
};

const cellToString = (cell: ExcelJS.Cell): string => {
  if (cell.isMerged && cell.master.address !== cell.address) return "";
  const value = cell.value;
  if (value == null) return "";
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? "" : value.toISOString();
  }
  if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return String(value);
  }
  if ("richText" in value && Array.isArray(value.richText)) {
    return value.richText.map((part) => part.text).join("");
  }
  if ("text" in value && typeof value.text === "string") return value.text;
  if ("result" in value) return stringifyCellValue(value.result);
  return cell.text;
};

const readCsvFile = async (file: File): Promise<WorkbookSheet> => {
  const parsed = Papa.parse<string[]>(await file.text(), {
    skipEmptyLines: "greedy",
    dynamicTyping: false,
  });
  if (parsed.errors.length > 0) {
    const firstError = parsed.errors[0];
    throw new Error(`CSV parsing failed on row ${(firstError.row ?? 0) + 1}: ${firstError.message}`);
  }
  return {
    name: file.name.replace(/\.csv$/i, "") || "CSV",
    rows: parsed.data.map((row) => row.map((cell) => cell ?? "")),
  };
};

const readXlsxFile = async (file: File): Promise<WorkbookSheet> => {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(await file.arrayBuffer());
  const worksheet = workbook.worksheets[0];
  if (!worksheet) throw new Error("The uploaded file does not contain a worksheet.");

  const rows: string[][] = [];
  for (let rowNumber = 1; rowNumber <= worksheet.rowCount; rowNumber += 1) {
    const row = worksheet.getRow(rowNumber);
    const values: string[] = [];
    for (let columnNumber = 1; columnNumber <= worksheet.columnCount; columnNumber += 1) {
      values[columnNumber - 1] = cellToString(row.getCell(columnNumber));
    }
    rows[rowNumber - 1] = values;
  }
  return { name: worksheet.name, rows };
};

/** Read the first worksheet from a CSV or XLSX browser File. */
export const readSpreadsheetFile = async (file: File): Promise<WorkbookSheet> => {
  if (/\.csv$/i.test(file.name)) return readCsvFile(file);
  if (!/\.xlsx$/i.test(file.name)) {
    throw new Error("Choose an XLSX workbook or CSV file.");
  }
  return readXlsxFile(file);
};

/** Normalize spreadsheet labels for matching, ignoring punctuation and accents. */
export const normalizeHeader = (value: string): string =>
  value
    .normalize("NFD")
    .replace(/\p{Diacritic}/gu, "")
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "");

export const getTemplateFields = (
  template: SheetIngestTemplate,
): SheetIngestField[] => template.groups.flatMap((group) => group.fields);

const headerAliases = (field: SheetIngestField): string[] =>
  [field.key, field.label, field.templateHeader, ...(field.aliases ?? [])]
    .filter((item): item is string => Boolean(item?.trim()))
    .map(normalizeHeader);

const isFieldMatch = (header: string, field: SheetIngestField): boolean => {
  const normalized = normalizeHeader(header);
  return normalized.length > 0 && headerAliases(field).includes(normalized);
};

/** Select a likely header row by counting exact and alias matches. */
export const detectHeaderRow = (
  rows: readonly (readonly string[])[],
  fields: readonly SheetIngestField[],
  scanRows = 25,
): number => {
  let bestRow = 0;
  let bestScore = -1;
  const limit = Math.min(rows.length, Math.max(scanRows, 1));

  for (let rowIndex = 0; rowIndex < limit; rowIndex += 1) {
    const row = rows[rowIndex] ?? [];
    const nonEmptyHeaders = row.map((cell) => cell.trim()).filter(Boolean);
    if (nonEmptyHeaders.length === 0) continue;

    let score = 0;
    const matchedFields = new Set<string>();
    for (const header of nonEmptyHeaders) {
      const match = fields.find((field) => isFieldMatch(header, field));
      if (match && !matchedFields.has(match.key)) {
        matchedFields.add(match.key);
        score += 4;
      }
    }
    // Prefer the most complete candidate when no template aliases match. This
    // favors the actual field-header row over a sparse grouped-heading row.
    score += nonEmptyHeaders.length / 100;

    if (score > bestScore) {
      bestScore = score;
      bestRow = rowIndex;
    }
  }

  return bestRow;
};

const editDistance = (left: string, right: string): number => {
  const previous = Array.from({ length: right.length + 1 }, (_, index) => index);
  for (let leftIndex = 1; leftIndex <= left.length; leftIndex += 1) {
    const current = [leftIndex];
    for (let rightIndex = 1; rightIndex <= right.length; rightIndex += 1) {
      const substitutionCost =
        left[leftIndex - 1] === right[rightIndex - 1] ? 0 : 1;
      current[rightIndex] = Math.min(
        current[rightIndex - 1] + 1,
        previous[rightIndex] + 1,
        previous[rightIndex - 1] + substitutionCost,
      );
    }
    for (let index = 0; index < current.length; index += 1) {
      previous[index] = current[index];
    }
  }
  return previous[right.length];
};

const similarity = (header: string, alias: string): number => {
  if (!header || !alias) return 0;
  if (header === alias) return 1;
  if (header.includes(alias) || alias.includes(header)) {
    return Math.min(header.length, alias.length) / Math.max(header.length, alias.length);
  }
  return 1 - editDistance(header, alias) / Math.max(header.length, alias.length);
};

/** Suggest one-to-one target-to-source mappings from a selected header row. */
export const suggestColumnMappings = (
  headers: readonly string[],
  fields: readonly SheetIngestField[],
): Record<string, number | null> => {
  const candidates = fields.flatMap((field) =>
    headers.map((header, sourceIndex) => {
      const normalizedHeader = normalizeHeader(header);
      const bestScore = Math.max(
        0,
        ...headerAliases(field).map((alias) => similarity(normalizedHeader, alias)),
      );
      return { fieldKey: field.key, sourceIndex, score: bestScore };
    }),
  );

  const result: Record<string, number | null> = Object.fromEntries(
    fields.map((field) => [field.key, null]),
  );
  const assignedFields = new Set<string>();
  const assignedColumns = new Set<number>();

  candidates
    .filter((candidate) => candidate.score >= 0.6)
    .sort((left, right) => right.score - left.score)
    .forEach((candidate) => {
      if (
        assignedFields.has(candidate.fieldKey) ||
        assignedColumns.has(candidate.sourceIndex)
      ) {
        return;
      }
      result[candidate.fieldKey] = candidate.sourceIndex;
      assignedFields.add(candidate.fieldKey);
      assignedColumns.add(candidate.sourceIndex);
    });

  return result;
};

export interface SourceColumn {
  index: number;
  header: string;
  group: string;
  sample: string;
}

/** Build source-column labels, carrying merged group headings across blank cells. */
export const getSourceColumns = (
  rows: readonly (readonly string[])[],
  headerRows: number | readonly number[],
): SourceColumn[] => {
  // Keep the original numeric API behavior: a numeric header row uses the row
  // above it as its group row. Array input represents the exact selected levels.
  let requestedRows: number[];
  if (typeof headerRows === "number") {
    requestedRows = [headerRows];
    if (headerRows > 0) requestedRows.unshift(headerRows - 1);
  } else {
    requestedRows = [...headerRows];
  }
  const selectedRows = [...new Set(requestedRows)]
    .filter((index) => Number.isInteger(index) && index >= 0 && index < rows.length)
    .sort((left, right) => left - right);
  const lastHeaderRow = selectedRows.at(-1);
  if (lastHeaderRow === undefined) return [];
  const groupRows = selectedRows.slice(0, -1).map((rowIndex) => rows[rowIndex] ?? []);
  const headings = rows[lastHeaderRow] ?? [];
  const maxColumns = Math.max(
    ...selectedRows.map((rowIndex) => rows[rowIndex]?.length ?? 0),
    ...rows.slice(lastHeaderRow + 1).map((row) => row.length),
  );
  const groups = groupRows.map((groupRow) => {
    const level: string[] = [];
    let activeGroup = "";
    for (let column = 0; column < maxColumns; column += 1) {
      const value = groupRow[column]?.trim() ?? "";
      if (value) activeGroup = value;
      level[column] = activeGroup;
    }
    return level;
  });

  return Array.from({ length: maxColumns }, (_, index) => ({
    index,
    header: headings[index]?.trim() ?? "",
    group: groups.map((level) => level[index]).filter(Boolean).join(" / "),
    sample: rows[lastHeaderRow + 1]?.[index] ?? "",
  })).filter((column) => {
    if (column.header || column.group || column.sample) return true;
    return rows.slice(lastHeaderRow + 2).some((row) => Boolean(row[column.index]?.trim()));
  });
};

const toWorkbookCellValue = (value: unknown): string | number | boolean | Date => {
  if (value == null) return "";
  if (value instanceof Date || typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
    return value;
  }
  try {
    return JSON.stringify(value) ?? "";
  } catch {
    return "";
  }
};

/** Generate the grouped, merged-header template, optionally with host-generated records. */
export const createTemplateWorkbook = (
  template: SheetIngestTemplate,
  rows: readonly Readonly<Record<string, unknown>>[] = [],
): ExcelJS.Workbook => {
  const workbook = new ExcelJS.Workbook();
  const worksheet = workbook.addWorksheet((template.sheetName ?? "Template").slice(0, 31), {
    views: [{ state: "frozen", ySplit: 2 }],
  });
  const groupHeaders: string[] = [];
  const fieldHeaders: string[] = [];
  const mergeRanges: Array<[number, number]> = [];
  let column = 1;

  template.groups.forEach((group) => {
    if (group.fields.length === 0) return;
    const startColumn = column;
    group.fields.forEach((field, fieldIndex) => {
      groupHeaders.push(fieldIndex === 0 ? group.label : "");
      fieldHeaders.push(field.templateHeader ?? field.key);
      column += 1;
    });
    if (group.fields.length > 1) mergeRanges.push([startColumn, column - 1]);
  });

  worksheet.addRow(groupHeaders);
  worksheet.addRow(fieldHeaders);
  mergeRanges.forEach(([startColumn, endColumn]) => {
    worksheet.mergeCells(1, startColumn, 1, endColumn);
  });
  worksheet.columns = fieldHeaders.map((header) => ({
    width: Math.max(14, Math.min(36, header.length + 4)),
  }));
  worksheet.getRow(1).height = 26;
  worksheet.getRow(2).height = 24;
  for (let columnNumber = 1; columnNumber <= fieldHeaders.length; columnNumber += 1) {
    const categoryCell = worksheet.getCell(1, columnNumber);
    categoryCell.font = { bold: true, size: 12 };
    categoryCell.alignment = { horizontal: "center", vertical: "middle" };
    categoryCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD3D3D3" },
    };

    const fieldCell = worksheet.getCell(2, columnNumber);
    fieldCell.font = { bold: true };
    fieldCell.alignment = { horizontal: "center", vertical: "middle" };
    fieldCell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFE8E8E8" },
    };
  }
  const fields = getTemplateFields(template);
  rows.forEach((row) => {
    worksheet.addRow(fields.map((field) => toWorkbookCellValue(row[field.key])));
  });
  return workbook;
};

/** Download a template workbook, optionally prefilled with host-generated rows. */
export const downloadTemplate = async (
  template: SheetIngestTemplate,
  rows: readonly Readonly<Record<string, unknown>>[] = [],
): Promise<void> => {
  if (typeof document === "undefined") {
    throw new TypeError("Template downloads are only available in a browser.");
  }
  const workbook = createTemplateWorkbook(template, rows);
  const buffer = await workbook.xlsx.writeBuffer();
  const bytes = new Uint8Array(buffer.byteLength);
  bytes.set(buffer as unknown as Uint8Array);
  const url = URL.createObjectURL(new Blob([bytes], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  }));
  const link = document.createElement("a");
  link.href = url;
  link.download = template.fileName;
  link.click();
  window.setTimeout(() => URL.revokeObjectURL(url), 0);
};
