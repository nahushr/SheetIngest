import type { ReactNode } from "react";
import type { DialogProps } from "@mui/material/Dialog";
import type {
  SuperDataGridColumnOptions,
  SuperDataGridColumnType,
} from "@simplishelf/super-data-grid";

export type SheetIngestScalar = string | number | boolean | Date | null | undefined;
export type SheetIngestRow = object;
export type SheetIngestMaybePromise<T> = T | Promise<T>;

export interface SheetIngestField {
  /** Stable key returned to the host application. */
  key: string;
  /** Label shown in the mapper and preview grid. */
  label: string;
  /** Header written to the generated template. Defaults to `key`. */
  templateHeader?: string;
  /** Other names that should automatically map to this field. */
  aliases?: readonly string[];
  /** Marks the target field as required during review. */
  required?: boolean;
  /** Optional example shown beside this target field. */
  example?: string;
  /** Optional cell renderer supplied by SuperDataGrid. */
  type?: SuperDataGridColumnType;
  /** Options for built-in preview cells such as email links and phone flags. */
  columnOptions?: SuperDataGridColumnOptions;
  /** Width for this field in the preview table. */
  width?: number;
  /** Exclude the field from the preview grid while still allowing mapping. */
  hideInPreview?: boolean;
}

export interface SheetIngestGroup {
  /** Category name used for the merged first row in template workbooks. */
  label: string;
  fields: readonly SheetIngestField[];
}

export interface SheetIngestTemplate {
  /** Suggested name for the workbook sheet. */
  sheetName?: string;
  /** Download file name, including the `.xlsx` extension. */
  fileName: string;
  groups: readonly SheetIngestGroup[];
}

export interface SheetIngestIssue {
  message: string;
  field?: string;
  severity?: "error" | "warning" | "info";
}

export interface SheetIngestRowContext {
  rowNumber: number;
  file: File;
  sheetName: string;
  /** Original values from the uploaded spreadsheet, keyed by target field. */
  values: Readonly<Record<string, string>>;
}

export interface SheetIngestStepContext {
  step: "upload" | "header" | "mapping" | "preview";
  file: File | null;
  rowCount: number;
}

export interface SheetIngestProps<Row extends SheetIngestRow = SheetIngestRow> {
  /** Controls the import dialog. */
  isOpen: boolean;
  /** Called when the dialog closes or the import succeeds. */
  onClose: () => void;
  /** Entity-specific header groups and output fields. */
  template: SheetIngestTemplate;
  /** Handles the final import in the host application. */
  onSubmit: (rows: readonly Row[], file: File) => SheetIngestMaybePromise<void>;
  /** Maps raw strings into the host's entity model before validation. */
  mapRow?: (
    values: Readonly<Record<string, string>>,
    context: SheetIngestRowContext,
  ) => SheetIngestMaybePromise<Row>;
  /** Runs host-specific validation and returns issues for this row. */
  validateRow?: (
    row: Row,
    context: SheetIngestRowContext,
  ) => SheetIngestMaybePromise<void | string | SheetIngestIssue | readonly (string | SheetIngestIssue)[]>;
  /** Generate and download entity-specific realistic data in the host app. */
  onGenerateTestData?: (recordCount: number) => SheetIngestMaybePromise<void>;
  /** Optional custom content rendered below the active step. */
  renderStepContent?: (context: SheetIngestStepContext) => ReactNode;
  /** Maximum non-empty data rows to parse. Defaults to 1,000. */
  maxRecords?: number;
  /** Maximum upload size in bytes. Defaults to 10 MiB. */
  maxFileSize?: number;
  /** Default test-data count in the upload step. Defaults to 50. */
  defaultTestDataCount?: number;
  /** Maximum accepted test-data count. Defaults to 10,000. */
  maxTestDataCount?: number;
  /** Permit submission when one or more rows have validation errors. Defaults to false. */
  allowInvalidSubmit?: boolean;
  /** Override the wizard title. */
  title?: string;
  /** Override dialog sizing and styling. */
  dialogProps?: Omit<Partial<DialogProps>, "open" | "onClose" | "children">;
  /** Localized labels for common controls. */
  translations?: Partial<SheetIngestTranslations>;
}

export interface SheetIngestTranslations {
  title: string;
  steps: {
    upload: string;
    header: string;
    mapping: string;
    preview: string;
  };
  downloadTemplate: string;
  generateTestData: string;
  uploadFile: string;
  chooseFile: string;
  selectHeaderRow: string;
  matchColumns: string;
  dataPreview: string;
  back: string;
  next: string;
  submit: string;
  cancel: string;
  noMapping: string;
  ignoreColumn: string;
  required: string;
  valid: string;
  invalid: string;
  row: string;
  status: string;
  errorTitle: string;
  close: string;
  records: string;
  testRecordCount: string;
  fileHelp: string;
  headerHelp: string;
  mappingHelp: string;
  previewHelp: string;
}
