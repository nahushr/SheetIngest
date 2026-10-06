import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type React from "react";
import {
  Alert,
  Box,
  Button,
  Checkbox,
  Chip,
  Dialog,
  DialogActions,
  DialogContent,
  DialogTitle,
  Divider,
  FormControl,
  InputLabel,
  LinearProgress,
  MenuItem,
  Paper,
  Select,
  Stack,
  Step,
  StepLabel,
  Stepper,
  TextField,
  Tooltip,
  Typography,
} from "@mui/material";
import { DataGrid } from "@mui/x-data-grid";
import type { GridColDef, GridColumnGroupingModel } from "@mui/x-data-grid";
import {
  AutoAwesome as AutoAwesomeIcon,
  CheckCircleOutline as CheckCircleOutlineIcon,
  CloudUpload as CloudUploadIcon,
  Download as DownloadIcon,
  ErrorOutline as ErrorOutlineIcon,
  FilePresent as FilePresentIcon,
  NavigateBefore as NavigateBeforeIcon,
  NavigateNext as NavigateNextIcon,
} from "@mui/icons-material";
import SuperDataGrid from "@simplishelf/super-data-grid";
import type {
  SuperDataGridCellProps,
  SuperDataGridColumnConfiguration,
  SuperDataGridColumnOptions,
  SuperDataGridColumnType,
} from "@simplishelf/super-data-grid";
import {
  detectHeaderRow,
  downloadTemplate,
  getSourceColumns,
  getTemplateFields,
  normalizeHeader,
  readSpreadsheetFile,
  suggestColumnMappings,
} from "./workbook";
import type {
  SheetIngestField,
  SheetIngestIssue,
  SheetIngestProps,
  SheetIngestRow,
  SheetIngestRowContext,
  SheetIngestStepContext,
  SheetIngestTranslations,
} from "./types";

const DEFAULT_TRANSLATIONS: SheetIngestTranslations = {
  title: "Import data",
  steps: {
    upload: "Upload file",
    header: "Select header rows",
    mapping: "Match columns",
    preview: "Data preview",
  },
  downloadTemplate: "Download sample file",
  generateTestData: "Generate test data",
  uploadFile: "Upload spreadsheet",
  chooseFile: "Choose file",
  selectHeaderRow: "Header rows",
  matchColumns: "Match columns",
  dataPreview: "Data preview",
  back: "Back",
  next: "Continue",
  submit: "Import data",
  cancel: "Cancel",
  noMapping: "Select a spreadsheet column",
  ignoreColumn: "Do not import",
  required: "Required",
  valid: "Valid",
  invalid: "Needs attention",
  row: "Row",
  status: "Status",
  errorTitle: "Import could not continue",
  close: "Close",
  records: "records",
  testRecordCount: "Number of records",
  fileHelp: "Supported formats: CSV and XLSX. The first worksheet is used.",
  headerHelp: "We preselected the best matching row and its parent row when available. Select every row that belongs to your nested headers.",
  mappingHelp: "Match each template field to the column that contains its data.",
  previewHelp: "Review validation results before importing. Rows with errors are blocked by default.",
};

type WizardStep = "upload" | "header" | "mapping" | "preview";

interface ProcessedRow<Row extends SheetIngestRow> {
  rowNumber: number;
  data: Row;
  errors: SheetIngestIssue[];
}

interface KeyedRow {
  cells: string[];
  index: number;
  key: string;
}

interface HeaderPreviewGridRow {
  id: number;
  headerIndex: number;
  cells: string[];
}

type GridRow<Row extends SheetIngestRow> = Row & {
  rowNumber: number;
  status: string;
  __sheetIngestErrors: SheetIngestIssue[];
  __sheetIngestData: Row;
};

const STEPS: WizardStep[] = ["upload", "header", "mapping", "preview"];
const ACCEPTED_EXTENSIONS = new Set(["csv", "xlsx"]);
const DEFAULT_MAX_RECORDS = 1000;
const DEFAULT_MAX_FILE_SIZE = 10 * 1024 * 1024;
const DEFAULT_MAX_TEST_DATA_COUNT = 10_000;

const getFileExtension = (fileName: string): string =>
  fileName.split(".").pop()?.toLowerCase() ?? "";

const isBlankRow = (row: readonly string[]): boolean =>
  row.every((cell) => cell.trim() === "");

const createKeyedRows = (rows: readonly string[][]): KeyedRow[] => {
  const occurrences = new Map<string, number>();
  return rows.map((cells, index) => {
    const fingerprint = JSON.stringify(cells);
    const occurrence = occurrences.get(fingerprint) ?? 0;
    occurrences.set(fingerprint, occurrence + 1);
    return { cells, index, key: `${fingerprint}-${occurrence}` };
  });
};

const asIssue = (value: string | SheetIngestIssue): SheetIngestIssue =>
  typeof value === "string" ? { message: value } : value;

const toIssueList = (
  value: void | string | SheetIngestIssue | readonly (string | SheetIngestIssue)[],
): SheetIngestIssue[] => {
  if (value == null) return [];
  return (Array.isArray(value) ? value : [value]).map(asIssue);
};

const groupId = (label: string): string =>
  `sheet-ingest-${normalizeHeader(label) || "group"}`;

const createColumnGroups = (
  fields: readonly SheetIngestField[],
  groups: SheetIngestProps["template"]["groups"],
): GridColumnGroupingModel =>
  groups
    .map((group) => ({
      groupId: groupId(group.label),
      headerName: group.label,
      children: group.fields
        .filter((field) => !field.hideInPreview && fields.some((item) => item.key === field.key))
        .map((field) => ({ field: field.key })),
    }))
    .filter((group) => group.children.length > 0);

const cellText = (value: unknown): string => {
  if (value == null) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  if (value instanceof Date) return value.toLocaleString();
  try {
    return JSON.stringify(value) ?? "";
  } catch {
    return "";
  }
};

interface ProcessRecordOptions<Row extends SheetIngestRow> {
  sourceRow: string[];
  rowIndex: number;
  fields: readonly SheetIngestField[];
  columnMapping: Readonly<Record<string, number | null>>;
  matrix: string[][];
  headerRow: number;
  file: File;
  sheetName: string;
  mapRow: SheetIngestProps<Row>["mapRow"];
  validateRow: SheetIngestProps<Row>["validateRow"];
}

const processRecord = async <Row extends SheetIngestRow>({
  sourceRow,
  rowIndex,
  fields,
  columnMapping,
  matrix,
  headerRow,
  file,
  sheetName,
  mapRow,
  validateRow,
}: ProcessRecordOptions<Row>): Promise<ProcessedRow<Row>> => {
  const rawValues = Object.fromEntries(
    fields.map((field) => {
      const sourceIndex = columnMapping[field.key];
      return [
        field.key,
        sourceIndex == null || sourceIndex < 0 ? "" : sourceRow[sourceIndex] ?? "",
      ];
    }),
  ) as Record<string, string>;
  const physicalRowNumber = matrix.indexOf(sourceRow, headerRow + 1) + 1;
  const context: SheetIngestRowContext = {
    rowNumber: physicalRowNumber || headerRow + rowIndex + 2,
    file,
    sheetName,
    values: rawValues,
  };
  const errors: SheetIngestIssue[] = fields.flatMap((field) =>
    field.required && rawValues[field.key].trim() === ""
      ? [{ field: field.label, message: `${field.label} is required.` }]
      : [],
  );
  let data = rawValues as unknown as Row;

  try {
    if (mapRow) data = await mapRow(rawValues, context);
    if (validateRow) errors.push(...toIssueList(await validateRow(data, context)));
  } catch (cause) {
    errors.push({
      message: cause instanceof Error ? cause.message : "This row could not be processed.",
    });
  }

  return { rowNumber: context.rowNumber, data, errors };
};

const ErrorStatusCell = <Row extends SheetIngestRow>({
  row,
}: SuperDataGridCellProps<GridRow<Row>>): React.JSX.Element => {
  const issues = row.__sheetIngestErrors;
  if (issues.length === 0) {
    return <Chip icon={<CheckCircleOutlineIcon />} size="small" color="success" label="Valid" />;
  }
  const tooltip = (
    <Stack spacing={0.5} sx={{ maxWidth: 420, py: 0.5 }}>
      {issues.map((issue, index) => (
        <Typography key={`${issue.field ?? "row"}-${index}`} variant="caption">
          {issue.field ? `${issue.field}: ` : ""}{issue.message}
        </Typography>
      ))}
    </Stack>
  );
  return (
    <Tooltip arrow placement="top" title={tooltip}>
      <Chip icon={<ErrorOutlineIcon />} size="small" color="error" label={`${issues.length} issue${issues.length === 1 ? "" : "s"}`} />
    </Tooltip>
  );
};

const getCustomColumns = (
  fields: readonly SheetIngestField[],
  rowNumberLabel: string,
  statusLabel: string,
): {
  columns: string[];
  columnConfiguration: Partial<Record<string, SuperDataGridColumnConfiguration>>;
  columnOptions: Partial<Record<string, SuperDataGridColumnOptions>>;
  columnTypes: Partial<Record<string, SuperDataGridColumnType>>;
} => {
  const visibleFields = fields.filter((field) => !field.hideInPreview);
  const columnConfiguration: Partial<Record<string, SuperDataGridColumnConfiguration>> = {
    rowNumber: { headerName: rowNumberLabel, width: 76, minWidth: 70, maxWidth: 90 },
    status: { headerName: statusLabel, width: 150, minWidth: 130 },
  };
  const columnOptions: Partial<Record<string, SuperDataGridColumnOptions>> = {};
  const columnTypes: Partial<Record<string, SuperDataGridColumnType>> = {};
  for (const field of visibleFields) {
    columnConfiguration[field.key] = {
      headerName: field.label,
      width: field.width ?? 160,
      minWidth: Math.min(field.width ?? 160, 110),
    };
    const normalizedNames = [field.key, field.label, field.templateHeader, ...(field.aliases ?? [])]
      .filter((name): name is string => Boolean(name?.trim()))
      .map(normalizeHeader);
    const inferredType: SuperDataGridColumnType | undefined = field.type ??
      (normalizedNames.some((name) => name.includes("email"))
        ? "email"
        : normalizedNames.some((name) => /phone|telephone|mobile|cell|fax/.test(name))
          ? "phone"
          : undefined);

    if (inferredType) columnTypes[field.key] = inferredType;
    if (inferredType === "email") {
      columnOptions[field.key] = {
        ...field.columnOptions,
        email: { showIcon: true, ...field.columnOptions?.email },
      };
    } else if (inferredType === "phone") {
      columnOptions[field.key] = {
        ...field.columnOptions,
        phone: {
          showIcon: true,
          showFlag: true,
          countryCode: "US",
          format: "international",
          ...field.columnOptions?.phone,
        },
      };
    } else if (field.columnOptions) {
      columnOptions[field.key] = field.columnOptions;
    }
  }
  return {
    columns: ["rowNumber", ...visibleFields.map((field) => field.key), "status"],
    columnConfiguration,
    columnOptions,
    columnTypes,
  };
};

export const SheetIngest = <Row extends SheetIngestRow = SheetIngestRow>({
  isOpen,
  onClose,
  template,
  onSubmit,
  mapRow,
  validateRow,
  onGenerateTestData,
  renderStepContent,
  maxRecords = DEFAULT_MAX_RECORDS,
  maxFileSize = DEFAULT_MAX_FILE_SIZE,
  defaultTestDataCount = 50,
  maxTestDataCount = DEFAULT_MAX_TEST_DATA_COUNT,
  allowInvalidSubmit = false,
  title,
  dialogProps,
  translations,
}: SheetIngestProps<Row>): React.JSX.Element => {
  const t = useMemo<SheetIngestTranslations>(
    () => ({ ...DEFAULT_TRANSLATIONS, ...translations, steps: { ...DEFAULT_TRANSLATIONS.steps, ...translations?.steps } }),
    [translations],
  );
  const fields = useMemo(() => getTemplateFields(template), [template]);
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [step, setStep] = useState<WizardStep>("upload");
  const [file, setFile] = useState<File | null>(null);
  const [sheetName, setSheetName] = useState("");
  const [matrix, setMatrix] = useState<string[][]>([]);
  const [headerRow, setHeaderRow] = useState(0);
  const [headerRows, setHeaderRows] = useState<number[]>([0]);
  const [columnMapping, setColumnMapping] = useState<Record<string, number | null>>({});
  const [processedRows, setProcessedRows] = useState<ProcessedRow<Row>[]>([]);
  const [testDataCount, setTestDataCount] = useState(defaultTestDataCount);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);

  const reset = useCallback((): void => {
    setStep("upload");
    setFile(null);
    setSheetName("");
    setMatrix([]);
    setHeaderRow(0);
    setHeaderRows([0]);
    setColumnMapping({});
    setProcessedRows([]);
    setBusy(false);
    setError(null);
    setSuccess(null);
  }, []);

  useEffect(() => {
    if (isOpen) reset();
  }, [isOpen, reset]);

  const columns = useMemo(() => getSourceColumns(matrix, headerRows), [headerRows, matrix]);
  const keyedRows = useMemo(() => createKeyedRows(matrix), [matrix]);
  const headerPreviewColumnCount = useMemo(
    () => matrix.reduce((count, row) => Math.max(count, row.length), 0),
    [matrix],
  );
  const headerPreviewRows = useMemo<HeaderPreviewGridRow[]>(
    () => keyedRows.map(({ cells, index }) => ({ id: index, headerIndex: index, cells })),
    [keyedRows],
  );
  const headers = useMemo(
    () => columns.map((column) => column.header || `Column ${column.index + 1}`),
    [columns],
  );
  const records = useMemo(
    () => matrix.slice(headerRow + 1).filter((row) => !isBlankRow(row)),
    [headerRow, matrix],
  );
  const rowCount = processedRows.length;
  const errorRowCount = useMemo(
    () => processedRows.filter((row) => row.errors.some((issue) => issue.severity !== "warning" && issue.severity !== "info")).length,
    [processedRows],
  );
  const issueCount = useMemo(
    () => processedRows.reduce((count, row) => count + row.errors.length, 0),
    [processedRows],
  );
  const gridRows = useMemo<GridRow<Row>[]>(
    () => processedRows.map((item) => ({
      ...item.data,
      rowNumber: item.rowNumber,
      status: item.errors.length === 0 ? t.valid : t.invalid,
      __sheetIngestErrors: item.errors,
      __sheetIngestData: item.data,
    })),
    [processedRows, t.invalid, t.valid],
  );
  const grid = useMemo(
    () => getCustomColumns(fields, t.row, t.status),
    [fields, t.row, t.status],
  );
  const groupingModel = useMemo(
    () => createColumnGroups(fields, template.groups),
    [fields, template.groups],
  );
  const cellComponents = useMemo(
    () => ({ status: ErrorStatusCell<Row> }),
    [],
  );

  const contextForStep: SheetIngestStepContext = {
    step,
    file,
    rowCount: step === "preview" ? rowCount : records.length,
  };

  const handleClose = useCallback((): void => {
    reset();
    onClose();
  }, [onClose, reset]);

  const handleFile = useCallback(async (selectedFile: File): Promise<void> => {
    setError(null);
    setSuccess(null);
    setProcessedRows([]);
    const extension = getFileExtension(selectedFile.name);
    if (!ACCEPTED_EXTENSIONS.has(extension)) {
      setError("Choose a CSV or XLSX file.");
      return;
    }
    if (selectedFile.size > maxFileSize) {
      setError(`The file exceeds the ${Math.ceil(maxFileSize / (1024 * 1024))} MiB upload limit.`);
      return;
    }

    setBusy(true);
    try {
      const sheet = await readSpreadsheetFile(selectedFile);
      if (sheet.rows.length < 2) {
        throw new Error("The worksheet needs a header row and at least one data row.");
      }
      const detectedHeader = detectHeaderRow(sheet.rows, fields);
      const parentHeader = sheet.rows[detectedHeader - 1];
      const initialHeaderRows = detectedHeader > 0 && parentHeader?.some((cell) => cell.trim())
        ? [detectedHeader - 1, detectedHeader]
        : [detectedHeader];
      setFile(selectedFile);
      setSheetName(sheet.name);
      setMatrix(sheet.rows);
      setHeaderRow(detectedHeader);
      setHeaderRows(initialHeaderRows);
      setColumnMapping({});
      setStep("header");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "The file could not be read.");
    } finally {
      setBusy(false);
    }
  }, [fields, maxFileSize]);

  const handleFileInput = useCallback((event: React.ChangeEvent<HTMLInputElement>): void => {
    const selectedFile = event.currentTarget.files?.[0];
    if (selectedFile) void handleFile(selectedFile);
    event.currentTarget.value = "";
  }, [handleFile]);

  const handleDownloadTemplate = useCallback(async (): Promise<void> => {
    try {
      await downloadTemplate(template);
      setError(null);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "The sample file could not be created.");
    }
  }, [template]);

  const handleGenerateTestData = useCallback(async (): Promise<void> => {
    if (!onGenerateTestData) return;
    if (!Number.isInteger(testDataCount) || testDataCount < 1 || testDataCount > maxTestDataCount) {
      setError(`Enter a whole number from 1 to ${maxTestDataCount.toLocaleString()}.`);
      return;
    }
    setBusy(true);
    setError(null);
    setSuccess(null);
    try {
      await onGenerateTestData(testDataCount);
      setSuccess(`Generated ${testDataCount.toLocaleString()} test ${t.records}.`);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Test data could not be generated.");
    } finally {
      setBusy(false);
    }
  }, [maxTestDataCount, onGenerateTestData, t.records, testDataCount]);

  const applyHeaderRows = useCallback((rows: readonly number[]): void => {
    const nextHeaderRows = [...new Set(rows)]
      .filter((index) => Number.isInteger(index) && index >= 0 && index < matrix.length)
      .sort((left, right) => left - right);
    if (nextHeaderRows.length === 0) return;
    setHeaderRows(nextHeaderRows);
    setHeaderRow(nextHeaderRows[nextHeaderRows.length - 1]);
    setColumnMapping({});
  }, [matrix.length]);

  const handleHeaderRowsChange = useCallback((rowIndex: number): void => {
    const nextHeaderRows = headerRows.includes(rowIndex)
      ? headerRows.length > 1 ? headerRows.filter((index) => index !== rowIndex) : headerRows
      : [...headerRows, rowIndex].sort((left, right) => left - right);
    applyHeaderRows(nextHeaderRows);
  }, [applyHeaderRows, headerRows]);

  const headerPreviewColumns = useMemo<GridColDef<HeaderPreviewGridRow>[]>(() => [
    {
      field: "headerRow",
      headerName: t.selectHeaderRow,
      width: 112,
      sortable: false,
      filterable: false,
      disableColumnMenu: true,
      renderCell: ({ row }) => {
        const selected = headerRows.includes(row.headerIndex);
        return (
          <Stack direction="row" spacing={0.25} alignItems="center">
            <Checkbox
              size="small"
              checked={selected}
              onChange={() => handleHeaderRowsChange(row.headerIndex)}
              onClick={(event) => event.stopPropagation()}
              inputProps={{ "aria-label": `Select row ${row.headerIndex + 1} as a header row` }}
            />
            <Typography variant="body2" fontWeight={selected ? 700 : 400}>
              {row.headerIndex + 1}
            </Typography>
          </Stack>
        );
      },
    },
    ...Array.from({ length: headerPreviewColumnCount }, (_, cellIndex): GridColDef<HeaderPreviewGridRow> => ({
      field: `column${cellIndex}`,
      headerName: `Column ${cellIndex + 1}`,
      minWidth: 160,
      flex: 1,
      sortable: false,
      filterable: false,
      disableColumnMenu: true,
      renderCell: ({ row }) => (
        <Typography
          variant="body2"
          noWrap
          title={row.cells[cellIndex] ?? ""}
          fontWeight={headerRows.includes(row.headerIndex) ? 700 : 400}
        >
          {row.cells[cellIndex] ?? ""}
        </Typography>
      ),
    })),
  ], [handleHeaderRowsChange, headerPreviewColumnCount, headerRows, t.selectHeaderRow]);

  const handleContinueFromHeader = useCallback((): void => {
    if (columns.length === 0) {
      setError("The selected header row has no columns.");
      return;
    }
    setColumnMapping(suggestColumnMappings(headers, fields));
    setError(null);
    setStep("mapping");
  }, [columns.length, fields, headers]);

  const handleMappingChange = useCallback((fieldKey: string, sourceIndex: number | null): void => {
    setColumnMapping((current) => ({ ...current, [fieldKey]: sourceIndex }));
  }, []);

  const handleBuildPreview = useCallback(async (): Promise<void> => {
    if (!file) return;
    if (records.length === 0) {
      setError("No data rows were found after the selected header row.");
      return;
    }
    if (records.length > maxRecords) {
      setError(`The file has ${records.length.toLocaleString()} data rows. The maximum is ${maxRecords.toLocaleString()}.`);
      return;
    }

    setBusy(true);
    setError(null);
    try {
      const result = await Promise.all(
        records.map((sourceRow, rowIndex) =>
          processRecord({
            sourceRow,
            rowIndex,
            fields,
            columnMapping,
            matrix,
            headerRow,
            file,
            sheetName,
            mapRow,
            validateRow,
          }),
        ),
      );
      setProcessedRows(result);
      setStep("preview");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "The file could not be prepared for review.");
    } finally {
      setBusy(false);
    }
  }, [columnMapping, fields, file, headerRow, mapRow, maxRecords, matrix, records, sheetName, validateRow]);

  const handleSubmit = useCallback(async (): Promise<void> => {
    if (!file || processedRows.length === 0) return;
    if (!allowInvalidSubmit && errorRowCount > 0) {
      setError("Resolve the validation errors before importing this file.");
      return;
    }
    setBusy(true);
    setError(null);
    try {
      await onSubmit(processedRows.map((row) => row.data), file);
      handleClose();
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "The import could not be submitted.");
    } finally {
      setBusy(false);
    }
  }, [allowInvalidSubmit, errorRowCount, file, handleClose, onSubmit, processedRows]);

  const goBack = useCallback((): void => {
    setError(null);
    setStep((current) => {
      const currentIndex = STEPS.indexOf(current);
      return STEPS[Math.max(currentIndex - 1, 0)];
    });
  }, []);

  const goNext = useCallback((): void => {
    setError(null);
    if (step === "header") handleContinueFromHeader();
    if (step === "mapping") void handleBuildPreview();
  }, [handleBuildPreview, handleContinueFromHeader, step]);

  const handleDrop = useCallback((event: React.DragEvent<HTMLDivElement>): void => {
    event.preventDefault();
    const droppedFile = event.dataTransfer.files[0];
    if (droppedFile) void handleFile(droppedFile);
  }, [handleFile]);

  const activeStepIndex = STEPS.indexOf(step);
  const stepLabels: Record<WizardStep, string> = t.steps;
  const mappingComplete = fields.every((field) => {
    const mapping = columnMapping[field.key];
    return !field.required || (mapping != null && mapping >= 0);
  });
  const componentRows = rowCount;
  const previewRows = useMemo(
    () => gridRows.map((row) => ({
      ...row,
      ...Object.fromEntries(fields.map((field) => [field.key, cellText((row as Record<string, unknown>)[field.key])])),
    })),
    [fields, gridRows],
  );

  return (
    <Dialog
      {...dialogProps}
      open={isOpen}
      onClose={handleClose}
      fullWidth={dialogProps?.fullWidth ?? true}
      maxWidth={dialogProps?.maxWidth ?? "xl"}
      aria-labelledby="sheet-ingest-title"
    >
      <DialogTitle id="sheet-ingest-title" sx={{ pb: 1 }}>
        <Stack spacing={2}>
          <Typography variant="h6">{title ?? t.title}</Typography>
          <Stepper activeStep={activeStepIndex} alternativeLabel>
            {STEPS.map((item) => (
              <Step key={item}>
                <StepLabel>{stepLabels[item]}</StepLabel>
              </Step>
            ))}
          </Stepper>
        </Stack>
      </DialogTitle>
      <DialogContent dividers sx={{ minHeight: { xs: 360, md: 460 }, pt: 3 }}>
        {busy && <LinearProgress sx={{ mb: 2 }} />}
        {error && <Alert severity="error" sx={{ mb: 2 }}>{error}</Alert>}
        {success && <Alert severity="success" sx={{ mb: 2 }}>{success}</Alert>}

        {step === "upload" && (
          <Stack spacing={2}>
            <Typography color="text.secondary">{t.fileHelp}</Typography>
            <Box
              onDragOver={(event) => event.preventDefault()}
              onDrop={handleDrop}
              sx={{
                border: "1px dashed",
                borderColor: "divider",
                borderRadius: 2,
                bgcolor: "action.hover",
                p: { xs: 3, md: 5 },
                textAlign: "center",
              }}
            >
              <Stack spacing={1.5} alignItems="center">
                {file ? <FilePresentIcon color="primary" fontSize="large" /> : <CloudUploadIcon color="primary" fontSize="large" />}
                <Typography variant="subtitle1" fontWeight={600}>
                  {file?.name ?? t.uploadFile}
                </Typography>
                {file && <Typography variant="body2" color="text.secondary">{(file.size / 1024).toFixed(1)} KB</Typography>}
                <Button
                  variant="contained"
                  startIcon={<CloudUploadIcon />}
                  onClick={() => fileInputRef.current?.click()}
                  disabled={busy}
                >
                  {file ? t.chooseFile : t.uploadFile}
                </Button>
                <input
                  ref={fileInputRef}
                  type="file"
                  hidden
                  accept=".csv,.xlsx"
                  onChange={handleFileInput}
                  aria-label={t.uploadFile}
                />
                <Typography variant="caption" color="text.secondary">or drop a file here</Typography>
              </Stack>
            </Box>
            <Divider />
            <Stack direction={{ xs: "column", sm: "row" }} spacing={1.5} alignItems={{ xs: "stretch", sm: "center" }}>
              <Button variant="outlined" startIcon={<DownloadIcon />} onClick={() => void handleDownloadTemplate()}>
                {t.downloadTemplate}
              </Button>
              {onGenerateTestData && (
                <>
                  <TextField
                    size="small"
                    type="number"
                    label={t.testRecordCount}
                    value={testDataCount}
                    onChange={(event) => setTestDataCount(Number(event.target.value))}
                    inputProps={{ min: 1, max: maxTestDataCount, step: 1 }}
                    sx={{ width: { xs: "100%", sm: 190 } }}
                  />
                  <Button
                    variant="text"
                    startIcon={<AutoAwesomeIcon />}
                    onClick={() => void handleGenerateTestData()}
                    disabled={busy}
                  >
                    {t.generateTestData}
                  </Button>
                </>
              )}
            </Stack>
          </Stack>
        )}

        {step === "header" && (
          <Stack spacing={2}>
            <Stack direction={{ xs: "column", sm: "row" }} spacing={2} alignItems={{ sm: "center" }}>
              <Box sx={{ flex: 1 }}>
                <Typography variant="subtitle1" fontWeight={600}>{file?.name}</Typography>
                <Typography variant="body2" color="text.secondary">{t.headerHelp}</Typography>
              </Box>
              <FormControl size="small" sx={{ width: { xs: "100%", sm: 260 } }}>
                <InputLabel id="sheet-ingest-header-rows-label">{t.selectHeaderRow}</InputLabel>
                <Select
                  labelId="sheet-ingest-header-rows-label"
                  label={t.selectHeaderRow}
                  multiple
                  value={headerRows}
                  renderValue={(selected) => selected.map((index) => `Row ${index + 1}`).join(", ")}
                  onChange={(event) => {
                    const selected = event.target.value;
                    const nextHeaderRows = Array.isArray(selected)
                      ? selected.map(Number)
                      : String(selected).split(",").map(Number);
                    applyHeaderRows(nextHeaderRows);
                  }}
                >
                  {keyedRows.map(({ index, key }) => (
                    <MenuItem key={key} value={index}>
                      <Checkbox size="small" checked={headerRows.includes(index)} />
                      Row {index + 1}
                    </MenuItem>
                  ))}
                </Select>
              </FormControl>
            </Stack>
              <Box sx={{ width: "100%", height: 360 }}>
                <DataGrid
                  rows={headerPreviewRows}
                  columns={headerPreviewColumns}
                  getRowId={(row) => row.id}
                  getRowClassName={({ row }) =>
                    headerRows.includes(row.headerIndex) ? "sheet-ingest-selected-header" : ""
                  }
                  onRowClick={({ row }) => handleHeaderRowsChange(row.headerIndex)}
                  disableRowSelectionOnClick
                  disableColumnFilter
                  disableColumnSelector
                  disableDensitySelector
                  hideFooter
                  rowHeight={40}
                  columnHeaderHeight={40}
                  sx={{
                    borderRadius: 1,
                    "& .sheet-ingest-selected-header .MuiDataGrid-cell": {
                      fontWeight: 700,
                      backgroundColor: "action.selected",
                    },
                    "& .MuiDataGrid-row": { cursor: "pointer" },
                  }}
                />
              </Box>
            <Typography variant="body2" color="text.secondary">
              Worksheet: {sheetName} · {matrix.length.toLocaleString()} rows · {headerPreviewColumnCount.toLocaleString()} columns
            </Typography>
          </Stack>
        )}

        {step === "mapping" && (
          <Stack spacing={2}>
            <Box>
              <Typography variant="subtitle1" fontWeight={600}>{t.matchColumns}</Typography>
              <Typography variant="body2" color="text.secondary">{t.mappingHelp}</Typography>
            </Box>
            <Stack spacing={2}>
              {template.groups.map((group) => (
                <Paper key={group.label} variant="outlined" sx={{ overflow: "hidden" }}>
                  <Box sx={{ px: 2, py: 1, bgcolor: "action.hover" }}>
                    <Typography variant="subtitle2" fontWeight={700}>{group.label}</Typography>
                  </Box>
                  <Stack divider={<Divider flexItem />}>
                    {group.fields.map((field) => {
                      const sourceIndex = columnMapping[field.key] ?? null;
                      const mapped = columns.find((column) => column.index === sourceIndex);
                      const unavailableIndices = new Set(
                        Object.entries(columnMapping)
                          .filter(([mappedKey, index]) => mappedKey !== field.key && index != null)
                          .map(([, index]) => index as number),
                      );
                      return (
                        <Stack
                          key={field.key}
                          direction={{ xs: "column", md: "row" }}
                          spacing={1.5}
                          alignItems={{ md: "center" }}
                          sx={{ px: 2, py: 1.25 }}
                        >
                          <Box sx={{ flex: "1 1 30%", minWidth: 180 }}>
                            <Stack direction="row" spacing={0.75} alignItems="center">
                              <Typography variant="body2" fontWeight={600}>{field.label}</Typography>
                              {field.required && <Chip size="small" label={t.required} color="error" variant="outlined" />}
                            </Stack>
                            {field.example && <Typography variant="caption" color="text.secondary">Example: {field.example}</Typography>}
                          </Box>
                          <FormControl size="small" sx={{ flex: "1 1 38%", minWidth: 190 }}>
                            <InputLabel id={`sheet-ingest-map-${field.key}`}>{t.matchColumns}</InputLabel>
                            <Select
                              labelId={`sheet-ingest-map-${field.key}`}
                              label={t.matchColumns}
                              value={sourceIndex ?? ""}
                              onChange={(event) => {
                                const value = event.target.value === "" ? null : Number(event.target.value);
                                handleMappingChange(field.key, value);
                              }}
                            >
                              <MenuItem value=""><em>{t.noMapping}</em></MenuItem>
                              {columns.map((column) => (
                                <MenuItem
                                  key={column.index}
                                  value={column.index}
                                  disabled={unavailableIndices.has(column.index)}
                                >
                                  {column.group ? `${column.group} / ` : ""}{column.header || `Column ${column.index + 1}`}
                                </MenuItem>
                              ))}
                              <Divider />
                              <MenuItem value={-1}>{t.ignoreColumn}</MenuItem>
                            </Select>
                          </FormControl>
                          <Box sx={{ flex: "1 1 32%", minWidth: 160 }}>
                            <Typography variant="caption" color="text.secondary">Spreadsheet sample</Typography>
                            <Typography variant="body2" noWrap title={mapped?.sample ?? ""}>
                              {mapped?.sample || "—"}
                            </Typography>
                          </Box>
                        </Stack>
                      );
                    })}
                  </Stack>
                </Paper>
              ))}
            </Stack>
            {!mappingComplete && (
              <Alert severity="warning">One or more required template fields are not mapped. They will be marked as missing in the preview.</Alert>
            )}
          </Stack>
        )}

        {step === "preview" && (
          <Stack spacing={1.5}>
            <Stack direction={{ xs: "column", sm: "row" }} spacing={1} alignItems={{ sm: "center" }} justifyContent="space-between">
              <Box>
                <Typography variant="subtitle1" fontWeight={600}>{t.dataPreview}</Typography>
                <Typography variant="body2" color="text.secondary">{t.previewHelp}</Typography>
              </Box>
              <Stack direction="row" spacing={1}>
                <Chip size="small" color="success" icon={<CheckCircleOutlineIcon />} label={`${componentRows - errorRowCount} ${t.valid}`} />
                {errorRowCount > 0 && <Chip size="small" color="error" icon={<ErrorOutlineIcon />} label={`${errorRowCount} ${t.invalid}`} />}
                <Chip size="small" variant="outlined" label={`${componentRows} ${t.records}`} />
              </Stack>
            </Stack>
            {issueCount > 0 && (
              <Alert severity="warning">{issueCount.toLocaleString()} validation issue{issueCount === 1 ? "" : "s"} found. Hover over an issue status to review its details.</Alert>
            )}
            <Box sx={{ width: "100%", height: { xs: 360, md: 460 } }}>
              <SuperDataGrid
                columns={grid.columns}
                data={previewRows}
                columnConfiguration={grid.columnConfiguration}
                columnOptions={grid.columnOptions}
                columnTypes={grid.columnTypes}
                columnGroupingModel={groupingModel}
                cellComponents={cellComponents}
                getRowId={(row) => row.rowNumber}
                minHeight="100%"
                rowHeight={46}
                pageSizeOptions={[10, 25, 50, 100]}
                hideToolbar
                hideViews
                canAddViews={false}
              />
            </Box>
          </Stack>
        )}

        {renderStepContent?.(contextForStep)}
      </DialogContent>
      <DialogActions sx={{ px: 3, py: 2, justifyContent: "space-between" }}>
        <Button onClick={handleClose} disabled={busy}>{t.cancel}</Button>
        <Stack direction="row" spacing={1}>
          {activeStepIndex > 0 && (
            <Button startIcon={<NavigateBeforeIcon />} onClick={goBack} disabled={busy}>
              {t.back}
            </Button>
          )}
          {step === "header" && (
            <Button variant="contained" endIcon={<NavigateNextIcon />} onClick={goNext} disabled={busy || !file}>
              {t.next}
            </Button>
          )}
          {step === "mapping" && (
            <Button variant="contained" endIcon={<NavigateNextIcon />} onClick={goNext} disabled={busy || records.length === 0}>
              {busy ? "Preparing…" : t.next}
            </Button>
          )}
          {step === "preview" && (
            <Button variant="contained" onClick={() => void handleSubmit()} disabled={busy || rowCount === 0 || (!allowInvalidSubmit && errorRowCount > 0)}>
              {busy ? "Importing…" : `${t.submit} (${rowCount})`}
            </Button>
          )}
        </Stack>
      </DialogActions>
    </Dialog>
  );
};

export default SheetIngest;
