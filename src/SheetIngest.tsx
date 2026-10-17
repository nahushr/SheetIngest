import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type React from "react";
import type { GridPaginationModel } from "@mui/x-data-grid";
import { OpsModalDialog } from "@simplishelf/opscards";
import {
  Alert,
  Avatar,
  Box,
  Button,
  Chip,
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
import {
  AutoAwesome as AutoAwesomeIcon,
  Check as CheckIcon,
  CheckCircleOutline as CheckCircleOutlineIcon,
  CalendarMonth as CalendarMonthIcon,
  CloudUpload as CloudUploadIcon,
  Download as DownloadIcon,
  ErrorOutline as ErrorOutlineIcon,
  FilePresent as FilePresentIcon,
  NavigateBefore as NavigateBeforeIcon,
  NavigateNext as NavigateNextIcon,
} from "@mui/icons-material";
import SuperDataGrid, { DateCell, UserContactCell } from "@simplishelf/super-data-grid";
import type {
  SuperDataGridCellComponent,
  SuperDataGridCellProps,
  SuperDataGridColumnConfiguration,
  SuperDataGridColumnOptions,
  SuperDataGridColumnType,
} from "@simplishelf/super-data-grid";
import {
  detectHeaderRow,
  createTemplateBlob,
  downloadBlob,
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
import type { SourceColumn } from "./workbook";
import { SheetIngestIdPopover } from "./SheetIngestIdPopover";
import { IMPORT_GRID_DEFAULT_PAGE_SIZE, IMPORT_GRID_PAGE_SIZE_OPTIONS } from "./pagination";
import styles from "./SheetIngest.module.css";

const selectMenuProps = {
  classes: { root: styles.menuRoot },
  PaperProps: { className: styles.menuPaper },
  MenuListProps: { className: styles.menuList },
};

const DEFAULT_TRANSLATIONS: SheetIngestTranslations = {
  title: "Import data",
  steps: {
    reference: "Reference data",
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
  referenceHelp: "Use these current values and IDs when preparing the spreadsheet. This reference data is read-only; filtering, columns, density, and export are available.",
};

type WizardStep = "reference" | "upload" | "mapping" | "preview";

const STEP_HELP_KEYS: Record<WizardStep, keyof Pick<SheetIngestTranslations, "referenceHelp" | "fileHelp" | "mappingHelp" | "previewHelp">> = {
  reference: "referenceHelp",
  upload: "fileHelp",
  mapping: "mappingHelp",
  preview: "previewHelp",
};

interface ProcessedRow<Row extends SheetIngestRow> {
  rowNumber: number;
  data: Row;
  errors: SheetIngestIssue[];
  sourceValues: Array<{ header: string; value: string }>;
}

type GridRow<Row extends SheetIngestRow> = Row & {
  rowNumber: number;
  status: string;
  __sheetIngestErrors: SheetIngestIssue[];
  __sheetIngestData: Row;
  __sheetIngestSourceValues: Array<{ header: string; value: string }>;
};

const UPLOAD_STEPS: WizardStep[] = ["upload", "mapping", "preview"];
const ACCEPTED_EXTENSIONS = new Set(["csv", "xlsx"]);
const DEFAULT_MAX_RECORDS = 1000;
const DEFAULT_MAX_FILE_SIZE = 10 * 1024 * 1024;
const DEFAULT_MAX_TEST_DATA_COUNT = 10_000;
const PREVIEW_USER_DETAILS_FIELD = "__sheetIngestPreviewUserDetails";
const PREVIEW_ADDRESSES_FIELD = "__sheetIngestPreviewAddresses";
const PREVIEW_RAW_DATA_FIELD = "__sheetIngestPreviewRawData";

const USER_DETAIL_FIELDS = new Set([
  "firstname",
  "lastname",
  "name",
  "fullname",
  "loginname",
  "email",
  "phone",
  "imageurl",
  "profilepicture",
]);

const ADDRESS_PART_FIELDS = new Set([
  "streetaddress",
  "streetaddress2",
  "streetaddress3",
  "city",
  "state",
  "zipcode",
  "postalcode",
  "country",
  "addresstype",
  "nameonaddress",
  "emailonaddress",
  "phoneonaddress",
]);

const getFileExtension = (fileName: string): string =>
  fileName.split(".").pop()?.toLowerCase() ?? "";

const isBlankRow = (row: readonly string[]): boolean =>
  row.every((cell) => cell.trim() === "");

const asIssue = (value: string | SheetIngestIssue): SheetIngestIssue =>
  typeof value === "string" ? { message: value } : value;

const toIssueList = (
  value: void | string | SheetIngestIssue | readonly (string | SheetIngestIssue)[],
): SheetIngestIssue[] => {
  if (value == null) return [];
  return (Array.isArray(value) ? value : [value]).map(asIssue);
};

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
  sourceColumns: readonly SourceColumn[];
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
  sourceColumns,
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

  const sourceValues = sourceColumns.flatMap((column) => {
    const value = sourceRow[column.index] ?? "";
    if (value.trim() === "") return [];
    const baseHeader = column.header || `Column ${column.index + 1}`;
    const header = column.group ? `${column.group} · ${baseHeader}` : baseHeader;
    return [{ header, value }];
  });

  return { rowNumber: context.rowNumber, data, errors, sourceValues };
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

interface PreviewAddressFieldGroup {
  number: number;
  fields: SheetIngestField[];
}

interface PreviewAddressValue {
  label: string;
  addressType: string;
  summary: string;
  lines: string[];
  contact?: string;
}

interface PreviewSourceValue {
  header: string;
  value: string;
}

const getAddressFieldGroup = (
  field: SheetIngestField,
): { number: number; part: string } | null => {
  const normalizedKey = normalizeHeader(field.key);
  if (ADDRESS_PART_FIELDS.has(normalizedKey)) {
    return { number: 1, part: normalizedKey };
  }

  const match = /^address([1-9])(.+)$/.exec(normalizedKey);
  if (!match || !ADDRESS_PART_FIELDS.has(match[2] ?? "")) return null;
  return { number: Number(match[1]), part: match[2] ?? "" };
};

const getFieldNames = (field: SheetIngestField): string[] =>
  [field.key, field.label, field.templateHeader, ...(field.aliases ?? [])]
    .filter((name): name is string => Boolean(name?.trim()))
    .map(normalizeHeader);

const getPreviewFieldType = (
  field: SheetIngestField,
): SuperDataGridColumnType | undefined => {
  if (field.type) return field.type;
  const names = getFieldNames(field);
  if (names.some((name) => /role$/.test(name))) return "badge";
  if (names.some((name) => /datetime|dateandtime|timestamp/.test(name))) return "dateTime";
  if (names.some((name) => /^(dob|birthdate|dateofbirth)$/.test(name) || /date$/.test(name))) {
    return "date";
  }
  if (names.some((name) => name.includes("email"))) return "email";
  if (names.some((name) => /phone|telephone|mobile|cell|fax/.test(name))) return "phone";
  return undefined;
};

const isUrlField = (field: SheetIngestField): boolean =>
  getFieldNames(field).some((name) => /url|website|href|hyperlink/.test(name));

const getRecordText = (record: object, key: string): string => {
  const value = Reflect.get(record, key);
  return typeof value === "string" || typeof value === "number"
    ? String(value).trim()
    : value instanceof Date && !Number.isNaN(value.getTime())
      ? value.toLocaleDateString()
      : "";
};

const getSafeHttpUrl = (value: string): string | null => {
  try {
    const parsed = new URL(value);
    return parsed.protocol === "http:" || parsed.protocol === "https:" ? parsed.href : null;
  } catch {
    return null;
  }
};

const getListItems = (value: unknown): string[] => {
  const values = Array.isArray(value)
    ? value.map(cellText)
    : typeof value === "string"
      ? value.split(/[;,|]/)
      : [cellText(value)];
  return values.map((item) => item.trim()).filter(Boolean);
};

const isListField = (field: SheetIngestField): boolean =>
  getFieldNames(field).some((name) => /(?:ids|idlist|list)$/.test(name));

const isPermissionIdField = (field: SheetIngestField): boolean =>
  getFieldNames(field).some((name) => /permissionids?$/.test(name));

const isGroupIdField = (field: SheetIngestField): boolean =>
  getFieldNames(field).some((name) => /(?:group|usergroup)ids?$/.test(name));

const getListFieldLabel = (field: SheetIngestField): string => {
  if (isGroupIdField(field)) return "group";
  const normalized = normalizeHeader(field.key).replace(/ids?$/, "");
  if (normalized.includes("user") || normalized.includes("member")) return "user";
  if (normalized.includes("label")) return "label";
  if (normalized.includes("permission")) return "permission";
  const label = field.label.replace(/\s*\(.*\)\s*$/, "").trim();
  return label || "item";
};

const isPreviewAddressValue = (value: unknown): value is PreviewAddressValue => {
  if (typeof value !== "object" || value === null || Array.isArray(value)) return false;
  const record = value as Record<string, unknown>;
  return typeof record.label === "string" && typeof record.addressType === "string" &&
    typeof record.summary === "string" && Array.isArray(record.lines) &&
    record.lines.every((line) => typeof line === "string");
};

const PreviewUserDetailsCell = <Row extends SheetIngestRow>({
  value,
}: SuperDataGridCellProps<GridRow<Row>>): React.JSX.Element => {
  const details = typeof value === "object" && value !== null && !Array.isArray(value)
    ? value as Record<string, unknown>
    : {};
  const firstName = cellText(details.firstName).trim();
  const lastName = cellText(details.lastName).trim();
  const name = cellText(details.name || details.fullName).trim() || [firstName, lastName].filter(Boolean).join(" ");
  const email = cellText(details.email || details.loginName).trim();
  const phone = cellText(details.phone).trim();
  const photoUrl = getSafeHttpUrl(cellText(details.imageUrl || details.profilePicture).trim());
  const initials = [firstName[0], lastName[0]].filter(Boolean).join("").toUpperCase() || name.slice(0, 1).toUpperCase();
  const addresses = Array.isArray(details.addresses) ? details.addresses.filter(isPreviewAddressValue) : [];
  const avatar = (
    <Avatar className={styles.previewUserAvatar} src={photoUrl ?? undefined} alt={name || "User profile"}>
      {initials || "?"}
    </Avatar>
  );

  return (
    <div className={styles.previewUserDetails}>
      <div className={styles.previewUserMain}>
        {photoUrl ? (
          <a
            aria-label={`Open profile image for ${name || email || "user"}`}
            className={styles.previewAvatarLink}
            href={photoUrl}
            rel="noreferrer"
            target="_blank"
            title="Open profile image"
          >
            {avatar}
          </a>
        ) : avatar}
        <UserContactCell
          email={email}
          firstName={firstName}
          lastName={lastName}
          name={name}
          phone={phone}
          showNameIcon={false}
          wrap
        />
      </div>
      {addresses.length > 0 && (
        <div className={styles.previewUserAddresses}>
          <span className={styles.previewAddressSectionLabel}>Addresses · {addresses.length}</span>
          <div className={styles.previewAddressChips}>
            {addresses.map((address) => (
              <details className={styles.previewAddressDisclosure} key={address.label}>
                <summary className={styles.previewAddressChip} title={`${address.addressType || "Address"}: ${address.summary}`}>
                  <span className={styles.previewAddressChipType}>{address.addressType || address.label}</span>
                  <span className={styles.previewAddressSummary}>{address.summary}</span>
                </summary>
                <div className={styles.previewAddressPopover}>
                  <div className={styles.previewAddressPopoverHeading}>
                    <strong>{address.label}</strong>
                    {address.addressType && <span className={styles.previewAddressType}>{address.addressType}</span>}
                  </div>
                  {address.lines.map((line, index) => (
                    <span className={styles.previewAddressLine} key={`${address.label}-${index}`}>{line}</span>
                  ))}
                  {address.contact && <span className={styles.previewAddressContact}>{address.contact}</span>}
                </div>
              </details>
            ))}
          </div>
        </div>
      )}
    </div>
  );
};

const PreviewAddressesCell = <Row extends SheetIngestRow>({
  value,
}: SuperDataGridCellProps<GridRow<Row>>): React.JSX.Element => {
  const addresses = Array.isArray(value) ? value.filter(isPreviewAddressValue) : [];
  if (addresses.length === 0) return <span className={styles.previewEmptyValue}>—</span>;

  return (
    <div className={styles.previewAddressList}>
      {addresses.map((address) => (
        <div className={styles.previewAddress} key={address.label}>
          <div className={styles.previewAddressHeading}>
            <span className={styles.previewAddressLabel}>{address.label}</span>
            {address.addressType && <span className={styles.previewAddressType}>{address.addressType}</span>}
          </div>
          {address.lines.map((line, index) => (
            <span className={styles.previewAddressLine} key={`${address.label}-${index}`}>{line}</span>
          ))}
          {address.contact && <span className={styles.previewAddressContact}>{address.contact}</span>}
        </div>
      ))}
    </div>
  );
};

const PreviewListCountCell = <Row extends SheetIngestRow>({
  field,
  value,
}: SuperDataGridCellProps<GridRow<Row>>): React.JSX.Element => {
  const items = getListItems(value);
  if (items.length === 0) return <span className={styles.previewEmptyValue}>—</span>;
  const label = getListFieldLabel({ key: field, label: field });
  const chipLabel = `${label.charAt(0).toUpperCase()}${label.slice(1)}${label.endsWith("s") ? "" : "s"}`;

  return (
    <div className={styles.previewIdList}>
      <SheetIngestIdPopover label={chipLabel} values={items} />
    </div>
  );
};

const PreviewRawDataCell = <Row extends SheetIngestRow>({
  value,
}: SuperDataGridCellProps<GridRow<Row>>): React.JSX.Element => {
  const values = Array.isArray(value)
    ? value.filter((item): item is PreviewSourceValue =>
      typeof item === "object" && item !== null &&
      typeof (item as PreviewSourceValue).header === "string" &&
      typeof (item as PreviewSourceValue).value === "string")
    : [];
  if (values.length === 0) return <span className={styles.previewEmptyValue}>No entered source values</span>;

  return (
    <details className={styles.previewRawDataDisclosure}>
      <summary>All spreadsheet values ({values.length})</summary>
      <dl className={styles.previewRawDataList}>
        {values.map((item, index) => (
          <div className={styles.previewRawDataItem} key={`${item.header}-${index}`}>
            <dt>{item.header}</dt>
            <dd>{item.value}</dd>
          </div>
        ))}
      </dl>
    </details>
  );
};

const PreviewUrlCell = <Row extends SheetIngestRow>({
  value,
}: SuperDataGridCellProps<GridRow<Row>>): React.JSX.Element => {
  const text = cellText(value).trim();
  const href = getSafeHttpUrl(text);
  if (!text) return <span className={styles.previewEmptyValue}>—</span>;
  return href ? (
    <a className={styles.previewUrlLink} href={href} rel="noreferrer" target="_blank" title={text}>
      {text}
    </a>
  ) : (
    <span className={styles.previewWrappedValue}>{text}</span>
  );
};

const PreviewDateCell = <Row extends SheetIngestRow>({
  value,
}: SuperDataGridCellProps<GridRow<Row>>): React.JSX.Element => (
  value == null || value === "" ? <span className={styles.previewEmptyValue}>—</span> : (
    <span className={styles.previewDate}>
      <CalendarMonthIcon aria-hidden="true" />
      <DateCell emptyText="—" formatString="do MMM yyyy" options={{ showIcon: false, timeZone: "UTC" }} value={value} />
    </span>
  )
);

const getCustomColumns = (
  fields: readonly SheetIngestField[],
  rowNumberLabel: string,
  statusLabel: string,
): {
  columns: string[];
  columnConfiguration: Partial<Record<string, SuperDataGridColumnConfiguration>>;
  columnOptions: Partial<Record<string, SuperDataGridColumnOptions>>;
  columnTypes: Partial<Record<string, SuperDataGridColumnType>>;
  hasUserDetails: boolean;
  userDetailsFields: SheetIngestField[];
  addressGroups: PreviewAddressFieldGroup[];
  urlFields: string[];
  dateFields: string[];
  listFields: SheetIngestField[];
} => {
  const visibleFields = fields.filter((field) => !field.hideInPreview);
  const nameKeys = new Set(visibleFields.flatMap(getFieldNames));
  const hasName = (nameKeys.has("firstname") && nameKeys.has("lastname")) ||
    nameKeys.has("name") || nameKeys.has("fullname");
  const hasEmail = nameKeys.has("loginname") || nameKeys.has("email");
  const userDetailsFields = hasName && hasEmail
    ? visibleFields.filter((field) => getFieldNames(field).some((name) => USER_DETAIL_FIELDS.has(name)))
    : [];
  const hasUserDetails = userDetailsFields.length > 1;
  const addressGroupsByNumber = new Map<number, SheetIngestField[]>();
  for (const field of visibleFields) {
    const group = getAddressFieldGroup(field);
    if (!group) continue;
    const addressFields = addressGroupsByNumber.get(group.number) ?? [];
    addressFields.push(field);
    addressGroupsByNumber.set(group.number, addressFields);
  }
  const addressGroups = [...addressGroupsByNumber.entries()]
    .sort(([left], [right]) => left - right)
    .map(([number, addressFields]) => ({ number, fields: addressFields }));
  const hasRole = visibleFields.some((field) => getPreviewFieldType(field) === "badge");
  const permissionFields = hasRole ? visibleFields.filter(isPermissionIdField) : [];
  const permissionFieldKeys = new Set(permissionFields.map((field) => field.key));
  const listFields = visibleFields.filter((field) =>
    isListField(field) && !permissionFieldKeys.has(field.key),
  );
  const listFieldKeys = new Set(listFields.map((field) => field.key));
  const userDetailsFieldKeys = new Set(userDetailsFields.map((field) => field.key));
  const addressFieldKeys = new Set(addressGroups.flatMap((group) => group.fields.map((field) => field.key)));
  const addressInUserDetails = hasUserDetails && addressGroups.length > 0;
  const columnConfiguration: Partial<Record<string, SuperDataGridColumnConfiguration>> = {
    rowNumber: { headerName: rowNumberLabel, width: 52, minWidth: 52, maxWidth: 58 },
    status: { headerName: statusLabel, minWidth: 100, flex: 0.75 },
  };
  const columnOptions: Partial<Record<string, SuperDataGridColumnOptions>> = {};
  const columnTypes: Partial<Record<string, SuperDataGridColumnType>> = {};
  const urlFields: string[] = [];
  const dateFields: string[] = [];

  if (hasUserDetails) {
    columnConfiguration[PREVIEW_USER_DETAILS_FIELD] = {
      headerName: "User Details",
      minWidth: 340,
      flex: 2.6,
    };
    columnTypes[PREVIEW_USER_DETAILS_FIELD] = "peopleDetails";
  }
  if (addressGroups.length > 0 && !addressInUserDetails) {
    columnConfiguration[PREVIEW_ADDRESSES_FIELD] = {
      headerName: addressGroups.length > 1 ? "Addresses" : "Address",
      minWidth: 190,
      flex: 1.3,
    };
  }

  const columns = ["rowNumber"];
  let addedUserDetails = false;
  let addedAddresses = false;
  for (const field of visibleFields) {
    if (userDetailsFieldKeys.has(field.key)) {
      if (!addedUserDetails) {
        columns.push(PREVIEW_USER_DETAILS_FIELD);
        addedUserDetails = true;
      }
      continue;
    }

    if (addressFieldKeys.has(field.key)) {
      if (addressInUserDetails) continue;
      if (!addedAddresses) {
        columns.push(PREVIEW_ADDRESSES_FIELD);
        addedAddresses = true;
      }
      continue;
    }

    if (permissionFieldKeys.has(field.key)) continue;

    columns.push(field.key);
    const inferredType = getPreviewFieldType(field);
    const urlField = isUrlField(field);
    const listField = listFieldKeys.has(field.key);
    const minWidth = field.width == null
      ? inferredType === "badge" ? 104 : inferredType === "date" || inferredType === "dateTime" ? 128 : urlField ? 160 : listField ? 80 : field.key.toLowerCase() === "notes" ? 90 : 100
      : Math.min(field.width, 110);
    columnConfiguration[field.key] = {
      headerName: listField && isGroupIdField(field) ? "Groups" : field.label,
      ...(field.width == null ? { flex: inferredType === "badge" || listField ? 0.65 : inferredType === "date" ? 0.8 : field.key.toLowerCase() === "notes" ? 0.7 : 0.85 } : { width: field.width }),
      minWidth,
    };

    if (listField) columnTypes[field.key] = "countChip";
    else if (inferredType) columnTypes[field.key] = inferredType;
    if (inferredType === "date") dateFields.push(field.key);
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
          format: "national",
          ...field.columnOptions?.phone,
        },
      };
    } else if (inferredType === "date") {
      columnOptions[field.key] = {
        ...field.columnOptions,
        date: { showIcon: false, timeZone: "UTC", ...field.columnOptions?.date },
      };
    } else if (inferredType === "badge") {
      columnOptions[field.key] = {
        ...field.columnOptions,
        badge: {
          fallbackColor: "neutral",
          ...field.columnOptions?.badge,
          colors: {
            SUPERADMIN: "purple",
            ADMIN: "purple",
            MANAGER: "cyan",
            VIEWER: "slate",
            CUSTOMER: "slate",
            CUSTOM: "orange",
            ...field.columnOptions?.badge?.colors,
          },
        },
      };
    } else if (field.columnOptions) {
      columnOptions[field.key] = field.columnOptions;
    }
    if (urlField) urlFields.push(field.key);
  }
  columnConfiguration[PREVIEW_RAW_DATA_FIELD] = {
    headerName: "All spreadsheet data",
    minWidth: 220,
    flex: 1.4,
  };
  columns.push(PREVIEW_RAW_DATA_FIELD);
  columns.push("status");

  return {
    columns,
    columnConfiguration,
    columnOptions,
    columnTypes,
    hasUserDetails,
    userDetailsFields,
    addressGroups,
    urlFields,
    dateFields,
    listFields,
  };
};

export const SheetIngest = <Row extends SheetIngestRow = SheetIngestRow>({
  isOpen,
  onClose,
  template,
  onSubmit,
  mapRow,
  validateRow,
  previewCellRenderers,
  referenceDataStep = false,
  onGenerateTestData,
  showTestDataGenerator = true,
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
  const steps = useMemo<WizardStep[]>(
    () => referenceDataStep ? ["reference", ...UPLOAD_STEPS] : UPLOAD_STEPS,
    [referenceDataStep],
  );
  const fileInputRef = useRef<HTMLInputElement>(null);
  const [step, setStep] = useState<WizardStep>(referenceDataStep ? "reference" : "upload");
  const stepHelp = t[STEP_HELP_KEYS[step]];
  const [file, setFile] = useState<File | null>(null);
  const [sheetName, setSheetName] = useState("");
  const [matrix, setMatrix] = useState<string[][]>([]);
  const [headerRow, setHeaderRow] = useState(0);
  const [headerRows, setHeaderRows] = useState<number[]>([0]);
  const [columnMapping, setColumnMapping] = useState<Record<string, number | null>>({});
  const [processedRows, setProcessedRows] = useState<ProcessedRow<Row>[]>([]);
  const [previewPaginationModel, setPreviewPaginationModel] = useState<GridPaginationModel>(() => ({
    page: 0,
    pageSize: IMPORT_GRID_DEFAULT_PAGE_SIZE,
  }));
  const [testDataCount, setTestDataCount] = useState(defaultTestDataCount);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);
  const [generatedTestFile, setGeneratedTestFile] = useState<File | null>(null);

  const reset = useCallback((): void => {
    setStep(referenceDataStep ? "reference" : "upload");
    setFile(null);
    setSheetName("");
    setMatrix([]);
    setHeaderRow(0);
    setHeaderRows([0]);
    setColumnMapping({});
    setProcessedRows([]);
    setPreviewPaginationModel({ page: 0, pageSize: IMPORT_GRID_DEFAULT_PAGE_SIZE });
    setBusy(false);
    setError(null);
    setSuccess(null);
    setGeneratedTestFile(null);
  }, [referenceDataStep]);

  useEffect(() => {
    if (isOpen) reset();
  }, [isOpen, reset]);

  const columns = useMemo(() => getSourceColumns(matrix, headerRows), [headerRows, matrix]);
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
      __sheetIngestSourceValues: item.sourceValues,
    })),
    [processedRows, t.invalid, t.valid],
  );
  const grid = useMemo(
    () => getCustomColumns(fields, t.row, t.status),
    [fields, t.row, t.status],
  );
  const cellComponents = useMemo(() => {
    const components: Partial<Record<string, SuperDataGridCellComponent<GridRow<Row>>>> = {
      status: ErrorStatusCell<Row>,
    };
    if (grid.hasUserDetails) {
      components[PREVIEW_USER_DETAILS_FIELD] = PreviewUserDetailsCell<Row>;
    }
    if (grid.addressGroups.length > 0) {
      components[PREVIEW_ADDRESSES_FIELD] = PreviewAddressesCell<Row>;
    }
    for (const field of grid.urlFields) {
      components[field] = PreviewUrlCell<Row>;
    }
    for (const field of grid.dateFields) {
      components[field] = PreviewDateCell<Row>;
    }
    for (const field of grid.listFields) {
      components[field.key] = PreviewListCountCell<Row>;
    }
    components[PREVIEW_RAW_DATA_FIELD] = PreviewRawDataCell<Row>;
    for (const field of Object.keys(previewCellRenderers ?? {})) {
      const renderer = previewCellRenderers?.[field];
      if (!renderer) continue;
      components[field] = (props) => renderer({
        field: props.field,
        value: props.value,
        row: props.row.__sheetIngestData,
      });
    }
    return components;
  }, [grid, previewCellRenderers]);

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
      const detectedColumns = getSourceColumns(sheet.rows, initialHeaderRows);
      if (detectedColumns.length === 0) {
        throw new Error("No spreadsheet columns could be detected from the header row.");
      }
      const detectedHeaders = detectedColumns.map(
        (column) => column.header || `Column ${column.index + 1}`,
      );
      setFile(selectedFile);
      setSheetName(sheet.name);
      setMatrix(sheet.rows);
      setHeaderRow(detectedHeader);
      setHeaderRows(initialHeaderRows);
      setColumnMapping(suggestColumnMappings(detectedHeaders, fields));
      setStep("mapping");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "The file could not be read.");
    } finally {
      setBusy(false);
    }
  }, [fields, maxFileSize]);

  const handleFileInput = useCallback((event: React.ChangeEvent<HTMLInputElement>): void => {
    const selectedFile = event.currentTarget.files?.[0];
    if (selectedFile) {
      setGeneratedTestFile(null);
      void handleFile(selectedFile);
    }
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
    setGeneratedTestFile(null);
    try {
      const rows = await onGenerateTestData(testDataCount);
      if (rows) {
        if (rows.length === 0) throw new Error("The test-data generator returned no records.");
        const blob = await createTemplateBlob(template, rows);
        const preparedFile = new File([blob], template.fileName, { type: blob.type });
        setGeneratedTestFile(preparedFile);
        setSuccess(`${rows.length.toLocaleString()} test ${t.records} are ready. Download and preview them to continue.`);
      } else {
        setSuccess(`Generated ${testDataCount.toLocaleString()} test ${t.records}.`);
      }
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "Test data could not be generated.");
    } finally {
      setBusy(false);
    }
  }, [maxTestDataCount, onGenerateTestData, t.records, template, testDataCount]);

  const handleDownloadAndPreviewGeneratedFile = useCallback((): void => {
    if (!generatedTestFile) return;
    try {
      downloadBlob(generatedTestFile, generatedTestFile.name);
      void handleFile(generatedTestFile);
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "The generated test-data file could not be downloaded.");
    }
  }, [generatedTestFile, handleFile]);

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
            sourceColumns: columns,
            file,
            sheetName,
            mapRow,
            validateRow,
          }),
        ),
      );
      setProcessedRows(result);
      setPreviewPaginationModel({ page: 0, pageSize: IMPORT_GRID_DEFAULT_PAGE_SIZE });
      setStep("preview");
    } catch (cause) {
      setError(cause instanceof Error ? cause.message : "The file could not be prepared for review.");
    } finally {
      setBusy(false);
    }
  }, [columnMapping, columns, fields, file, headerRow, mapRow, maxRecords, matrix, records, sheetName, validateRow]);

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
      const currentIndex = steps.indexOf(current);
      return steps[Math.max(currentIndex - 1, 0)] ?? (referenceDataStep ? "reference" : "upload");
    });
  }, [referenceDataStep, steps]);

  const goNext = useCallback((): void => {
    setError(null);
    if (step === "reference") setStep("upload");
    else if (step === "mapping") void handleBuildPreview();
  }, [handleBuildPreview, step]);

  const handleDrop = useCallback((event: React.DragEvent<HTMLDivElement>): void => {
    event.preventDefault();
    const droppedFile = event.dataTransfer.files[0];
    if (droppedFile) void handleFile(droppedFile);
  }, [handleFile]);

  const activeStepIndex = steps.indexOf(step);
  const stepLabels: Record<WizardStep, string> = {
    reference: t.steps.reference ?? "Reference data",
    upload: t.steps.upload,
    mapping: t.steps.mapping,
    preview: t.steps.preview,
  };
  const mappingComplete = fields.every((field) => {
    const mapping = columnMapping[field.key];
    return !field.required || (mapping != null && mapping >= 0);
  });
  const componentRows = rowCount;
  const previewRows = useMemo(
    () => gridRows.map((row) => {
      const rowData = row.__sheetIngestData;
      const formattedFields = Object.fromEntries(
        fields.map((field) => [field.key, cellText(Reflect.get(rowData, field.key))]),
      );
      const userDetails: Record<string, unknown> = Object.fromEntries(
        grid.userDetailsFields.map((field) => [field.key, cellText(Reflect.get(rowData, field.key))]),
      );
      const userDetailValue = (...names: string[]): string => {
        const field = grid.userDetailsFields.find((candidate) =>
          getFieldNames(candidate).some((name) => names.includes(name)),
        );
        return field ? getRecordText(rowData, field.key) : "";
      };
      const userName = [userDetailValue("firstname"), userDetailValue("lastname")]
        .filter(Boolean)
        .join(" ") || userDetailValue("name", "fullname");
      const userEmail = userDetailValue("loginname", "email");
      const userPhone = userDetailValue("phone");
      const isSameContact = (addressValue: string, userValue: string, phone = false): boolean => {
        if (!addressValue || !userValue) return false;
        const normalize = (value: string): string => phone
          ? value.replace(/\D/g, "")
          : value.trim().toLocaleLowerCase();
        return normalize(addressValue) === normalize(userValue);
      };
      const addresses = grid.addressGroups.flatMap((group) => {
        const values = new Map<string, string>();
        for (const field of group.fields) {
          const descriptor = getAddressFieldGroup(field);
          if (!descriptor) continue;
          const text = getRecordText(rowData, field.key);
          if (text) values.set(descriptor.part, text);
        }
        const street = [values.get("streetaddress"), values.get("streetaddress2"), values.get("streetaddress3")]
          .filter(Boolean)
          .join(", ");
        const locality = [
          values.get("city"),
          values.get("state"),
          values.get("zipcode") || values.get("postalcode"),
          values.get("country"),
        ].filter(Boolean).join(", ");
        const addressLine = [street, locality].filter(Boolean).join(", ");
        const summary = [
          values.get("city"),
          values.get("state"),
          values.get("zipcode") || values.get("postalcode"),
        ].filter(Boolean).join(", ") || street || "Open address details";
        const contact = [
          values.get("nameonaddress") && !isSameContact(values.get("nameonaddress") ?? "", userName)
            ? `Name: ${values.get("nameonaddress")}`
            : "",
          values.get("emailonaddress") && !isSameContact(values.get("emailonaddress") ?? "", userEmail)
            ? `Email: ${values.get("emailonaddress")}`
            : "",
          values.get("phoneonaddress") && !isSameContact(values.get("phoneonaddress") ?? "", userPhone, true)
            ? `Phone: ${values.get("phoneonaddress")}`
            : "",
        ].filter(Boolean).join(" · ");
        const lines = [addressLine].filter(Boolean);
        if (lines.length === 0 && !values.get("addresstype")) return [];
        return [{
          label: group.number === 1 ? "Primary" : `Address ${group.number}`,
          addressType: values.get("addresstype") ?? "",
          summary,
          lines,
          contact,
        }];
      });

      return {
        ...row,
        ...formattedFields,
        ...(grid.hasUserDetails ? { [PREVIEW_USER_DETAILS_FIELD]: { ...userDetails, addresses } } : {}),
        ...(grid.addressGroups.length > 0 ? { [PREVIEW_ADDRESSES_FIELD]: addresses } : {}),
        [PREVIEW_RAW_DATA_FIELD]: row.__sheetIngestSourceValues,
      };
    }),
    [fields, grid, gridRows],
  );

  return (
    <OpsModalDialog
      {...dialogProps}
      open={isOpen}
      onClose={handleClose}
      fullWidth={dialogProps?.fullWidth ?? true}
      maxWidth={dialogProps?.maxWidth ?? "sm"}
      title={title ?? t.title}
      closeButtonLabel={t.close}
      classNames={{
        ...dialogProps?.classNames,
        paper: `${styles.previewDialog} ${dialogProps?.classNames?.paper ?? ""}`.trim(),
        content: `${styles.dialogContent} ${step === "upload" ? styles.uploadContent : ""} ${dialogProps?.classNames?.content ?? ""}`,
        actions: `${styles.dialogActions} ${dialogProps?.classNames?.actions ?? ""}`,
      }}
      actions={(
        <>
          <Stack direction="row" spacing={1}>
            {activeStepIndex > 0 && (
              <Button className={styles.navigationButton} startIcon={<NavigateBeforeIcon />} onClick={goBack} disabled={busy}>
                {t.back}
              </Button>
            )}
            {(step === "reference" || step === "mapping") && (
              <Button className={styles.primaryButton} variant="contained" endIcon={<NavigateNextIcon />} onClick={goNext} disabled={busy || (step === "mapping" && records.length === 0)}>
                {busy ? "Preparing…" : step === "reference" ? "Continue to upload" : t.next}
              </Button>
            )}
            {step === "preview" && (
              <Button className={styles.primaryButton} variant="contained" onClick={() => void handleSubmit()} disabled={busy || rowCount === 0 || (!allowInvalidSubmit && errorRowCount > 0)}>
                {busy ? "Importing…" : `${t.submit} (${rowCount})`}
              </Button>
            )}
          </Stack>
        </>
      )}
    >
      <Box className={styles.stepperBar}>
        <Stepper className={styles.stepper} activeStep={activeStepIndex}>
          {steps.map((item) => (
            <Step className={styles.step} key={item}>
              <StepLabel className={styles.stepLabel}>{stepLabels[item]}</StepLabel>
            </Step>
          ))}
        </Stepper>
      </Box>
      <Box className={styles.subtitleBanner}>
        {stepHelp}
      </Box>
        {busy && <LinearProgress className={styles.progress} sx={{ mb: 2 }} />}
        {error && <Alert className={styles.alert} severity="error" sx={{ mb: 2 }}>{error}</Alert>}
        {success && <Alert className={styles.alert} severity="success" sx={{ mb: 2 }}>{success}</Alert>}

        {step === "upload" && (
          <Stack className={styles.uploadStep}>
            <Box
              className={styles.dropzone}
              onClick={(event) => {
                if ((event.target as HTMLElement).closest("button")) return;
                fileInputRef.current?.click();
              }}
              onDragOver={(event) => event.preventDefault()}
              onDrop={handleDrop}
            >
              <Stack spacing={1.5} alignItems="center">
                {file ? (
                  <FilePresentIcon className={styles.uploadIcon} />
                ) : (
                  <CloudUploadIcon className={styles.uploadIcon} />
                )}
                <Typography className={styles.dropzoneHeadline} variant="subtitle1" fontWeight={600}>
                  {file?.name ?? t.uploadFile}
                </Typography>
                {file && <Typography variant="body2" color="text.secondary">{(file.size / 1024).toFixed(1)} KB</Typography>}
                <Button
                  variant="contained"
                  className={styles.uploadButton}
                  startIcon={<CloudUploadIcon />}
                  onClick={(event) => {
                    event.stopPropagation();
                    fileInputRef.current?.click();
                  }}
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
                <Typography className={styles.dropzoneHelper} variant="caption">or drop a file here</Typography>
              </Stack>
            </Box>
            <Stack className={styles.utilityBar} direction={{ xs: "column", sm: "row" }} spacing={1.5} alignItems={{ xs: "stretch", sm: "center" }}>
              <Button className={styles.downloadButton} variant="outlined" startIcon={<DownloadIcon />} onClick={() => void handleDownloadTemplate()}>
                {t.downloadTemplate}
              </Button>
              {showTestDataGenerator && onGenerateTestData && (
                <>
                  <Box className={styles.recordCountControl}>
                    <Typography
                      className={styles.recordCountLabel}
                      component="label"
                      htmlFor="sheet-ingest-test-record-count"
                    >
                      {t.testRecordCount}
                    </Typography>
                    <TextField
                      id="sheet-ingest-test-record-count"
                      size="small"
                      type="number"
                      className={styles.recordCountField}
                      value={testDataCount}
                      onChange={(event) => {
                        setGeneratedTestFile(null);
                        setTestDataCount(Number(event.target.value));
                      }}
                      inputProps={{
                        min: 1,
                        max: maxTestDataCount,
                        step: 1,
                        "aria-label": t.testRecordCount,
                      }}
                    />
                  </Box>
                  <Button
                    className={styles.generateButton}
                    variant="text"
                    startIcon={<AutoAwesomeIcon />}
                    onClick={() => void handleGenerateTestData()}
                    disabled={busy}
                  >
                    {t.generateTestData}
                  </Button>
                  {generatedTestFile && (
                    <Button
                      className={styles.generatedDataButton}
                      variant="contained"
                      startIcon={<DownloadIcon />}
                      onClick={handleDownloadAndPreviewGeneratedFile}
                      disabled={busy}
                    >
                      Download &amp; preview {testDataCount} records
                    </Button>
                  )}
                </>
              )}
            </Stack>
          </Stack>
        )}

        {step === "mapping" && (
          <Stack spacing={2}>
            <Stack direction={{ xs: "column", sm: "row" }} spacing={1.5} justifyContent="space-between">
              <Box>
                <Typography className={styles.mappingTitle} variant="subtitle1" fontWeight={600}>{t.matchColumns}</Typography>
              </Box>
              <Box sx={{ textAlign: { sm: "right" } }}>
                <Typography variant="body2" fontWeight={600}>{file?.name}</Typography>
                <Typography variant="caption" color="text.secondary">
                  {sheetName} · {records.length.toLocaleString()} {records.length === 1 ? "row" : "rows"} · {columns.length.toLocaleString()} columns
                </Typography>
              </Box>
            </Stack>
            <Stack spacing={2}>
              {template.groups.map((group) => (
                <Paper className={styles.groupCard} key={group.label} variant="outlined" sx={{ overflow: "hidden" }}>
                  <Box className={styles.groupHeader} sx={{ px: 2, py: 1, bgcolor: "action.hover" }}>
                    <Typography className={styles.groupTitle} variant="subtitle2" fontWeight={700}>{group.label}</Typography>
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
                          className={styles.mappingRow}
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
                              className={styles.selectControl}
                              labelId={`sheet-ingest-map-${field.key}`}
                              label={t.matchColumns}
                              value={sourceIndex ?? ""}
                              MenuProps={selectMenuProps}
                              onChange={(event) => {
                                const value = event.target.value === "" ? null : Number(event.target.value);
                                handleMappingChange(field.key, value);
                              }}
                            >
                              <MenuItem className={styles.menuItem} value="">
                                <em>{t.noMapping}</em>
                                {sourceIndex == null && <CheckIcon className={styles.checkIcon} />}
                              </MenuItem>
                              {columns.map((column) => (
                                <MenuItem
                                  className={styles.menuItem}
                                  key={column.index}
                                  value={column.index}
                                  disabled={unavailableIndices.has(column.index)}
                                >
                                  {column.group ? `${column.group} / ` : ""}{column.header || `Column ${column.index + 1}`}
                                  {sourceIndex === column.index && <CheckIcon className={styles.checkIcon} />}
                                </MenuItem>
                              ))}
                              <Divider />
                              <MenuItem className={styles.menuItem} value={-1}>
                                {t.ignoreColumn}
                                {sourceIndex === -1 && <CheckIcon className={styles.checkIcon} />}
                              </MenuItem>
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
                <Typography className={styles.mappingTitle} variant="subtitle1" fontWeight={600}>{t.dataPreview}</Typography>
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
            <Box className={styles.previewGrid}>
              <SuperDataGrid
                columns={grid.columns}
                data={previewRows}
                columnConfiguration={grid.columnConfiguration}
                columnOptions={grid.columnOptions}
                columnTypes={grid.columnTypes}
                cellComponents={cellComponents}
                paginationModel={previewPaginationModel}
                onPaginationModelChange={setPreviewPaginationModel}
                gridClassName={styles.previewGridRoot}
                getRowId={(row) => row.rowNumber}
                height="100%"
                minHeight="100%"
                getRowHeight={() => "auto"}
                getEstimatedRowHeight={() => 88}
                pageSizeOptions={IMPORT_GRID_PAGE_SIZE_OPTIONS}
                hideToolbar
                hideViews
                canAddViews={false}
              />
            </Box>
          </Stack>
        )}

        {renderStepContent?.(contextForStep)}
    </OpsModalDialog>
  );
};

export default SheetIngest;
