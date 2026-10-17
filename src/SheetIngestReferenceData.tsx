import type { ReactNode } from "react";
import { ReferenceDataCard } from "@simplishelf/opscards";
import SuperDataGrid from "@simplishelf/super-data-grid";
import type {
  SuperDataGridColumnOptions,
  SuperDataGridColumnConfiguration,
  SuperDataGridColumnType,
  SuperDataGridFilterField,
} from "@simplishelf/super-data-grid";

import { IMPORT_GRID_PAGE_SIZE_OPTIONS } from "./pagination";
import styles from "./SheetIngestReferenceData.module.css";

export interface SheetIngestReferenceGrid {
  /** Stable key for the section. */
  id: string;
  /** Section title shown above its grid. */
  title: string;
  /** Short guidance shown below the title. */
  description: string;
  /** Grid rows supplied by the host application. */
  rows: readonly object[];
  /** Row fields, ordered from left to right. */
  columns: readonly string[];
  /** The row field that uniquely identifies each record. */
  rowIdField: string;
  /** Human-readable labels for grid fields. */
  columnLabels: Readonly<Record<string, string>>;
  /** Optional package-owned cell renderers for standard grid presentation. */
  columnTypes?: Partial<Record<string, SuperDataGridColumnType>>;
  /** Options for package-owned cell renderers, keyed by column field. */
  columnOptions?: Partial<Record<string, SuperDataGridColumnOptions>>;
  /** Optional width and layout overrides for the reference grid. */
  columnConfiguration?: Partial<Record<string, SuperDataGridColumnConfiguration>>;
  /** Typed fields used by the grid's filter builder. */
  filterFields: readonly SuperDataGridFilterField[];
  /** Optional empty-state message. */
  emptyMessage?: string;
}

export interface SheetIngestReferenceDataProps {
  /** Heading shown above the reference sections. */
  title: string;
  /** Instructions explaining how the reference data should be used. */
  description: string;
  /** Read-only grid sections. */
  sections: readonly SheetIngestReferenceGrid[];
  /** Show loading overlays while reference data is being fetched. */
  loading?: boolean;
  /** Optional error text. */
  error?: ReactNode;
  /** Optional retry/refresh action. */
  onRefresh?: () => void;
}

const createColumnConfiguration = (
  section: SheetIngestReferenceGrid,
): Partial<Record<string, SuperDataGridColumnConfiguration>> =>
  Object.fromEntries(section.columns.map((field) => {
    const values = section.rows.map((row) => String(Reflect.get(row, field) ?? ""));
    const longestValue = Math.max(
      section.columnLabels[field]?.length ?? field.length,
      ...values.map((value) => value.length),
    );
    const minWidth = Math.max(116, Math.min(320, longestValue * 7 + 36));
    return [field, {
      ...section.columnConfiguration?.[field],
      headerName: section.columnLabels[field] ?? field,
      minWidth: section.columnConfiguration?.[field]?.minWidth ?? minWidth,
      flex: section.columnConfiguration?.[field]?.flex ?? Math.max(0.7, minWidth / 180),
    }];
  }));

const createEmptyOverlay = (message: string) => function ReferenceGridEmptyOverlay(): React.JSX.Element {
  return <div className={styles.emptyGrid}>{message}</div>;
};

export const SheetIngestReferenceData = ({
  title,
  description,
  sections,
  loading = false,
  error,
  onRefresh,
}: SheetIngestReferenceDataProps): React.JSX.Element => (
  <section className={styles.referenceStep} aria-label={title}>
    <header className={styles.referenceHeader}>
      <div>
        <h2>{title}</h2>
        <p>{description}</p>
      </div>
      {onRefresh && (
        <button className={styles.refreshButton} type="button" onClick={onRefresh} disabled={loading}>
          {loading ? "Loading…" : "Refresh lists"}
        </button>
      )}
    </header>

    {error && (
      <div className={styles.errorNotice} role="alert">
        <span>{error}</span>
        {onRefresh && <button type="button" onClick={onRefresh} disabled={loading}>Retry</button>}
      </div>
    )}

    <div className={styles.referenceSections}>
      {sections.map((section, index) => (
        <ReferenceDataCard
          key={section.id}
          ordinal={index + 1}
          title={section.title}
          description={section.description}
          aria-label={section.title}
        >
          <div className={styles.referenceGridInstance} aria-label={section.title} role="region">
            <SuperDataGrid
              columns={section.columns}
              columnConfiguration={createColumnConfiguration(section)}
              columnTypes={section.columnTypes}
              columnOptions={section.columnOptions}
              data={section.rows}
              filterFields={section.filterFields}
              filterFieldsMode="replace"
              filterMode="client"
              paginationMode="client"
              sortingMode="client"
              checkboxSelection={false}
              isRowSelectable={() => false}
              hideBulkDelete
              hideViews
              canAddViews={false}
              canEditViews={false}
              canDeleteViews={false}
              getRowId={(row) => String(Reflect.get(row, section.rowIdField) ?? "")}
              height={460}
              minHeight={440}
              workspacePadding="none"
              getRowHeight={() => "auto"}
              pageSizeOptions={IMPORT_GRID_PAGE_SIZE_OPTIONS}
              loading={loading}
              dataGridSlots={{ noRowsOverlay: createEmptyOverlay(section.emptyMessage ?? "No reference data available.") }}
            />
          </div>
        </ReferenceDataCard>
      ))}
    </div>
  </section>
);

export default SheetIngestReferenceData;
