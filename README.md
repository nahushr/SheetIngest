# SheetIngest

SheetIngest is a configurable React import wizard for CSV and XLSX files. It provides a four-step flow for uploading a file, choosing its header row, matching grouped template fields, and reviewing validated rows. Import and test-data callbacks stay in the host application, so the package can be used with different entity models and APIs.

## Install

```sh
npm install @simplishelf/sheet-ingest @simplishelf/super-data-grid
```

React, React DOM, Material UI, Emotion, and SuperDataGrid are peer dependencies. The host app should use compatible versions of those packages.

## Basic usage

```tsx
import { useState } from "react";
import { SheetIngest } from "@simplishelf/sheet-ingest";
import { downloadTemplate } from "@simplishelf/sheet-ingest";
import "@simplishelf/sheet-ingest/style.css";

const template = {
  fileName: "contacts_template.xlsx",
  sheetName: "Contacts",
  groups: [
    {
      label: "Contact details",
      fields: [
        {
          key: "firstName",
          label: "First name",
          aliases: ["Given name"],
          required: true,
        },
        {
          key: "email",
          label: "Email address",
          aliases: ["Email"],
          required: true,
          type: "email",
        },
      ],
    },
  ],
} as const;

function ContactsPage() {
  const [isImportOpen, setIsImportOpen] = useState(false);

  return (
    <>
      <button onClick={() => setIsImportOpen(true)}>Import contacts</button>
      <SheetIngest
        isOpen={isImportOpen}
        onClose={() => setIsImportOpen(false)}
        template={template}
        mapRow={(values) => ({
          firstName: values.firstName.trim(),
          email: values.email.trim().toLowerCase(),
        })}
        validateRow={(row) =>
          row.email.includes("@") ? undefined : "Enter a valid email address."
        }
        onGenerateTestData={async (count) => {
          const generatedRows = await generateContacts(count);
          await downloadTemplate(template, generatedRows);
        }}
        onSubmit={async (rows) => {
          await createContacts(rows);
        }}
      />
    </>
  );
}
```

The generated sample workbook has one merged category row followed by a field row. `downloadTemplate(template, rows)` can also write host-generated records beneath those headers; each record is keyed by the template field keys. `templateHeader` can override the field key written into that workbook. Header matching considers each field's key, label, template header, and aliases. A field marked `required` receives a built-in missing-value check; `validateRow` can add host-specific rules. Validation callbacks may return a string, an issue object, an array of either, or `undefined`.

`onSubmit` receives the mapped rows and original `File`. Submission is blocked when any row has an error unless `allowInvalidSubmit` is enabled. Warning and informational issues are shown in the preview but do not block submission.

The package accepts `maxRecords`, `maxFileSize`, `defaultTestDataCount`, `maxTestDataCount`, `title`, `translations`, `dialogProps`, and `renderStepContent` for host-level configuration. See the exported `SheetIngestProps`, `SheetIngestTemplate`, and `SheetIngestField` TypeScript types for details.

## Development

```sh
npm install
npm run lint
npm test
npm run typecheck
npm run build
npm pack --dry-run
```

The GitHub Actions workflow builds and verifies the package, runs SonarCloud and Snyk analysis, fails when SonarCloud reports any unresolved issue, and publishes public releases to npm from `main`. Add `SONAR_TOKEN`, `SNYK_TOKEN`, and `NPM_TOKEN` as repository Actions secrets. Sonar coverage is excluded for the project.
