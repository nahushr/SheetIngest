<p align="center">
  <img src="https://raw.githubusercontent.com/nahushr/SheetIngest/main/assets/sheet-ingest-logo.svg" alt="SheetIngest" width="600" />
</p>

<p align="center">
  <a href="https://github.com/nahushr/SheetIngest/actions/workflows/deploy.yml"><img alt="CI" src="https://github.com/nahushr/SheetIngest/actions/workflows/deploy.yml/badge.svg?branch=main" /></a>
  <a href="https://www.npmjs.com/package/@simplishelf/sheet-ingest"><img alt="npm version" src="https://img.shields.io/npm/v/@simplishelf/sheet-ingest?logo=npm" /></a>
  <img alt="React 18+" src="https://img.shields.io/badge/React-18%2B-61DAFB?logo=react&logoColor=111827" />
  <img alt="TypeScript" src="https://img.shields.io/badge/TypeScript-types%20included-3178C6?logo=typescript&logoColor=white" />
  <a href="LICENSE"><img alt="MIT license" src="https://img.shields.io/badge/license-MIT-16a085.svg" /></a>
</p>

<p align="center">
  <a href="https://stackblitz.com/github/nahushr/SheetIngest/tree/main/examples/vite?file=src/App.tsx"><img alt="Open the SheetIngest example in StackBlitz" src="https://developer.stackblitz.com/img/open_in_stackblitz.svg" /></a>
</p>

SheetIngest is a configurable React wizard for importing CSV and Excel files. Define grouped fields and supply your own mapping, validation, test-data, and submit callbacks. The package handles file upload, automatic header detection, column matching, and data preview.

## Demo

| Online | Local |
|---|---|
| [Open the Vite example in StackBlitz](https://stackblitz.com/github/nahushr/SheetIngest/tree/main/examples/vite?file=src/App.tsx) | `npm ci` → `npm run dev:example` → [localhost:7013](http://localhost:7013) |

| Example | What to try |
|---|---|
| Contact import | Open the three-step import dialog from the page toolbar. |
| Grouped template | Download the sample workbook and inspect its merged Contact and Organization headings. |
| Test data | Choose a row count, generate realistic contact records, and upload the resulting workbook. |
| Mapping and validation | Change headers in the workbook, map them in the wizard, and review row-level email validation. |

## Install

```sh
npm install @simplishelf/sheet-ingest @simplishelf/super-data-grid
```

The package uses React 18+, Material UI, Emotion, and SuperDataGrid as peer dependencies. Install compatible versions in the host app. Import the package stylesheet once from the app entry point:

```tsx
import "@simplishelf/sheet-ingest/style.css";
```

## Quick start

Describe the fields your application accepts. `groups` control the merged category row in the downloaded workbook and keep related columns together in the mapper.

```tsx
import { useState } from "react";
import { SheetIngest, downloadTemplate } from "@simplishelf/sheet-ingest";
import type {
  SheetIngestTemplate,
  SheetIngestIssue,
} from "@simplishelf/sheet-ingest";
import "@simplishelf/sheet-ingest/style.css";

type Contact = {
  firstName: string;
  lastName: string;
  email: string;
  company: string;
};

const contactTemplate: SheetIngestTemplate = {
  fileName: "contacts_template.xlsx",
  sheetName: "Contacts",
  groups: [
    {
      label: "Contact",
      fields: [
        { key: "firstName", label: "First name", aliases: ["Given name"], required: true },
        { key: "lastName", label: "Last name", aliases: ["Surname"], required: true },
        { key: "email", label: "Email address", aliases: ["Email"], required: true },
      ],
    },
    {
      label: "Organization",
      fields: [
        { key: "company", label: "Company", aliases: ["Organization"] },
      ],
    },
  ],
};

function makeContacts(count: number): Contact[] {
  return Array.from({ length: count }, (_, index) => ({
    firstName: `Contact${index + 1}`,
    lastName: "Example",
    email: `contact${index + 1}@example.com`,
    company: "Northstar Studio",
  }));
}

async function saveContacts(contacts: readonly Contact[], file: File): Promise<void> {
  console.info("Send these contacts to your API", {
    fileName: file.name,
    contacts,
  });
}

export function ContactsPage() {
  const [isOpen, setIsOpen] = useState(false);

  return (
    <>
      <button onClick={() => setIsOpen(true)}>Import contacts</button>
      <SheetIngest<Contact>
        isOpen={isOpen}
        onClose={() => setIsOpen(false)}
        template={contactTemplate}
        mapRow={(values) => ({
          firstName: values.firstName.trim(),
          lastName: values.lastName.trim(),
          email: values.email.trim().toLowerCase(),
          company: values.company.trim(),
        })}
        validateRow={(contact): SheetIngestIssue | undefined =>
          /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(contact.email)
            ? undefined
            : { field: "email", message: "Enter a valid email address." }
        }
        onGenerateTestData={async (count) => {
          await downloadTemplate(contactTemplate, makeContacts(count));
        }}
        onSubmit={async (contacts, file) => {
          await saveContacts(contacts, file);
        }}
      />
    </>
  );
}
```

Replace the example `saveContacts` handler with the host application's API call. The example project contains a complete runnable page, including the generated workbook callback.

## How the import flow works

| Step | Behavior |
|---|---|
| Upload file | Accept a CSV or XLSX file, download the grouped sample template, or generate host-provided test data. The header rows are inferred automatically from template field names and aliases. |
| Match columns | Show the uploaded file and worksheet, suggest one-to-one field matches, and allow each source column to be remapped or ignored. Nested group labels from the workbook remain visible. |
| Data preview | Show mapped rows and row-level validation issues in SuperDataGrid. Email and phone fields use linked cells with icons, country-specific phone formatting, and country flags. |

The first workbook row can contain merged group labels; the next row contains field headers. SheetIngest detects the field-header row and carries its preceding group row into the column mappings. `templateHeader` changes the header written to the workbook, while `aliases` add names that can be matched automatically. Required fields receive a built-in missing-value check. `mapRow` converts source strings into the host's model; `validateRow` applies entity-specific rules. Errors block submission by default, while warning and informational issues remain visible without blocking it.

`onGenerateTestData(count)` runs in the host application. Generate records using your own domain data or a test-data library, then call `downloadTemplate(template, rows)` to return an XLSX file using the same grouped template. Rows should be keyed by the template field keys.

## Component options

| Prop | Description |
|---|---|
| `isOpen`, `onClose` | Control the dialog from the host page. |
| `template` | File name, sheet name, grouped fields, aliases, required flags, and preview configuration. |
| `mapRow` | Convert mapped spreadsheet strings into the application's row type. |
| `validateRow` | Return a message, issue object, array of issues, or `undefined` for each row. |
| `onSubmit` | Receive validated rows and the original uploaded `File`. |
| `onGenerateTestData` | Generate and download host-specific sample rows for the requested count. |
| `showTestDataGenerator` | Show or hide the count input and generate button when a callback is supplied. Defaults to `true`. |
| `maxRecords`, `maxFileSize` | Set parsing limits. Defaults are 1,000 rows and 10 MiB. |
| `defaultTestDataCount`, `maxTestDataCount` | Set the test-data count input defaults. Defaults are 50 and 10,000. |
| `allowInvalidSubmit` | Allow submission when rows contain validation errors. Defaults to `false`. |
| `title`, `translations`, `dialogProps` | Customize the dialog title, labels, and MUI dialog props. |
| `renderStepContent` | Render host-specific content below the active step. |

Email and phone preview types are inferred from each field's key, label, template header, and aliases. Set `type: "email"` or `type: "phone"` to override inference. Phone fields default to country-specific national formatting, including the familiar parentheses for US numbers. Set `columnOptions.phone.format` to `"international"` or `"original"`, and `columnOptions.phone.countryCode` to choose a different default country. Email links use `mailto:`, and phone links use `tel:`.

The exported `SheetIngestProps`, `SheetIngestTemplate`, `SheetIngestGroup`, `SheetIngestField`, `SheetIngestIssue`, and row-context types are available for TypeScript apps.

## Development

```sh
npm ci
npm run lint
npm test
npm run typecheck
npm run build
npm run build:example
npm run dev:example
npm pack --dry-run
```

The example runs at [http://localhost:7013](http://localhost:7013). The GitHub Actions workflow builds the package and example, runs lint, tests, type checks, SonarCloud, and Snyk, then publishes public releases from `main`. Sonar fails on any unresolved issue; coverage is excluded. Configure `SONAR_TOKEN`, `SNYK_TOKEN`, and `NPM_TOKEN` as repository Actions secrets.

## License

MIT. See [LICENSE](LICENSE).
